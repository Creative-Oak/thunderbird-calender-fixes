"use strict";

// NOTE: Every experiment parent script in this add-on is loaded into the SAME
// shared global, so top-level `const`/`let` names would collide across scripts
// (e.g. MAIN_WINDOW_URL). We wrap the whole body in an IIFE so its declarations
// stay local; `.call(this)` keeps `this` pointing at the shared global so the
// `this.autoInvite = …` API registration still works.
(function () {

// =============================================================================
// Auto-add emailed invitations to the calendar  (WebExtension Experiment)
// =============================================================================
//
// GOAL
// ----
// When an email containing a calendar invitation (iMIP REQUEST) arrives, put a
// copy of the event into the calendar *immediately* — before the user clicks
// Accept/Decline — shown as an un-answered invitation (dotted outline). No RSVP
// email is sent; the user still accepts/declines normally from the message's
// invitation bar, which then updates the same item (matched by UID) in place.
//
// HOW THE PIECES FIT (verified against comm-central tip)
// ------------------------------------------------------
// * New mail: nsIMsgFolderNotificationService ("mfn"). We listen with the
//   `msgsClassified` flag, which fires once per newly-arrived, classified
//   message (mailnews/base/public/nsIMsgFolderListener.idl).
// * Extract the invite: a message's `text/calendar` part is only turned into a
//   calIItipItem by the MIME converter while the message is *displayed*
//   (CalMimeConverter sets channel.imipItem). On arrival there is no such
//   shortcut, so we parse the MIME ourselves with MsgHdrToMimeMessage
//   (resource:///modules/gloda/MimeMessage.sys.mjs) and read the inline
//   text/calendar body.
// * Only METHOD:REQUEST is auto-added (a real invitation). REPLY/CANCEL/COUNTER
//   and plain .ics attachments (no METHOD) are ignored.
// * Recognise the user: cal.itip.getInvitedAttendee(item, calendar) returns the
//   attendee that matches that calendar's own email identity — the exact check
//   Thunderbird uses to decide something is an "invitation" and to apply the
//   dotted `invitation-status` styling. We add to the first calendar that
//   recognises the invite (preferring the default calendar).
// * Add without replying: we call calendar.addItem() directly. Thunderbird's
//   own Accept path wraps this in `ItipOpListener`, whose completion handler
//   calls itip.checkAndSend() to email the reply — we deliberately do NOT use
//   that wrapper, so nothing is sent.
// * De-duplicate: calendar.getItem(uid) is the same UID lookup the invitation
//   bar uses, so a later Accept finds our copy and modifies it instead of
//   creating a second one.
//
// LIMITATION: if no calendar is associated with the address the invite was sent
// to, getInvitedAttendee() returns null for every calendar and we skip the
// message (Thunderbird itself wouldn't consider it an invitation either).

var { ExtensionCommon } = ChromeUtils.importESModule(
  "resource://gre/modules/ExtensionCommon.sys.mjs"
);
var { ExtensionSupport } = ChromeUtils.importESModule(
  "resource:///modules/ExtensionSupport.sys.mjs"
);
var { MailServices } = ChromeUtils.importESModule(
  "resource:///modules/MailServices.sys.mjs"
);
var { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
  "resource:///modules/gloda/MimeMessage.sys.mjs"
);
var { cal } = ChromeUtils.importESModule(
  "resource:///modules/calendar/calUtils.sys.mjs"
);
var { NetUtil } = ChromeUtils.importESModule(
  "resource://gre/modules/NetUtil.sys.mjs"
);

const LOG = "[calendar-tweaks/auto-invite]";
// Flip to true to re-enable the step-by-step diagnostics for troubleshooting.
const DEBUG = false;
function dbg(...args) {
  if (DEBUG) {
    console.log(LOG, ...args);
  }
}
const MAIN_WINDOW_URL = "chrome://messenger/content/messenger.xhtml";
const MAIN_WINDOW_LISTENER_ID = "autoInvite-mainWindowListener";
const STYLE_ELEMENT_ID = "autoInvite-invitation-style";
// Stashed on the calendar command controller so we can restore doCommand.
const DOCMD_MARKER = "_calendarTweaks_originalDoCommand";
// The calendar's Synchronize / "reload remote calendars" command.
const RELOAD_COMMAND = "calendar_reload_remote_calendars";

// Folders whose new messages are never incoming invitations to us.
function skippableFolder(folder) {
  const F = Ci.nsMsgFolderFlags;
  const skip = F.SentMail | F.Drafts | F.Templates | F.Queue | F.Junk | F.Newsgroup;
  return !folder || (folder.flags & skip) != 0;
}

// -----------------------------------------------------------------------------
// New-mail listener
// -----------------------------------------------------------------------------

const folderListener = {
  QueryInterface: ChromeUtils.generateQI(["nsIMsgFolderListener"]),
  msgsClassified(msgs /*, junkProcessed, traitProcessed */) {
    dbg(`msgsClassified fired for ${msgs.length} message(s)`);
    for (const msgHdr of msgs) {
      try {
        handleMessage(msgHdr);
      } catch (error) {
        console.error(LOG, "handleMessage failed:", error);
      }
    }
  },
};

function handleMessage(msgHdr) {
  if (skippableFolder(msgHdr.folder)) {
    dbg("skipping message in folder:", msgHdr.folder?.name);
    return;
  }
  dbg("examining message:", msgHdr.mime2DecodedSubject || msgHdr.subject);
  // On arrival, allow download so IMAP bodies not yet cached are fetched.
  parseMessageForInvite(msgHdr, true);
}

/**
 * Parse one message's MIME and, if it carries an inline text/calendar
 * invitation, process it. Shared by the on-arrival listener and the manual
 * backfill scan.
 *
 * @param {nsIMsgDBHdr} msgHdr
 * @param {boolean} allowDownload - fetch the body if not already offline.
 */
function parseMessageForInvite(msgHdr, allowDownload) {
  MsgHdrToMimeMessage(
    msgHdr,
    null,
    (hdr, mimeMessage) => {
      const subject = hdr.mime2DecodedSubject || hdr.subject;
      (async () => {
        try {
          if (!mimeMessage) {
            dbg("no MIME available (message not downloaded?):", subject);
            return;
          }
          // 1) Inline text/calendar part (Google-style invitations).
          let ics = findCalendarBody(mimeMessage);
          // 2) Otherwise a calendar attachment (invite.ics / application/ics,
          //    as sent by Outlook and many servers).
          if (!ics) {
            ics = await findCalendarInAttachments(mimeMessage);
          }
          if (ics) {
            processInvite(ics);
          } else {
            dbg("no calendar data in:", subject, "| parts:", describeParts(mimeMessage));
          }
        } catch (error) {
          console.error(LOG, "invite processing failed:", error);
        }
      })();
    },
    allowDownload,
    { partsOnDemand: false }
  );
}

/** Does this attachment look like an iCalendar file? */
function isCalendarAttachment(att) {
  const ct = (att.contentType || "").toLowerCase();
  const name = (att.name || att.url || "").toLowerCase();
  return (
    ct == "text/calendar" ||
    ct == "application/ics" ||
    ct.includes("calendar") ||
    name.endsWith(".ics")
  );
}

/** Fetch + decode any calendar attachment and return its ICS text, or null. */
async function findCalendarInAttachments(mimeMessage) {
  let attachments = [];
  try {
    attachments = (mimeMessage.allAttachments || []).filter(isCalendarAttachment);
  } catch (error) {
    return null;
  }
  for (const att of attachments) {
    try {
      const text = await fetchAttachmentText(att.url);
      if (text && text.includes("BEGIN:VCALENDAR")) {
        dbg("found calendar attachment:", att.name || att.contentType);
        return text;
      }
    } catch (error) {
      dbg("attachment fetch failed:", att.name, "-", error?.message || error);
    }
  }
  return null;
}

/** Read an attachment part URL (mailbox:/imap:) and return its decoded text. */
function fetchAttachmentText(url) {
  return new Promise((resolve, reject) => {
    let channel;
    try {
      channel = NetUtil.newChannel({ uri: url, loadUsingSystemPrincipal: true });
    } catch (error) {
      reject(error);
      return;
    }
    NetUtil.asyncFetch(channel, (inputStream, status) => {
      try {
        // A failed fetch throws when we try to read the stream.
        const text = NetUtil.readInputStreamToString(
          inputStream,
          inputStream.available(),
          { charset: "UTF-8", replacement: "�" }
        );
        resolve(text);
      } catch (error) {
        reject(error);
      }
    });
  });
}

/** For diagnostics: a compact list of a message's attachment content types. */
function describeParts(mimeMessage) {
  try {
    const atts = mimeMessage.allAttachments || [];
    if (!atts.length) {
      return "(no attachments)";
    }
    return atts
      .map(a => `${a.contentType || "?"}${a.name ? ` "${a.name}"` : ""}`)
      .join(", ");
  } catch (error) {
    return "(unavailable)";
  }
}

// -----------------------------------------------------------------------------
// Manual backfill: scan existing Inbox mail for invitations
// -----------------------------------------------------------------------------

// Safety cap so a huge inbox can't spawn an unbounded number of MIME parses.
const BACKFILL_MAX_MESSAGES = 500;

/**
 * Scan the Inbox folder(s) of every account for invitation emails and load any
 * that aren't already in the calendar. Triggered when the user presses the
 * calendar's Synchronize/reload button. Idempotent: events already present are
 * skipped by UID.
 *
 * To keep it fast we only look at messages flagged as having an attachment
 * (invitations carry the .ics as an attachment), capped at BACKFILL_MAX_MESSAGES.
 */
function backfillInvites() {
  let scanned = 0;
  let inboxes = 0;
  try {
    for (const server of MailServices.accounts.allServers) {
      const root = server.rootFolder;
      if (!root) {
        continue;
      }
      const inboxFolders = root.getFoldersWithFlags(Ci.nsMsgFolderFlags.Inbox);
      for (const folder of inboxFolders) {
        inboxes++;
        for (const msgHdr of folder.messages) {
          // Cheap DB-level filter: only messages that carry an attachment.
          if (!(msgHdr.flags & Ci.nsMsgMessageFlags.Attachment)) {
            continue;
          }
          if (scanned >= BACKFILL_MAX_MESSAGES) {
            console.log(
              LOG,
              `backfill: reached the ${BACKFILL_MAX_MESSAGES}-message cap; ` +
                "stopping scan (older messages not checked)"
            );
            return;
          }
          scanned++;
          // Allow download so we can read the MIME structure / attachment even
          // if the body isn't cached offline yet. Bounded by the attachment
          // filter + the message cap above.
          parseMessageForInvite(msgHdr, true);
        }
      }
    }
    console.log(
      LOG,
      `backfill: scanned ${scanned} attachment-bearing message(s) across ` +
        `${inboxes} inbox folder(s); any new invitations are being added.`
    );
  } catch (error) {
    console.error(LOG, "backfill scan failed:", error);
  }
}

/**
 * Depth-first search for an inline text/calendar part containing a VCALENDAR.
 * (Invitations from Google/Outlook/etc. include the invite as an inline
 * text/calendar alternative part, which MsgHdrToMimeMessage exposes as a
 * MimeBody with a decoded `.body`.)
 *
 * @param {object} part - a MimeMessage/MimeContainer/MimeBody node.
 * @returns {?string} the ICS text, or null.
 */
function findCalendarBody(part) {
  const type = (part.contentType || "").toLowerCase();
  if (
    type == "text/calendar" &&
    typeof part.body == "string" &&
    part.body.includes("BEGIN:VCALENDAR")
  ) {
    return part.body;
  }
  if (Array.isArray(part.parts)) {
    for (const child of part.parts) {
      const found = findCalendarBody(child);
      if (found) {
        return found;
      }
    }
  }
  return null;
}

// -----------------------------------------------------------------------------
// Invitation → calendar item
// -----------------------------------------------------------------------------

function processInvite(icsText) {
  // Only auto-add true invitations. METHOD:REQUEST is "please attend"; we skip
  // REPLY/CANCEL/COUNTER/REFRESH and plain event files that carry no METHOD.
  const method = (icsText.match(/^METHOD:(.+)$/im)?.[1] || "").trim().toUpperCase();
  dbg("calendar METHOD =", method || "(none)");
  if (method != "REQUEST") {
    dbg("not a REQUEST invitation; ignoring");
    return;
  }

  const parser = Cc["@mozilla.org/calendar/ics-parser;1"].createInstance(
    Ci.calIIcsParser
  );
  parser.parseString(icsText);

  for (const item of parser.getItems()) {
    if (item.isEvent && item.isEvent()) {
      addInvitation(item).catch(error =>
        console.error(LOG, "addInvitation failed:", error)
      );
    }
  }
}

async function addInvitation(item) {
  // Pick a calendar that (a) we can add events to and (b) recognises us as an
  // invited attendee — the latter is what makes Thunderbird treat the item as
  // an invitation and draw the dotted `invitation-status` styling. Prefer the
  // default calendar among those.
  const allCals = cal.manager.getCalendars();
  const candidates = allCals.filter(calendar => {
    if (calendar.readOnly || calendar.getProperty("disabled")) {
      return false;
    }
    if (calendar.getProperty("capabilities.events.supported") === false) {
      return false;
    }
    try {
      return cal.itip.getInvitedAttendee(item, calendar) != null;
    } catch (error) {
      return false;
    }
  });

  dbg(
    `event "${item.title}": ${allCals.length} calendar(s) total, ` +
      `${candidates.length} recognise you as an invited attendee`
  );
  if (!candidates.length) {
    console.log(
      LOG,
      "no calendar is associated with this invitation's recipient address " +
        "(set Calendar Properties > Email); skipping:",
      item.title
    );
    return;
  }

  candidates.sort(
    (a, b) =>
      (b.getProperty("calendar-main-default") ? 1 : 0) -
      (a.getProperty("calendar-main-default") ? 1 : 0)
  );
  const calendar = candidates[0];
  dbg(`adding to calendar "${calendar.name}" (uid ${item.id})`);

  // De-duplicate by UID: if this invitation (or a copy the user already acted
  // on) is present, do nothing.
  let existing = null;
  try {
    existing = await calendar.getItem(item.id);
  } catch (error) {
    existing = null;
  }
  if (existing) {
    dbg("already present in calendar; not adding again");
    return;
  }

  const newItem = item.clone();
  newItem.calendar = calendar;

  // Mark OUR attendance as not-yet-answered. This is the honest state and is
  // what triggers the "un-answered invitation" (dotted) styling.
  const attendee = cal.itip.getInvitedAttendee(newItem, calendar);
  if (attendee) {
    const updated = attendee.clone();
    updated.participationStatus = "NEEDS-ACTION";
    newItem.removeAttendee(attendee);
    newItem.addAttendee(updated);
  }

  // Add directly. NOTE: we intentionally do NOT route through Thunderbird's
  // ItipOpListener, so NO iTIP reply email is sent.
  await calendar.addItem(newItem);
  console.log(LOG, "auto-added invitation as NEEDS-ACTION:", newItem.title);
}

// -----------------------------------------------------------------------------
// Styling: make the un-answered-invitation border clearly visible (both themes)
// -----------------------------------------------------------------------------

// Thunderbird already outlines invitation-status="NEEDS-ACTION" items with
// `2px dotted black`, which is invisible on dark event backgrounds. We add a
// theme-adaptive dashed outline using currentColor so it reads in both themes.
const INVITATION_CSS = `
  calendar-event-box[invitation-status="NEEDS-ACTION"],
  calendar-editable-item[invitation-status="NEEDS-ACTION"],
  calendar-month-day-box-item[invitation-status="NEEDS-ACTION"] {
    outline: 2px dashed currentColor !important;
    outline-offset: -2px;
  }
`;

function injectStyle(window) {
  try {
    const doc = window.document;
    if (doc.getElementById(STYLE_ELEMENT_ID)) {
      return;
    }
    const style = doc.createElement("style");
    style.id = STYLE_ELEMENT_ID;
    style.textContent = INVITATION_CSS;
    doc.documentElement.appendChild(style);
  } catch (error) {
    console.error(LOG, "style injection failed:", error);
  }
}

function removeStyle(window) {
  try {
    const style = window.document?.getElementById(STYLE_ELEMENT_ID);
    style?.remove();
  } catch (error) {
    /* window already gone */
  }
}

// -----------------------------------------------------------------------------
// Reload/Synchronize hook: backfill invites when the user syncs calendars
// -----------------------------------------------------------------------------

// The Synchronize toolbar button, menu item and shortcut all route through
// `calendarController.doCommand("calendar_reload_remote_calendars")`
// (calendar-command-controller.js). We wrap doCommand so a sync also scans the
// Inbox for invitations that never got auto-added (arrived before install, or
// while the add-on was disabled).
function patchReloadHook(window) {
  const controller = window.calendarController;
  if (
    !controller ||
    typeof controller.doCommand != "function" ||
    controller[DOCMD_MARKER]
  ) {
    return;
  }
  const original = controller.doCommand;
  controller[DOCMD_MARKER] = original;
  controller.doCommand = function (command) {
    const result = original.apply(this, arguments);
    if (command == RELOAD_COMMAND) {
      dbg("sync pressed → scanning Inbox for un-loaded invitations");
      try {
        backfillInvites();
      } catch (error) {
        console.error(LOG, "backfill trigger failed:", error);
      }
    }
    return result;
  };
  dbg("reload/synchronize hook installed");
}

function unpatchReloadHook(window) {
  const controller = window.calendarController;
  if (controller && controller[DOCMD_MARKER]) {
    controller.doCommand = controller[DOCMD_MARKER];
    delete controller[DOCMD_MARKER];
  }
}

/**
 * Set up per-window pieces (CSS + reload hook). calendarController is defined by
 * the calendar's chrome scripts at window load, which may be slightly after
 * onLoadWindow fires, so poll briefly for it.
 */
function setupWindow(window, attemptsLeft = 40) {
  if (window.closed) {
    return;
  }
  injectStyle(window);
  if (window.calendarController && typeof window.calendarController.doCommand == "function") {
    patchReloadHook(window);
    return;
  }
  if (attemptsLeft <= 0) {
    return;
  }
  window.setTimeout(() => setupWindow(window, attemptsLeft - 1), 250);
}

// -----------------------------------------------------------------------------
// Lifecycle
// -----------------------------------------------------------------------------

let listenerAdded = false;

this.autoInvite = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    return {
      autoInvite: {
        enable() {
          // Catch and surface the real error here — the extension framework
          // otherwise replaces it with a generic "unexpected error" message.
          // Global new-mail listener (installed once).
          try {
            if (!listenerAdded) {
              MailServices.mfn.addListener(
                folderListener,
                Ci.nsIMsgFolderNotificationService.msgsClassified
              );
              listenerAdded = true;
              dbg("new-mail listener registered (msgsClassified)");
            }
          } catch (error) {
            console.error(
              LOG,
              "step 1 (mfn.addListener) failed:",
              error?.message || error,
              "\n",
              error?.stack || ""
            );
          }
          // Per-window setup: dotted-invitation CSS + the Synchronize→backfill
          // hook. onLoadWindow fires for current and future main windows.
          try {
            ExtensionSupport.registerWindowListener(MAIN_WINDOW_LISTENER_ID, {
              chromeURLs: [MAIN_WINDOW_URL],
              onLoadWindow(window) {
                setupWindow(window);
              },
            });
            dbg("main-window listener registered");
          } catch (error) {
            console.error(
              LOG,
              "step 2 (registerWindowListener) failed:",
              error?.message || error,
              "\n",
              error?.stack || ""
            );
          }
        },
      },
    };
  }

  onShutdown(isAppShutdown) {
    if (isAppShutdown) {
      return;
    }
    if (listenerAdded) {
      try {
        MailServices.mfn.removeListener(folderListener);
      } catch (error) {
        console.error(LOG, "removeListener failed:", error);
      }
      listenerAdded = false;
    }
    try {
      ExtensionSupport.unregisterWindowListener(MAIN_WINDOW_LISTENER_ID);
    } catch (error) {
      /* not registered */
    }
    try {
      for (const window of Services.wm.getEnumerator("mail:3pane")) {
        removeStyle(window);
        unpatchReloadHook(window);
      }
    } catch (error) {
      console.error(LOG, "window cleanup failed:", error);
    }
  }
};

}).call(this);
