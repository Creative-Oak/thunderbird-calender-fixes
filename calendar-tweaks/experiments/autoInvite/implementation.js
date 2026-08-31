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
// * Extract the invite: we stream the message's RAW rfc822 source
//   (nsIMsgMessageService.streamMessage with aConvertData=false) and parse it
//   with MimeParser (resource:///modules/mimeParser.sys.mjs). We deliberately
//   avoid gloda's MsgHdrToMimeMessage / libmime: Thunderbird registers a MIME
//   converter for text/calendar, so libmime-based parsing can transform or hide
//   the invitation part (Outlook/Exchange invites especially). Raw parsing sees
//   both inline text/calendar parts and .ics attachments uniformly.
// * Only METHOD:REQUEST is auto-added; METHOD:CANCEL removes the event; other
//   methods (REPLY/COUNTER/REFRESH) and non-invitation .ics files are ignored.
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
var { MimeParser } = ChromeUtils.importESModule(
  "resource:///modules/mimeParser.sys.mjs"
);
var { cal } = ChromeUtils.importESModule(
  "resource:///modules/calendar/calUtils.sys.mjs"
);

const LOG = "[calendar-tweaks/auto-invite]";
// Flip to false to quiet the step-by-step diagnostics once things work.
const DEBUG = true;
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
  parseMessageForInvite(msgHdr);
}

/**
 * Read a message's RAW source and, if it carries a calendar invitation, process
 * it. Shared by the on-arrival listener and the manual backfill scan.
 *
 * We parse the raw RFC822 source ourselves with MimeParser instead of gloda's
 * MsgHdrToMimeMessage. That matters: Thunderbird registers a MIME converter for
 * text/calendar, so libmime-based parsing (what gloda uses) can transform or
 * hide the invitation part — Outlook/Exchange invites in particular then become
 * invisible. Streaming the raw source (aConvertData=false) bypasses libmime
 * entirely, and one code path then handles both inline text/calendar parts and
 * .ics attachments.
 *
 * @param {nsIMsgDBHdr} msgHdr
 */
async function parseMessageForInvite(msgHdr) {
  const subject = msgHdr.mime2DecodedSubject || msgHdr.subject;
  let rawSource;
  try {
    rawSource = await streamRawMessage(msgHdr);
  } catch (error) {
    dbg("could not read message source:", subject, "-", error?.message || error);
    return;
  }
  let ics;
  try {
    ics = extractCalendarFromRaw(rawSource);
  } catch (error) {
    console.error(LOG, "raw parse failed:", error);
    return;
  }
  if (ics) {
    processInvite(ics);
  } else {
    dbg("no calendar part in:", subject);
  }
}

/** Stream a message's raw RFC822 source (no libmime conversion). */
function streamRawMessage(msgHdr) {
  return new Promise((resolve, reject) => {
    let uri;
    try {
      uri = msgHdr.folder.getUriForMsg(msgHdr);
    } catch (error) {
      reject(error);
      return;
    }
    const service = MailServices.messageServiceFromURI(uri);
    let data = "";
    const listener = {
      QueryInterface: ChromeUtils.generateQI([
        "nsIStreamListener",
        "nsIRequestObserver",
      ]),
      onStartRequest() {},
      onDataAvailable(request, inputStream, offset, count) {
        const sis = Cc["@mozilla.org/scriptableinputstream;1"].createInstance(
          Ci.nsIScriptableInputStream
        );
        sis.init(inputStream);
        data += sis.read(count);
      },
      onStopRequest(request, statusCode) {
        if (data.length) {
          resolve(data);
        } else {
          reject(new Error("empty stream (status " + statusCode + ")"));
        }
      },
    };
    try {
      // aConvertData=false → raw source; aLocalOnly=false → may fetch if needed.
      service.streamMessage(uri, listener, null, null, false, "", false);
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * Parse raw RFC822 source and return the ICS text of the first calendar part
 * (inline text/calendar or an .ics/application/ics attachment), or null.
 */
function extractCalendarFromRaw(rawSource) {
  const parts = new Map(); // partNum -> { type, data }
  const emitter = {
    startPart(partNum, headers) {
      let type = "";
      try {
        if (headers.has("content-type")) {
          type = (headers.contentType.type || "").toLowerCase();
        }
      } catch (error) {
        /* part without a content-type */
      }
      parts.set(partNum, { type, data: "" });
    },
    deliverPartData(partNum, data) {
      const part = parts.get(partNum);
      if (part) {
        part.data += data;
      }
    },
  };
  // strformat:"unicode" + bodyformat:"decode" → parts arrive transfer-decoded
  // and charset-decoded to a JS string.
  MimeParser.parseSync(rawSource, emitter, {
    strformat: "unicode",
    bodyformat: "decode",
  });

  const isCalType = t =>
    t == "text/calendar" || t == "application/ics" || t.includes("calendar");
  // Prefer a properly-typed calendar part...
  for (const { type, data } of parts.values()) {
    if (isCalType(type) && data && data.includes("BEGIN:VCALENDAR")) {
      return data;
    }
  }
  // ...otherwise any part whose decoded body is actually an iCalendar object
  // (covers generic content types like application/octet-stream for .ics).
  for (const { data } of parts.values()) {
    if (data && data.includes("BEGIN:VCALENDAR") && /^METHOD:/im.test(data)) {
      return data;
    }
  }
  return null;
}

// -----------------------------------------------------------------------------
// Manual backfill: scan existing Inbox mail for invitations
// -----------------------------------------------------------------------------

// Safety cap so a huge inbox can't spawn an unbounded number of parses.
const BACKFILL_MAX_MESSAGES = 500;

/**
 * Scan the Inbox folder(s) of every account for invitation emails and load any
 * that aren't already in the calendar. Triggered when the user presses the
 * calendar's Synchronize/reload button. Idempotent: events already present are
 * skipped by UID.
 *
 * To keep it fast we only look at messages flagged as having an attachment
 * (invitations carry the .ics), capped at BACKFILL_MAX_MESSAGES.
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
          parseMessageForInvite(msgHdr);
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

// -----------------------------------------------------------------------------
// Invitation → calendar item
// -----------------------------------------------------------------------------

function processInvite(icsText) {
  // METHOD:REQUEST is "please attend" → add. METHOD:CANCEL is "it's off" →
  // remove the matching event. Everything else (REPLY/COUNTER/REFRESH, or a
  // plain event file with no METHOD) is left to Thunderbird's invitation bar.
  const method = (icsText.match(/^METHOD:(.+)$/im)?.[1] || "").trim().toUpperCase();
  dbg("calendar METHOD =", method || "(none)");
  if (method != "REQUEST" && method != "CANCEL") {
    dbg("not a REQUEST/CANCEL invitation; ignoring");
    return;
  }

  const parser = Cc["@mozilla.org/calendar/ics-parser;1"].createInstance(
    Ci.calIIcsParser
  );
  parser.parseString(icsText);

  for (const item of parser.getItems()) {
    if (!(item.isEvent && item.isEvent())) {
      continue;
    }
    const action = method == "CANCEL" ? removeInvitation : addInvitation;
    action(item).catch(error =>
      console.error(LOG, `${action.name} failed:`, error)
    );
  }
}

/**
 * Handle a METHOD:CANCEL: delete the matching event (by UID) from any calendar
 * that holds it. This covers both our auto-added tentative copy and one the
 * user has already accepted, matching Thunderbird's normal cancellation
 * behavior. (Whole-event cancellations only; per-occurrence cancels are left to
 * the invitation bar.)
 */
async function removeInvitation(item) {
  let removed = 0;
  for (const calendar of cal.manager.getCalendars()) {
    if (calendar.readOnly) {
      continue;
    }
    let existing = null;
    try {
      existing = await calendar.getItem(item.id);
    } catch (error) {
      existing = null;
    }
    if (!existing) {
      continue;
    }
    try {
      await calendar.deleteItem(existing);
      removed++;
      console.log(LOG, "removed cancelled invitation:", existing.title);
    } catch (error) {
      console.error(LOG, "deleteItem failed:", error);
    }
  }
  if (!removed) {
    dbg("cancellation for an event not in any calendar; nothing to remove:", item.title);
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
