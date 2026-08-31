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

const LOG = "[calendar-tweaks/auto-invite]";
// Flip to false to quiet the step-by-step diagnostics once things work.
const DEBUG = true;
function dbg(...args) {
  if (DEBUG) {
    console.log(LOG, ...args);
  }
}
const MAIN_WINDOW_URL = "chrome://messenger/content/messenger.xhtml";
const CSS_WINDOW_LISTENER_ID = "autoInvite-cssListener";
const STYLE_ELEMENT_ID = "autoInvite-invitation-style";

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
  // Parse the MIME structure (allowDownload=true so IMAP bodies not yet cached
  // are fetched). The callback runs asynchronously.
  MsgHdrToMimeMessage(
    msgHdr,
    null,
    (hdr, mimeMessage) => {
      try {
        if (!mimeMessage) {
          dbg("no MIME message returned (not downloaded?)");
          return;
        }
        const ics = findCalendarBody(mimeMessage);
        if (ics) {
          dbg("found inline text/calendar part");
          processInvite(ics);
        } else if (hasCalendarAttachment(mimeMessage)) {
          // Some senders (notably Outlook) ship the invitation only as an
          // attachment (invite.ics) rather than an inline body. v1 parses the
          // inline form; log so we know this is the case.
          dbg(
            "invitation present only as an attachment (not inline); not parsed in this version"
          );
        }
      } catch (error) {
        console.error(LOG, "invite processing failed:", error);
      }
    },
    true,
    { partsOnDemand: false }
  );
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

/**
 * Whether the message carries a text/calendar part in any form (including as a
 * non-inline attachment). Used only for diagnostics.
 */
function hasCalendarAttachment(mimeMessage) {
  try {
    const all = mimeMessage.allAttachments || [];
    return all.some(a => (a.contentType || "").toLowerCase() == "text/calendar");
  } catch (error) {
    return false;
  }
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
          // Per-window CSS for the dotted invitation border. onLoadWindow fires
          // for current and future main windows.
          try {
            ExtensionSupport.registerWindowListener(CSS_WINDOW_LISTENER_ID, {
              chromeURLs: [MAIN_WINDOW_URL],
              onLoadWindow: injectStyle,
            });
            dbg("CSS window listener registered");
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
      ExtensionSupport.unregisterWindowListener(CSS_WINDOW_LISTENER_ID);
    } catch (error) {
      /* not registered */
    }
    try {
      for (const window of Services.wm.getEnumerator("mail:3pane")) {
        removeStyle(window);
      }
    } catch (error) {
      console.error(LOG, "style cleanup failed:", error);
    }
  }
};

}).call(this);
