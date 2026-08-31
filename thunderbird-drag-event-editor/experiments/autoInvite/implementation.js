"use strict";

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

const LOG = "[calendar-fixes/auto-invite]";
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
  msgsClassified(msgs /*, junkProcessed, traitProcessed */) {
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
    return;
  }
  // Parse the MIME structure (allowDownload=true so IMAP bodies not yet cached
  // are fetched). The callback runs asynchronously.
  MsgHdrToMimeMessage(
    msgHdr,
    null,
    (hdr, mimeMessage) => {
      try {
        if (!mimeMessage) {
          return;
        }
        const ics = findCalendarBody(mimeMessage);
        if (ics) {
          processInvite(ics);
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

// -----------------------------------------------------------------------------
// Invitation → calendar item
// -----------------------------------------------------------------------------

function processInvite(icsText) {
  // Only auto-add true invitations. METHOD:REQUEST is "please attend"; we skip
  // REPLY/CANCEL/COUNTER/REFRESH and plain event files that carry no METHOD.
  const method = (icsText.match(/^METHOD:(.+)$/im)?.[1] || "").trim().toUpperCase();
  if (method != "REQUEST") {
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
  const candidates = cal.manager.getCalendars().filter(calendar => {
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

  if (!candidates.length) {
    console.log(
      LOG,
      "no calendar is associated with this invitation's address; skipping:",
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

  // De-duplicate by UID: if this invitation (or a copy the user already acted
  // on) is present, do nothing.
  let existing = null;
  try {
    existing = await calendar.getItem(item.id);
  } catch (error) {
    existing = null;
  }
  if (existing) {
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
          // Global new-mail listener (installed once).
          if (!listenerAdded) {
            MailServices.mfn.addListener(
              folderListener,
              Ci.nsIMsgFolderNotificationService.msgsClassified
            );
            listenerAdded = true;
          }
          // Per-window CSS for the dotted invitation border. onLoadWindow fires
          // for current and future main windows.
          ExtensionSupport.registerWindowListener(CSS_WINDOW_LISTENER_ID, {
            chromeURLs: [MAIN_WINDOW_URL],
            onLoadWindow: injectStyle,
          });
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
    for (const window of Services.wm.getEnumerator("mail:3pane")) {
      removeStyle(window);
    }
  }
};
