"use strict";

// =============================================================================
// Calendar Drag → Full Event Editor  (WebExtension Experiment)
// =============================================================================
//
// WHAT THIS PATCHES
// -----------------
// Thunderbird's calendar day/week views handle "drag across a time range to
// create an event" internally. Once the drag finishes, the view has already
// resolved the start/end times (snapping, timezone, cross-day, etc.) and calls:
//
//     col.calendarView.controller.createNewEvent(null, newStart, newEnd);
//         (calendar/base/content/calendar-multiday-view.js)
//
// `controller` is the single shared object `calendarViewController`, defined in
//     calendar/base/content/calendar-views-utils.js
// and assigned onto each view in calendar/base/content/calendar-tabs.js. Its
// createNewEvent() looks (as of Thunderbird 128–current) like this:
//
//     createNewEvent(calendar, startTime, endTime, forceAllday) {
//       // if we're given both times, skip the dialog
//       if (startTime && endTime && !startTime.isDate && !endTime.isDate) {
//         const item = new CalEvent();
//         setDefaultItemValues(item, calendar, startTime, endTime);
//         doTransaction("add", item, item.calendar, null, null);   // <-- inline
//       } else {
//         createEventWithDialog(calendar, startTime, null, null, null, forceAllday);
//       }
//     }
//
// So the ONLY case that creates an event inline (and then shows the inline
// title editor) is a drag that produced two *timed* endpoints. Every other
// entry point — single click, all-day drag, month-view, the New Event command —
// already goes through createEventWithDialog(). That makes the interception
// tiny and low-risk: we wrap createNewEvent and, for that one inline case,
// call createEventWithDialog() with BOTH times instead.
//
// All of calendarViewController, createEventWithDialog, CalEvent, etc. are plain
// globals on the main window (messenger.xhtml loads the calendar scripts with
// <script src="chrome://calendar/content/...">), so we can reach them directly.
//
// WHY A WRAPPER (not a DOM listener): we deliberately do NOT touch drag/mouse
// handling, snapping, timezones or cross-day logic. We intercept exactly after
// Thunderbird has resolved the times, so all of that native behavior is reused.

var { ExtensionCommon } = ChromeUtils.importESModule(
  "resource://gre/modules/ExtensionCommon.sys.mjs"
);
var { ExtensionSupport } = ChromeUtils.importESModule(
  "resource:///modules/ExtensionSupport.sys.mjs"
);

// The main 3-pane window that hosts the calendar UI and its global controller.
const MAIN_WINDOW_URL = "chrome://messenger/content/messenger.xhtml";

// A unique id for our ExtensionSupport window listener.
const WINDOW_LISTENER_ID = "dragEventEditor-windowListener";

// Stashed on the controller object so we can (a) detect an existing patch and
// avoid double-wrapping, and (b) restore the original on shutdown.
const ORIGINAL_MARKER = "_dragEventEditor_originalCreateNewEvent";

// Every window whose controller we successfully patched, so we can restore them.
const patchedWindows = new Set();

/**
 * Wrap calendarViewController.createNewEvent on a single main window.
 *
 * @param {Window} window - a chrome://messenger/content/messenger.xhtml window.
 */
function patchController(window) {
  const controller = window.calendarViewController;

  // Nothing to do if the calendar isn't present, the API changed, or we already
  // patched this controller.
  if (
    !controller ||
    typeof controller.createNewEvent != "function" ||
    controller[ORIGINAL_MARKER]
  ) {
    return;
  }

  const original = controller.createNewEvent;
  controller[ORIGINAL_MARKER] = original;

  controller.createNewEvent = function (calendar, startTime, endTime, forceAllday) {
    // Redirect ONLY the drag-to-create case that stock TB would create inline:
    // both endpoints present and both timed (not all-day). Anything else keeps
    // its original behavior untouched (single click, all-day, month view, the
    // New Event command that calls createEventWithDialog directly, etc.).
    const bothTimed =
      startTime && endTime && !startTime.isDate && !endTime.isDate;

    if (bothTimed) {
      // Open the normal full editor pre-filled with the dragged start AND end.
      //   - Passing `calendar` through preserves the selected/target calendar
      //     (null makes createEventWithDialog fall back to getSelectedCalendar).
      //   - The calIDateTime objects carry their own timezone, so tz behavior
      //     is preserved.
      //   - createEventWithDialog() does NOT persist anything until the user
      //     clicks Save, satisfying "no event until confirmed".
      window.createEventWithDialog(
        calendar,
        startTime,
        endTime,
        null, // summary
        null, // template event
        forceAllday
      );
      return undefined;
    }

    // Defer to stock Thunderbird for every other case.
    return original.call(this, calendar, startTime, endTime, forceAllday);
  };

  patchedWindows.add(window);
}

/**
 * Restore the original createNewEvent on a single window.
 *
 * @param {Window} window
 */
function unpatchController(window) {
  const controller = window.calendarViewController;
  if (controller && controller[ORIGINAL_MARKER]) {
    controller.createNewEvent = controller[ORIGINAL_MARKER];
    delete controller[ORIGINAL_MARKER];
  }
  patchedWindows.delete(window);
}

/**
 * The calendar's chrome scripts define `calendarViewController` at window load,
 * but on a cold start that can land slightly after our onLoadWindow callback
 * fires. Poll briefly until it exists, then patch. If it never appears (e.g. a
 * window without the calendar), we simply give up quietly.
 *
 * @param {Window} window
 * @param {number} attemptsLeft
 */
function patchWhenReady(window, attemptsLeft = 40) {
  if (window.closed) {
    return;
  }
  if (
    window.calendarViewController &&
    typeof window.calendarViewController.createNewEvent == "function"
  ) {
    patchController(window);
    return;
  }
  if (attemptsLeft <= 0) {
    return;
  }
  // ~40 * 250ms = up to 10s grace for the calendar to finish initializing.
  window.setTimeout(() => patchWhenReady(window, attemptsLeft - 1), 250);
}

this.dragEvent = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    return {
      dragEvent: {
        enable() {
          // registerWindowListener is self-guarding: a duplicate id just warns
          // and returns false, so it is safe to call once per background load.
          // onLoadWindow is invoked for already-open windows AND every future
          // one, which covers the "handle newly opened windows" requirement.
          ExtensionSupport.registerWindowListener(WINDOW_LISTENER_ID, {
            chromeURLs: [MAIN_WINDOW_URL],
            onLoadWindow(window) {
              patchWhenReady(window);
            },
            onUnloadWindow(window) {
              // Window is going away; just drop our reference to it.
              patchedWindows.delete(window);
            },
          });
        },
      },
    };
  }

  onShutdown(isAppShutdown) {
    // On full application shutdown the windows are torn down anyway, so there is
    // nothing to restore. On disable/uninstall/update we undo everything so the
    // stock behavior returns immediately, without needing a restart.
    if (isAppShutdown) {
      return;
    }
    ExtensionSupport.unregisterWindowListener(WINDOW_LISTENER_ID);
    for (const window of [...patchedWindows]) {
      unpatchController(window);
    }
  }
};
