"use strict";

// NOTE: Every experiment parent script in this add-on shares ONE global, so
// top-level `const`/`let` names would collide across scripts. Wrap the body in
// an IIFE to keep declarations local; `.call(this)` keeps `this` as the shared
// global so the `this.dragEvent = …` API registration still works.
(function () {

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
//
// OPTIONAL: KEEP THE DRAGGED SLOT HIGHLIGHTED
// -------------------------------------------
// Natively, the drag draws a "shadow" box (the .fgdragbox element, toggled by a
// `dragging` attribute) while sweeping, and clears it on mouseup *before*
// createNewEvent runs. Because we now open a modeless dialog, the calendar
// behind it would otherwise show no trace of the selection. So we optionally
// re-draw that same shadow box and keep it until the dialog closes.
//
// Crucially we do NOT recompute any geometry: the drag's pixel/time geometry is
// stored on the column's `mDragState` (startMin/endMin/offset/shadows) and drawn
// by the column's own `updateColumnShadows()`. We snapshot mDragState just
// before Thunderbird wipes it (by wrapping `clearDragging`), then briefly
// restore it and call the native `updateColumnShadows()` to redraw. This whole
// feature is cosmetic and fully wrapped in try/catch — if any of it fails on a
// future Thunderbird, the core "open the editor" behavior is unaffected.

var { ExtensionCommon } = ChromeUtils.importESModule(
  "resource://gre/modules/ExtensionCommon.sys.mjs"
);
var { ExtensionSupport } = ChromeUtils.importESModule(
  "resource:///modules/ExtensionSupport.sys.mjs"
);

// The main 3-pane window that hosts the calendar UI and its global controller.
const MAIN_WINDOW_URL = "chrome://messenger/content/messenger.xhtml";

// The custom element tag for a single day column in the day/week views.
const EVENT_COLUMN_TAG = "calendar-event-column";

// Document URLs of the event editor window we open (windowed + in-tab iframe).
const EVENT_DIALOG_URLS = [
  "chrome://calendar/content/calendar-event-dialog.xhtml",
  "chrome://calendar/content/calendar-item-iframe.xhtml",
];

// A unique id for our ExtensionSupport window listener.
const WINDOW_LISTENER_ID = "dragEventEditor-windowListener";

// Markers stashed on the patched objects so we can detect an existing patch and
// restore the originals on shutdown.
const CONTROLLER_MARKER = "_dragEventEditor_originalCreateNewEvent";
const COLUMN_MARKER = "_dragEventEditor_originalClearDragging";

// Every window we patched, so we can restore them on shutdown.
const patchedWindows = new Set();

// Transient hand-off between the drag's clearDragging (which owns the geometry)
// and our createNewEvent (which decides whether to redraw). Set and consumed
// within the same synchronous mouseup, so a single slot is enough.
let pendingHighlight = null; // { window, column, view, snapshot }

// -----------------------------------------------------------------------------
// Core: redirect drag-to-create to the full editor
// -----------------------------------------------------------------------------

function patchController(window) {
  const controller = window.calendarViewController;
  if (
    !controller ||
    typeof controller.createNewEvent != "function" ||
    controller[CONTROLLER_MARKER]
  ) {
    return;
  }

  const original = controller.createNewEvent;
  controller[CONTROLLER_MARKER] = original;

  controller.createNewEvent = function (calendar, startTime, endTime, forceAllday) {
    // Redirect ONLY the drag-to-create case that stock TB would create inline:
    // both endpoints present and both timed (not all-day). Anything else keeps
    // its original behavior untouched.
    const bothTimed =
      startTime && endTime && !startTime.isDate && !endTime.isDate;

    if (bothTimed) {
      // Cosmetic: keep the dragged slot highlighted behind the dialog. Must
      // never interfere with actually opening the editor, hence try/catch.
      let highlightedView = null;
      try {
        highlightedView = redrawPendingHighlight(window);
        // Arrange to clear the highlight once the editor closes. This MUST be
        // set up *before* opening the dialog: the editor window opens
        // synchronously inside createEventWithDialog(), so a watcher registered
        // afterwards would miss it.
        if (highlightedView) {
          installHighlightCleanup(window, highlightedView);
        }
      } catch (error) {
        console.error("[calendar-tweaks] highlight setup failed:", error);
      }

      // Open the normal full editor pre-filled with the dragged start AND end.
      //   - Passing `calendar` through preserves the selected/target calendar
      //     (null makes createEventWithDialog fall back to getSelectedCalendar).
      //   - The calIDateTime objects carry their own timezone.
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
    pendingHighlight = null;
    return original.call(this, calendar, startTime, endTime, forceAllday);
  };
}

function unpatchController(window) {
  const controller = window.calendarViewController;
  if (controller && controller[CONTROLLER_MARKER]) {
    controller.createNewEvent = controller[CONTROLLER_MARKER];
    delete controller[CONTROLLER_MARKER];
  }
}

// -----------------------------------------------------------------------------
// Optional: persistent drag highlight
// -----------------------------------------------------------------------------

/**
 * Wrap the event column's clearDragging so we can capture the finished NEW-drag
 * geometry (mDragState) *before* the original wipes it. We store it in
 * `pendingHighlight`; createNewEvent decides whether to actually redraw.
 */
function patchDragHighlight(window) {
  const columnClass = window.customElements?.get(EVENT_COLUMN_TAG);
  if (!columnClass) {
    return; // calendar day/week view not registered; skip cosmetics.
  }
  const proto = columnClass.prototype;
  if (proto[COLUMN_MARKER]) {
    return;
  }

  const originalClear = proto.clearDragging;
  proto[COLUMN_MARKER] = originalClear;

  proto.clearDragging = function () {
    let snapshot = null;
    try {
      const ds = this.mDragState;
      if (ds && ds.dragType == "new") {
        // Only the fields updateColumnShadows() reads. dragOccurrence is
        // intentionally omitted (undefined) so the todo-label branch is skipped.
        snapshot = {
          dragType: "new",
          startMin: ds.startMin,
          endMin: ds.endMin,
          offset: ds.offset,
          shadows: ds.shadows,
        };
      }
    } catch (error) {
      console.error("[calendar-tweaks] snapshot failed:", error);
    }

    // Let Thunderbird do its normal cleanup (removes listeners, clears boxes,
    // nulls mDragState).
    originalClear.call(this);

    pendingHighlight = snapshot
      ? { window, column: this, view: this.calendarView, snapshot }
      : null;
  };
}

function unpatchDragHighlight(window) {
  const columnClass = window.customElements?.get(EVENT_COLUMN_TAG);
  const proto = columnClass?.prototype;
  if (proto && proto[COLUMN_MARKER]) {
    proto.clearDragging = proto[COLUMN_MARKER];
    delete proto[COLUMN_MARKER];
  }
}

/**
 * If a NEW-drag snapshot is waiting for this window, re-draw the native shadow
 * box by briefly restoring mDragState and calling the column's own
 * updateColumnShadows(). Returns the view the highlight was drawn on, or null.
 */
function redrawPendingHighlight(window) {
  const p = pendingHighlight;
  pendingHighlight = null;

  if (!p || p.window !== window || !p.column || !p.column.isConnected) {
    return null;
  }

  const column = p.column;
  const previous = column.mDragState;
  column.mDragState = p.snapshot;
  try {
    column.updateColumnShadows(); // native drawing, native geometry.
  } finally {
    column.mDragState = previous; // restore (null after clearDragging).
  }
  return p.view;
}

/**
 * Remove the highlight from every column of the given view. The shadow box is
 * only visible while the `dragging` attribute is present, so removing that hides
 * it; we also reset the inline sizes/labels to leave the DOM clean.
 */
function clearHighlight(view) {
  for (const column of view.getEventColumns()) {
    const fg = column.fgboxes;
    fg.dragbox.removeAttribute("dragging");
    fg.box.removeAttribute("dragging");
    fg.dragbox.style.removeProperty("height");
    fg.dragbox.style.removeProperty("width");
    fg.dragspacer.style.removeProperty("height");
    fg.dragspacer.style.removeProperty("width");
    fg.startlabel.value = "";
    fg.endlabel.value = "";
  }
}

/**
 * Clear the highlight when the just-opened event dialog closes. Two triggers,
 * both one-shot and idempotent:
 *   1. the editor window's `unload` (the default, windowed editor), and
 *   2. the next mousedown in the main window (covers the "edit in a tab" pref
 *      and any missed dialog-close).
 */
function installHighlightCleanup(window, view) {
  // Use the chrome window's own Services (guaranteed present) rather than
  // relying on the experiment sandbox global.
  const services = window.Services;
  let cleared = false;
  let observer = null;

  const stopWatching = () => {
    if (observer && services) {
      try {
        services.ww.unregisterNotification(observer);
      } catch (error) {
        /* already unregistered */
      }
      observer = null;
    }
  };

  const clearOnce = () => {
    stopWatching();
    if (cleared) {
      return;
    }
    cleared = true;
    try {
      clearHighlight(view);
    } catch (error) {
      console.error("[calendar-tweaks] clearHighlight failed:", error);
    }
  };

  // (1) Clear as soon as the editor window closes. domwindowclosed fires with
  // the closing window as the subject, whose URL still identifies the editor.
  // (Registered before the dialog is opened by the caller.)
  try {
    if (services?.ww) {
      observer = {
        observe(subject, topic) {
          if (topic != "domwindowclosed") {
            return;
          }
          let href = "";
          try {
            href = subject.location?.href || subject.document?.documentURI || "";
          } catch (error) {
            /* window already torn down */
          }
          if (EVENT_DIALOG_URLS.some(url => href.includes(url))) {
            clearOnce();
          }
        },
      };
      services.ww.registerNotification(observer);
      // Safety: never watch forever.
      window.setTimeout(stopWatching, 10 * 60 * 1000);
    }
  } catch (error) {
    console.error("[calendar-tweaks] dialog-close watcher failed:", error);
  }

  // (2) Fallback: any further interaction with the main window ends the hint
  // (covers the "edit in a tab" pref, where no separate window closes).
  try {
    window.addEventListener("mousedown", clearOnce, { capture: true, once: true });
  } catch (error) {
    /* non-fatal */
  }
}

// -----------------------------------------------------------------------------
// Window lifecycle
// -----------------------------------------------------------------------------

function patchWindow(window) {
  // Core behavior (required).
  patchController(window);
  // Cosmetic highlight (optional; must never break the core).
  try {
    patchDragHighlight(window);
  } catch (error) {
    console.error("[calendar-tweaks] highlight patch failed:", error);
  }
  patchedWindows.add(window);
}

function unpatchWindow(window) {
  unpatchController(window);
  try {
    unpatchDragHighlight(window);
  } catch (error) {
    console.error("[calendar-tweaks] highlight unpatch failed:", error);
  }
  patchedWindows.delete(window);
}

/**
 * The calendar's chrome scripts define `calendarViewController` (and register
 * the column custom element) at window load, but on a cold start that can land
 * slightly after our onLoadWindow callback fires. Poll briefly until it exists.
 */
function patchWhenReady(window, attemptsLeft = 40) {
  if (window.closed) {
    return;
  }
  if (
    window.calendarViewController &&
    typeof window.calendarViewController.createNewEvent == "function"
  ) {
    patchWindow(window);
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
      unpatchWindow(window);
    }
  }
};

}).call(this);
