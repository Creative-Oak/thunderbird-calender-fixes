"use strict";

// The entire fix lives in the privileged Experiment API (experiments/dragEvent).
// A plain MailExtension cannot reach the calendar's internal controller object,
// so all this background page does is switch the Experiment on.
//
// We log success/failure to the Browser Console (Tools ▸ Developer Tools ▸
// Error Console, or Ctrl/Cmd+Shift+J) so problems are visible instead of silent.
(async () => {
  try {
    await browser.dragEvent.enable();
    console.log("[drag-event-editor] enabled: calendar drag now opens the full event editor");
  } catch (error) {
    console.error("[drag-event-editor] failed to enable:", error);
  }
})();
