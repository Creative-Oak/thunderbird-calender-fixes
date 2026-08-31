"use strict";

// All logic lives in the privileged Experiment APIs; a plain MailExtension
// cannot reach the calendar controller or the mail notification service. This
// background page just switches the fixes on.
//
// Logs go to the Browser Console (Tools ▸ Developer Tools ▸ Error Console, or
// Ctrl/Cmd+Shift+J) so problems are visible instead of silent.
(async () => {
  try {
    await browser.dragEvent.enable();
    console.log("[calendar-tweaks] drag → full event editor enabled");
  } catch (error) {
    console.error("[calendar-tweaks] dragEvent failed to enable:", error);
  }

  try {
    await browser.autoInvite.enable();
    console.log("[calendar-tweaks] auto-add emailed invitations enabled");
  } catch (error) {
    console.error("[calendar-tweaks] autoInvite failed to enable:", error);
  }
})();
