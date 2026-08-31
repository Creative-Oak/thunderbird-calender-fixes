# Calendar Tweaks

A minimal Thunderbird add-on bundling two calendar quality-of-life fixes:

1. **Drag → full event editor.** When you drag across a time range in the
   day/week view, open Thunderbird's normal full event editor (pre-filled with
   the dragged start/end) instead of the inline title editor — and keep the
   dragged slot highlighted behind the editor.
2. **Auto-add emailed invitations.** When an email with a calendar invitation
   arrives, add it to your calendar right away as an *un-answered* (dotted)
   item — before you click Accept/Decline — **without sending any reply**.

Each fix is a self-contained WebExtension Experiment; the two are independent.

---

## Feature 1 — Drag → full event editor

### How the patch works

### The interception point

When you finish a drag-to-create in the calendar, Thunderbird has already done
all the hard work (snapping, timezone resolution, cross-day handling) and calls
a single method on the shared calendar view controller:

```js
// calendar/base/content/calendar-multiday-view.js  (MozCalendarEventColumn.onEventSweepMouseUp)
col.calendarView.controller.createNewEvent(null, newStart, newEnd);
```

`controller` is the object **`calendarViewController`**, defined in
`calendar/base/content/calendar-views-utils.js` and assigned onto each view in
`calendar/base/content/calendar-tabs.js`. Its `createNewEvent` looks like this
(Thunderbird 128 → current):

```js
createNewEvent(calendar, startTime, endTime, forceAllday) {
  // if we're given both times, skip the dialog
  if (startTime && endTime && !startTime.isDate && !endTime.isDate) {
    const item = new CalEvent();
    setDefaultItemValues(item, calendar, startTime, endTime);
    doTransaction("add", item, item.calendar, null, null);   // ← inline create
  } else {
    createEventWithDialog(calendar, startTime, null, null, null, forceAllday);
  }
}
```

The **only** code path that creates an event inline (and then shows the inline
title editor) is the first branch: a drag that produced **two timed
endpoints**. Every other entry point — a single click, an all-day drag, the
month view, the "New Event" command — already goes through
`createEventWithDialog()`.

### What we change

The add-on wraps `calendarViewController.createNewEvent`. For the one inline
case (both endpoints present and timed), it calls the existing global
`createEventWithDialog(calendar, startTime, endTime, …)` — passing **both**
times — so the full editor opens. Every other case is delegated, untouched, to
the original method.

```
drag (both timed)   →  createEventWithDialog(calendar, start, end, …)   [full editor]
everything else     →  original createNewEvent(…)                       [unchanged]
```

### Keeping the dragged slot highlighted (cosmetic)

While you drag, Thunderbird draws a "shadow" box over the swept time range (the
`.fgdragbox` element, shown via a `dragging` attribute). Natively that box is
cleared on mouse-up, *before* `createNewEvent` runs — so once we open a modeless
dialog, the calendar behind it would show no trace of the selection.

The add-on optionally re-draws that same shadow box and keeps it until the
editor closes. It does **not** recompute any geometry: the drag's geometry lives
on the column's `mDragState` (`startMin`/`endMin`/`offset`/`shadows`) and is
drawn by the column's own `updateColumnShadows()`. We snapshot `mDragState` just
before Thunderbird wipes it (by wrapping the column's `clearDragging`), then
briefly restore it and call the native `updateColumnShadows()` to redraw. The
highlight is cleared when the editor window closes (or on the next click in the
main window).

This whole feature is cosmetic and wrapped in `try/catch`: if it ever breaks on
a future Thunderbird, the core "open the editor" behavior is unaffected.

### Why this is safe and minimal

- **No drag/mouse re-implementation.** We hook the point *after* Thunderbird
  has resolved the start/end times, so snapping, timezones, cross-day and
  all-day logic are all reused, not rebuilt.
- **Preserves the calendar.** The `calendar` argument is passed straight
  through; `createEventWithDialog` falls back to the selected calendar when it
  is `null`, exactly like stock behavior.
- **Preserves timezone / all-day.** The `calIDateTime` objects carry their own
  timezone. All-day drags have `startTime.isDate === true`, so they fail the
  "both timed" test and fall through to the original code (which already opens
  the dialog).
- **Nothing is saved until you confirm.** `createEventWithDialog` only persists
  the event in its own OK/Save callback.
- **Reversible.** The original method is stashed and restored when the add-on
  is disabled/uninstalled — no restart needed.

### Why a WebExtension Experiment

`calendarViewController` is an internal chrome object on the main Thunderbird
window (`messenger.xhtml` loads the calendar scripts with plain
`<script src="chrome://calendar/content/…">`). A standard MailExtension cannot
reach it, so the add-on uses a **WebExtension Experiment** (`experiments/dragEvent`)
which runs privileged code. The `background.js` page does nothing but switch the
Experiment on; all logic lives in `experiments/dragEvent/implementation.js`.

The Experiment uses `ExtensionSupport.registerWindowListener` to patch the
current main window and any window opened later, and restores the original
method on shutdown.

---

## Feature 2 — Auto-add emailed invitations

When an email containing a calendar invitation (an iMIP `METHOD:REQUEST`)
arrives, the add-on adds the event to your calendar immediately as an
**un-answered invitation** — drawn with a dotted/dashed outline — *before* you
click Accept or Decline, and **without sending any reply email**. You still
accept/decline normally from the message's invitation bar; that updates the same
item in place (no duplicate).

### How it works (verified against comm-central tip)

- **Detect new mail.** Listens on `MailServices.mfn` with the `msgsClassified`
  flag (`nsIMsgFolderNotificationService`), which fires once per newly-arrived,
  classified message.
- **Extract the invite.** We stream the message's **raw** RFC822 source
  (`nsIMsgMessageService.streamMessage` with `aConvertData=false`) and parse it
  with `MimeParser` (`resource:///modules/mimeParser.sys.mjs`). We deliberately
  avoid gloda's `MsgHdrToMimeMessage`/libmime: Thunderbird registers a MIME
  converter for `text/calendar`, so libmime-based parsing can transform or hide
  the invitation part (Outlook/Exchange invites especially). Raw parsing handles
  both **inline** `text/calendar` parts and **`.ics` attachments** uniformly.
  `METHOD:REQUEST` is added, `METHOD:CANCEL` removes the event, and other
  methods are ignored.
- **Recognise you.** `cal.itip.getInvitedAttendee(item, calendar)` returns the
  attendee matching a calendar's own email identity — the same check Thunderbird
  uses to treat something as an "invitation" and apply the dotted
  `invitation-status` styling. We add to the first calendar that recognises the
  invite, preferring the **default** calendar.
- **Add without replying.** We set your attendee's participation status to
  `NEEDS-ACTION` ("not answered") and call `calendar.addItem()` **directly**.
  Thunderbird's own Accept path wraps this in `ItipOpListener`, whose completion
  handler emails the iTIP reply — we deliberately never use that wrapper, so
  nothing is sent.
- **No duplicates.** `calendar.getItem(uid)` is the same UID lookup the
  invitation bar uses, so a later Accept finds our copy and *modifies* it rather
  than adding a second one.
- **Styling.** `invitation-status="NEEDS-ACTION"` already gets a dotted outline
  from Thunderbird's CSS; because that outline is hard-coded black (invisible on
  dark event backgrounds), the add-on injects one small theme-adaptive rule
  (`outline: 2px dashed currentColor`) into the calendar window.

### Limitation

The invite is only auto-added if one of your calendars is associated with the
email address the invitation was sent to. This is automatic for CalDAV accounts
(Google, Fastmail, Nextcloud, …). For a **local** calendar, set it in
**Calendar Properties ▸ Email**. If no calendar matches, the message is skipped
(Thunderbird itself would not consider it an invitation either) — you'll see a
`no calendar is associated…` line in the Error Console.

### Scope (v1)

Handles `REQUEST` invitations (adds them) and `CANCEL` messages (removes the
matching event by UID from any calendar that holds it — whole-event
cancellations only). `REPLY`/`COUNTER` and per-occurrence cancellations still go
through Thunderbird's normal invitation bar.

Auto-add happens on **arrival**. For invitations that predate the add-on (or
arrived while it was disabled), pressing the calendar's **Synchronize / reload**
button also **backfills**: it scans the Inbox of every account for
attachment-bearing messages and loads any invitations not already in the
calendar (capped at 500 messages, idempotent — existing events are skipped by
UID). It uses the same raw parser, so both inline and `.ics`-attachment
invitations are imported.

---

## Project layout

```
calendar-tweaks/
├── manifest.json                       # MV2 MailExtension + experiment_apis
├── background.js                       # enables both experiments
├── experiments/
│   ├── dragEvent/
│   │   ├── schema.json                 # dragEvent.enable()
│   │   └── implementation.js           # drag → editor wrapper + highlight
│   └── autoInvite/
│       ├── schema.json                 # autoInvite.enable()
│       └── implementation.js           # new-mail listener + invite auto-add
└── README.md
```

---

## Supported Thunderbird versions

`manifest.json` sets `strict_min_version` to **128.0** (the previous ESR) and no
maximum. The wrapped method (`calendarViewController.createNewEvent` with the
inline `doTransaction` branch) has been stable across Thunderbird 128 through
the current release; it was verified against `comm-central` tip while building
this add-on. If a future Thunderbird renames or restructures the controller, the
add-on degrades gracefully — it simply finds nothing to patch and stock behavior
remains.

---

## Install temporarily for development

Experiments only load from an **unpacked/temporary** install or a signed XPI;
they do not run from a normal drag-and-drop install of an unsigned XPI.

1. Open Thunderbird.
2. Menu ▸ **Tools ▸ Developer Tools ▸ Debug Add-ons**
   (or type `about:debugging#/runtime/this-firefox` in the address of a
   `chrome://` context; easiest is the menu).
3. Click **Load Temporary Add-on…**.
4. Select this folder's **`manifest.json`**.
5. Open the calendar, switch to Day or Week view, and drag across a time range.
   The full event editor should open with the dragged times.

Notes:

- A temporary add-on is removed when Thunderbird restarts — reload it the same
  way after a restart during development.
- To see logs, open **Tools ▸ Developer Tools ▸ Error Console**
  (`Ctrl/Cmd+Shift+J`) and look for `[calendar-tweaks]` messages.
- After editing `implementation.js`, click **Reload** on the add-on in the
  Debug Add-ons page.

---

## Package as an `.xpi`

An XPI is just a ZIP of the add-on's contents with the `manifest.json` at the
**root** of the archive. From inside this folder:

```bash
cd calendar-tweaks
zip -r -FS ../calendar-tweaks.xpi \
  manifest.json background.js experiments \
  -x '*.DS_Store'
```

That produces `calendar-tweaks.xpi` one level up.

Because this add-on contains an Experiment, Thunderbird will **not** run it from
an unsigned XPI installed normally. Your options:

- **Development / personal use:** keep using **Load Temporary Add-on** (above),
  which loads the XPI or the unpacked folder without signing. To install the XPI
  persistently for personal use you can set the pref
  `xpinstall.signatures.required = false` in a build that allows it
  (Daily/Developer Edition), then install the XPI via
  **Tools ▸ Add-ons ▸ gear ▸ Install Add-on From File…**.
- **Distribution:** submit the XPI to
  [addons.thunderbird.net](https://addons.thunderbird.net/) for signing/hosting.

---

## Testing checklist

### Feature 1 — drag → editor

Day or Week view unless noted:

- [ ] **Drag 10:00 → 11:00:** full editor opens with those exact times.
- [ ] **Drag 10:00 → 11:30:** 90-minute duration preserved.
- [ ] **Drag upward/backwards** (release above the start): start/end are in the
      correct chronological order (Thunderbird resolves this before our hook).
- [ ] **Drag across a day boundary** (where supported): times span correctly.
- [ ] **Dragged slot stays highlighted** behind the open editor window.
- [ ] **Highlight clears** after the editor is saved/cancelled (or on next click).
- [ ] **Cancel the editor:** no event remains in the calendar.
- [ ] **Save the editor:** exactly one event is created.
- [ ] **Double-click a time slot / single click-drag of zero length:** normal
      Thunderbird behavior (this is not the "both timed" case).
- [ ] **All-day area drag:** unchanged (still opens the dialog as an all-day
      event).
- [ ] **Edit an existing event** (click / double-click an event): unaffected.
- [ ] **Switch the active calendar**, then drag: the new event targets the
      appropriate/selected calendar.
- [ ] **Restart Thunderbird** (with the add-on installed persistently): still
      works.
- [ ] **Disable / uninstall the add-on:** stock inline drag-create behavior
      returns immediately, no restart required.

### Feature 2 — auto-add invitations

- [ ] **Send yourself an invite** (e.g. from Google Calendar / another account)
      to an address tied to one of your calendars. When it arrives, the event
      appears in the calendar with a dotted/dashed outline, **without** you
      accepting — and no reply is sent (check the organizer doesn't receive an
      RSVP).
- [ ] **Accept it** from the message's invitation bar: the same event turns
      solid; there is exactly **one** copy (no duplicate).
- [ ] **A non-invite `.ics` attachment** (no `METHOD:REQUEST`) is ignored.
- [ ] **An invite to an address not tied to any calendar** is skipped (see the
      `no calendar is associated…` line in the Error Console).
- [ ] **Disable the add-on:** no further invites are auto-added; the injected
      dashed-border style is removed.

---

## Uninstalling

Remove or disable the add-on from **Tools ▸ Add-ons and Themes**. The wrapper
restores the original `createNewEvent` on shutdown, so the calendar reverts to
stock behavior right away.
