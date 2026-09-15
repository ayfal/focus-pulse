**Plan**
I’d replace the Word macro flow with a Windows PowerShell app that keeps the same behavior pattern from [ThisDocument.txt](ThisDocument.txt#L1), [InitializerForm.frm](InitializerForm.frm#L22), and [Module1.bas](Module1.bas#L32): open the UI, start a fixed work interval, warn when the interval ends, and reschedule the active task. The task store will move from one Word document to one file per task, with each task saved as a separate `.txt` file named by due date using the safe sortable format you picked: `yyyy-MM-dd-hh-mm`.

1. Build the storage layer first.
   - Define a strict filename contract for due dates.
   - Treat file contents as the task body only.
   - Add parsing, sorting, and rename helpers so task persistence is isolated from the GUI.

2. Build the GUI around selection + editing.
   - Use a Windows GUI stack that handles rich text and bidi input well, preferably WPF.
   - Show a task list sorted by due date, a details pane, and a text editor for the selected task.
   - Include controls for start, stop, reschedule, and folder selection.

3. Recreate the timer behavior.
   - Preserve the VBA-style session timer model rather than switching to elapsed-time accounting.
   - Keep the working state, started-at timestamp, and popup reminder flow.
   - Reset the active state and update the status label when the interval completes.

4. Wire rescheduling to the selected task.
   - Reschedule should act on the currently selected task, not the earliest file globally.
   - Saving should update the task body, then rename the file to the new due date.
   - Refresh the list immediately after rename so ordering stays consistent.

5. Add startup, shutdown, and validation.
   - Load the task folder on launch and stop timers cleanly on exit.
   - Add focused tests for filename parsing, sorting, and rename behavior if the script is split into helper functions.
   - Manually verify both left-to-right and right-to-left editing with mixed-direction text.

**Decisions**
- Use the safe filename format `yyyy-MM-dd-hh-mm` because the literal VBA format is not valid on Windows.
- Keep counters and reminders session-scoped, matching the current macro’s behavior.
- Make rescheduling operate on the selected task because the PowerShell UI is editor-centric rather than Word-paragraph-centric.
- Keep persistence file-based only; do not reintroduce a single master document.

If you want, I can next turn this into an implementation checklist with the exact PowerShell files and function boundaries to create.
