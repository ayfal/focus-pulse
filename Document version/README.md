# **Using FocusPulse in Word**

## **1. The scheduler**

You probably have a lot of tasks to do in your life. Some of them are recurring, and belong in a timetable. Some of them are one-off tasks that you schedule in your calendar. And some of them are unscheduled tasks that you just have to get done sometime, and you keep them in a tasklist. Somehow they all need to fit together. The novel idea here is to put them all in one document, and schedule the unscheduled tasks as well. You don't have to put much effort into scheduling, as you'll continuously reschedule tasks as you go.

The FocusPulse form has 3 rescheduling buttons, They all reschedule the first task in the document. For recurring tasks, press "Next Day" or "Next Week" and they will be rescheduled to their next occurrence. For one-off tasks, press "Next Task" and it will be rescheduled for a few minutes later (according to the timer interval you set in the form). You can of course edit the document manually to reschedule tasks as you see fit, and use Word's sorting function to keep tasks in order.

## **2. The timer**

When you open the document, the FocusPulse form appears, it indicates that the timer is running in the background.

In the document version, the timer doesn't just nudge you periodically, like in the browser and app versions, it also tracks how long you've been working and how long you've been slacking off.

You tell the timer you are working by clicking the "Start" button on the form, and the message next to the button changes to "Work until (the set time)".

When that time comes, you'll get a focus-pulse: a message box appears telling you to reschedule this task, and what are your current stats.
When you click OK, you're back to "not working" mode, and the message next to the button changes to "Get to work!".
The timer keeps running in the background in the same pace, and if it makes another focus-pulse while you're not working, the message box will remind you to get to work, and it will count that lapse as slacking off time.

When you click "Start" again, the message changes to "Work until (the set time)" again, that timer is killed and a new one is started. If you finish a task early, you can just click "Start" again to reset the timer.

You can kill the timer by closing the form (clicking the X in the top-right corner of the form), or the document.

### **Word of caution**
To lower the CPU impact, the timer uses Windows API calls to set up a timer that runs in the background, instead of using a VBA timer. If you edit the code in the VBA editor while the timer is running, the code will be recompiled, but the timer will keep running in the background with the old code. When the timer fires, it will try to run the old code, which is no longer there, and this will crash all open documents in Word! Stop the timer first by closing the form, then edit the code. You have been warned.

# **Instructions: How to Install FocusPulse in Word**

## **1. Download the VBA files**

Download all files from the `FocusPulse/word-vba` folder:

* `InitializerForm.frm`
* `InitializerForm.frx`
* `Module1.bas`
* `ThisDocument.txt`

Keep them together in the same folder.

---

## **2. Create a new macro-enabled Word file**

Because Word does not allow macros in `.docx`, you must use: **.docm** — macro-enabled document

**Steps:**

1. Open Word
2. Click **File → New → Blank document**
3. Save as:
   **File → Save As → Word Macro-Enabled Document (*.docm)**

---

## **3. Open the VBA editor**

Press: **ALT + F11**

This opens the **VBA Editor**.

---

## **4. Open the Project Explorer**

If it’s not visible, press: **CTRL + R**

You will see something like:

```
Project (Document1)
    Microsoft Word Objects
        ThisDocument
    References
```

---

## **5. Import the VBA files**

### **Import the module (.bas)**

1. Right-click the project (not a folder)
   Example: `Project (Document1)`
2. Click **Import File…**
3. Select the `Module1.bas` file you downloaded earlier.

The module will now appear under **Modules** in the Project Explorer.
---

### **Import the UserForm (.frm + .frx)**

1. Right-click the project (not a folder)
   Example: `Project (Document1)`
2. Click **Import File…**
3. Select the `InitializerForm.frm` file (the `InitializerForm.frx` file loads automatically)

The form will now appear under **Forms** in the Project Explorer.

---

### **Copy the code for ThisDocument**
1. Double-click **ThisDocument** under **Microsoft Word Objects**
2. Copy all text from `ThisDocument.txt` that you downloaded earlier and paste it into the code window that opened in the VBA editor.

---

## **6. Close and save the document**
Close the Word document (click the X) and press Yes to save changes when prompted.

---

## **7. Reopen the document and Enable macros**

When you open your `.docm` file, Word will show:

> **Security Warning: Macros have been disabled**

Click: **Enable Content**