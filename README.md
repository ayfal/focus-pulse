# FocusPulse
This is a very simple and surprisingly effective productivity tool:

It just alerts you every few minutes to switch tasks, so you get a new burst of focus on a different task. Also if you get distracted, it will remind you to get back on track.

There are 3 versions of this tool:

1. A simple browser version that you can open in your browser locally and is also available online at https://ayfal.github.io/focus-pulse/Browser%20version/FocusPulse.html. It supports a small semicolon separated list of tasks.

2. A tiny app version that you can run on your computer (you need to have python installed). It supports a small comma separated list of tasks. 

3. A document version. This is for managing an entire task list and calendar. It also introduces a novel way of prioritizing tasks: by scheduling them.
You put your entire schedule and tasks in this document, and have a simple form for the timer and for rescheduling tasks.
I personally use this every day all day, and put the document in windows startup so it opens automatically when I start my computer.

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

Press:

```
ALT + F11
```

This opens the **VBA Editor**.

---

## **4. Open the Project Explorer**

If it’s not visible, press:

```
CTRL + R
```

You will see something like:

```
Project (Document1)
    Microsoft Word Objects
        ThisDocument
    Modules
    Forms
```

---

## **5. Import the VBA files**

### **Import the module (.bas)**

1. Right-click the project (not a folder)
   Example: `Project (Document1)`
2. Click **Import File…**
3. Select the `Module1.bas` file you downloaded earlier.

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

Click:

**Enable Content**


# **Using FocusPulse in Word**

## **1. The timer**

When you open the document, the FocusPulse initializer form appears, it indicates that the timer is running in the background.

In the document version, the timer doesn't just nudge you periodically, like in the browser and app versions, it also tracks how long you've been working and how long you've been slacking off.

You tell the timer you are working by clicking the "Start" button on the form, the message next to the button changes to "Work until (the set time)".

When the timer finishes counting down, a message box appears telling you to reschedule this task, and your current stats.
When you click OK, you're back to "not working" mode, and the message next to the button changes to "Get to work!".
The timer keeps running in the background in the same pace, and if it ticks while you're not working, it will count that as slacking off time.

When you click "Start" again, the message changes to "Work until (the set time)" again, that timer is killed and a new one is started.
You can also kill the timer by closing the form (clicking the X in the top-right corner of the form), or the document.

## **2. The scheduler**

Normally you'd have scheduled tasks and unscheduled tasks in your life. The scheduled tasks are probably written in a calendar, if they're one-off tasks, or in a timetable, if they're recurring tasks. The unscheduled tasks are just a list. Somehow they all need to get along. The novel idea here is to put them all in one document, and schedule the unscheduled tasks as well. You don't have to put much effort into scheduling, as you'll continuously reschedule tasks as you go.

The initializer form has 3 rescheduling buttons, They all reschedule the first task in the document. For recurring tasks, press "Next Day" or "Next Week" and they will be rescheduled to their next occurrence. For one-off tasks, press "Next Task" and it will be rescheduled for a few minutes later (according to the timer interval you set in the form). You can of course edit the document manually to reschedule tasks as you see fit, and use Word's sorting functionality to keep tasks in order.

---

# **Word of caution**
If you edit the code in the VBA editor while the timer is running, you're likely to crash Word. All of your open documents(!) will crash when the running timer will try to bring up the message box. Stop the timer first by closing the form, then edit the code. You have been warned.