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

* `FocusPulseForm.frm`
* `FocusPulseForm.frx`
* `Module_FocusPulse.bas`

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

1. Right-click **Modules**
2. Select **Import File…**
3. Choose the `.bas` file

(If "Modules" does not exist, right-click the project → **Insert → Module** → then right-click "Modules".)

---

### **Import the UserForm (.frm + .frx)**

1. Right-click the project (not a folder)
   Example: `Project (Document1)`
2. Click **Import File…**
3. Select the `.frm` file (the `.frx` file loads automatically)

The form will now appear under **Forms** in the Project Explorer.

---

## **6. Save and close the VBA editor**

Click **File → Save** inside the VBA editor, then close it.

---

## **7. Enable macros**

When you open your `.docm` file, Word will show:

> **Security Warning: Macros have been disabled**

Click:

**Enable Content**

FocusPulse will not run unless macros are enabled.

---

## **8. Reopen the document**

* The form will appear automatically on document open

