import time
import tkinter as tk
from tkinter import messagebox
from tkinter import simpledialog


def main():
    # Initialize the timer, task list and the windows:
    alert_root = tk.Tk() # Create a root window for the alert
    alert_root.withdraw()  # Hide the root window
    seconds = simpledialog.askinteger("Timer", "Enter the number of minutes for the timer: ", initialvalue=15) * 60
    task_list = simpledialog.askstring("Tasks", "Enter a list of tasks (separated by commas): ").split(",")
    progress_root = tk.Tk() # Create a progress window
    progress_root.withdraw() # Hide the progress window
    progress_window = tk.Toplevel(progress_root) # Create a toplevel window for the progress
    progress_window.title("Progress") # Set the title of the progress window
    progress_label = tk.Label(progress_window) # Create a label for the progress
    progress_label.pack() # Display the progress label

    # Update the progress label every second:
    while True:
        for task in task_list:
            for i in range(seconds):
                progress_label.config(text=f"Task: {task}\nTime remaining: {(seconds-i) // 60}:{(seconds-i) % 60} minutes")
                progress_root.update()
                time.sleep(1)

            # Show a message box when each task's timer is done:
            messagebox.showinfo("Time's up!", f"Next task: {task_list[(task_list.index(task) + 1) % len(task_list)]}", parent=alert_root)
            alert_root.attributes('-topmost', True)  # Make messagebox appear on top


if __name__ == "__main__":
    main()