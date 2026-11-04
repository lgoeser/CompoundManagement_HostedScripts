"""
Desktop GUI for the SampleSearch reformatter (no external dependencies).

Run with: python desktop_app.py
"""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox

from sample_search_reformat import reformat_sample_search

IDLE_TEXT = "Select a Sample Search export (.xlsx) to reformat."


def main():
    root = tk.Tk()
    root.title("Sample Search Reformatter")
    root.geometry("360x120")

    status_label = tk.Label(root, text=IDLE_TEXT, wraplength=320, justify="center")
    status_label.pack(pady=15)

    select_button = tk.Button(root, text="Select File...")
    select_button.pack(pady=5)

    def on_done(message, is_error):
        if is_error:
            select_button.config(state="normal")
            status_label.config(text=IDLE_TEXT)
            messagebox.showerror("Missing columns", message)
        else:
            messagebox.showinfo("Done", message)
            root.destroy()

    def run_reformat(input_path):
        # Runs on a background thread so the window stays responsive
        # while pandas/openpyxl do the actual work.
        try:
            output_path = reformat_sample_search(input_path)
        except ValueError as e:
            root.after(0, on_done, str(e), True)
        else:
            root.after(0, on_done, f"Formatted file saved to:\n{output_path}", False)

    def select_and_reformat():
        input_path = filedialog.askopenfilename(
            title="Select Sample Search export",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not input_path:
            return

        select_button.config(state="disabled")
        status_label.config(text="Processing, please wait...")
        threading.Thread(target=run_reformat, args=(input_path,), daemon=True).start()

    select_button.config(command=select_and_reformat)

    root.mainloop()


if __name__ == "__main__":
    main()
