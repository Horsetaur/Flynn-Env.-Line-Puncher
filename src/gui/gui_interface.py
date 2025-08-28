import threading
import tkinter as tk
from tkinter import messagebox


class LinePuncherGUI:
    def __init__(self, on_add_row, on_add_category):
        self.on_add_row = on_add_row
        self.on_add_category = on_add_category
        self.root = tk.Tk()
        self.root.title("Flynn Line Puncher")

        btn_row = tk.Button(self.root, text="Add Row to Category", width=24, command=self._wrap(self.on_add_row))
        btn_row.pack(padx=12, pady=8)

        btn_cat = tk.Button(self.root, text="Add New Category", width=24, command=self._wrap(self.on_add_category))
        btn_cat.pack(padx=12, pady=4)

        quit_btn = tk.Button(self.root, text="Quit", width=24, command=self.root.destroy)
        quit_btn.pack(padx=12, pady=8)

    def _wrap(self, fn):
        def handler():
            t = threading.Thread(target=self._safe_call, args=(fn,))
            t.daemon = True
            t.start()
        return handler

    def _safe_call(self, fn):
        try:
            fn()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    def _dummy_row():
        print("Add Row clicked")

    def _dummy_cat():
        print("Add Category clicked")

    LinePuncherGUI(_dummy_row, _dummy_cat).run()


