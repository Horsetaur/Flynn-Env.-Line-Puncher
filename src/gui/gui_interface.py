import tkinter as tk
from tkinter import messagebox
try:
    import keyboard  # global hotkeys
except Exception:
    keyboard = None


class LinePuncherGUI:
    def __init__(self, on_add_row, on_add_category):
        self.on_add_row = on_add_row
        self.on_add_category = on_add_category
        self.root = tk.Tk()
        self.root.title("Flynn Line Puncher")

        btn_row = tk.Button(self.root, text="Add Row to Category", width=24, command=self._call(self.on_add_row))
        btn_row.pack(padx=12, pady=8)

        btn_cat = tk.Button(self.root, text="Add New Category", width=24, command=self._call(self.on_add_category))
        btn_cat.pack(padx=12, pady=4)

        quit_btn = tk.Button(self.root, text="Quit", width=24, command=self.root.destroy)
        quit_btn.pack(padx=12, pady=8)

        # Bring window to front and center it shortly after launch
        self._bring_to_front()
        self.root.after(200, self._center_window)

    def _call(self, fn):
        def handler():
            self._safe_call(fn)
        return handler

    def _safe_call(self, fn):
        try:
            fn()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run(self):
        # Global hotkeys: Ctrl+Alt+A for Add Row, Ctrl+Alt+C for New Category
        if keyboard is not None:
            try:
                keyboard.add_hotkey('ctrl+alt+a', self._call(self.on_add_row))
                keyboard.add_hotkey('ctrl+alt+c', self._call(self.on_add_category))
            except Exception:
                pass
        self.root.mainloop()

    def _bring_to_front(self) -> None:
        try:
            self.root.lift()
            self.root.attributes('-topmost', True)
            # Release topmost after a short time so it doesn't stay above Excel
            self.root.after(1500, lambda: self.root.attributes('-topmost', False))
        except Exception:
            pass

    def _center_window(self) -> None:
        try:
            self.root.update_idletasks()
            w = self.root.winfo_width()
            h = self.root.winfo_height()
            ws = self.root.winfo_screenwidth()
            hs = self.root.winfo_screenheight()
            x = int((ws - w) / 2)
            y = int((hs - h) / 3)
            self.root.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass


if __name__ == "__main__":
    def _dummy_row():
        print("Add Row clicked")

    def _dummy_cat():
        print("Add Category clicked")

    LinePuncherGUI(_dummy_row, _dummy_cat).run()


