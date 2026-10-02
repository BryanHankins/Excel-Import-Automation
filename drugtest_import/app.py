"""Desktop app: pick a photo, let Claude read it, review/correct the fields, then save."""
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageOps, ImageTk

from .extract import ExtractionError, extract_fields
from .schema import FIELDS, RESULTS, TEST_TYPES, normalize, validate
from .storage import append_record

DATA_FILE = os.environ.get("DRUGTEST_CSV", "DrugTestingOrganizer.csv")
PREVIEW_SIZE = (420, 520)
CHOICES = {"TestType": TEST_TYPES, "Result": RESULTS}
LABELS = {"EmployeeID": "Employee ID", "TestDate": "Test date (YYYY-MM-DD)", "TestType": "Test type"}


class ReviewApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Drug Test Import")
        self.image_path = None
        self.uncertain = set()
        self.preview = None
        self.results = queue.Queue()  # worker thread -> UI thread; Tk isn't thread-safe

        toolbar = ttk.Frame(root, padding=8)
        toolbar.pack(fill="x")
        self.open_button = ttk.Button(toolbar, text="Open image…", command=self.open_image)
        self.open_button.pack(side="left")
        self.status = ttk.Label(toolbar, text=f"Saving to {os.path.abspath(DATA_FILE)}")
        self.status.pack(side="left", padx=12)

        body = ttk.Frame(root, padding=8)
        body.pack(fill="both", expand=True)
        self.image_label = ttk.Label(body, text="No image loaded", width=50, anchor="center")
        self.image_label.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        form = ttk.Frame(body)
        form.grid(row=0, column=1, sticky="n")
        self.vars, self.hints = {}, {}
        for row, key in enumerate(FIELDS):
            ttk.Label(form, text=LABELS.get(key, key)).grid(row=row * 2, column=0, sticky="w")
            var = tk.StringVar()
            if key in CHOICES:
                widget = ttk.Combobox(form, textvariable=var, values=CHOICES[key], width=32)
            else:
                widget = ttk.Entry(form, textvariable=var, width=35)
            widget.grid(row=row * 2, column=1, sticky="w", pady=(4, 0))
            hint = ttk.Label(form, text="", foreground="#b45309")
            hint.grid(row=row * 2 + 1, column=1, sticky="w")
            self.vars[key], self.hints[key] = var, hint

        buttons = ttk.Frame(form)
        buttons.grid(row=len(FIELDS) * 2, column=0, columnspan=2, pady=12, sticky="e")
        ttk.Button(buttons, text="Discard", command=self.clear).pack(side="right")
        self.save_button = ttk.Button(buttons, text="Save record", command=self.save, state="disabled")
        self.save_button.pack(side="right", padx=8)

    def open_image(self):
        path = filedialog.askopenfilename(filetypes=[("Images", "*.jpg *.jpeg *.png *.webp")])
        if not path:
            return
        self.clear()
        try:
            with Image.open(path) as img:
                img = ImageOps.exif_transpose(img)
                img.thumbnail(PREVIEW_SIZE)
                self.preview = ImageTk.PhotoImage(img)
        except OSError as e:
            messagebox.showerror("Can't open image", str(e))
            return
        self.image_path = path
        self.image_label.configure(image=self.preview, text="")
        self.set_busy(True, "Reading handwriting…")
        threading.Thread(target=self._extract, args=(path,), daemon=True).start()
        self.root.after(100, self._poll_result, path)

    def _extract(self, path):
        try:
            self.results.put((path, extract_fields(path), None))
        except (ExtractionError, OSError) as e:
            self.results.put((path, None, str(e)))

    def _poll_result(self, path):
        try:
            done_path, result, error = self.results.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll_result, path)
            return
        if done_path != self.image_path:
            return  # user discarded this image while it was being read
        if error:
            self._on_failed(error)
        else:
            self._on_extracted(result)

    def _on_extracted(self, result):
        record = normalize(result.model_dump())
        for key in FIELDS:
            self.vars[key].set(record[key] or "")
        self.uncertain = set(result.uncertain_fields)
        self.set_busy(False, "Check every field against the photo, then save.")
        self.show_hints()

    def _on_failed(self, message):
        self.set_busy(False, "Could not read the image. Enter the fields manually.")
        messagebox.showerror("Extraction failed", message)

    def set_busy(self, busy, message):
        self.status.configure(text=message)
        self.open_button.configure(state="disabled" if busy else "normal")
        self.save_button.configure(state="disabled" if busy or not self.image_path else "normal")

    def current_record(self):
        return normalize({key: var.get() for key, var in self.vars.items()})

    def show_hints(self, problems=None):
        problems = problems or {}
        for key, hint in self.hints.items():
            if key in problems:
                hint.configure(text=problems[key], foreground="#b91c1c")
            elif key in self.uncertain:
                hint.configure(text="Hard to read — double-check", foreground="#b45309")
            else:
                hint.configure(text="")

    def save(self):
        record = self.current_record()
        problems = validate(record)
        self.show_hints(problems)
        if problems:
            self.status.configure(text="Fix the highlighted fields before saving.")
            return
        if self.uncertain:
            names = ", ".join(LABELS.get(k, k) for k in FIELDS if k in self.uncertain)
            if not messagebox.askyesno("Confirm", f"These fields were hard to read: {names}.\n\nHave you checked them against the photo?"):
                return
        try:
            append_record(record, DATA_FILE)
        except OSError as e:
            messagebox.showerror("Save failed", f"Could not write {DATA_FILE}: {e}\nIs it open in Excel?")
            return
        self.clear()
        self.status.configure(text=f"Saved {record['Name']} to {DATA_FILE}.")

    def clear(self):
        self.image_path, self.preview, self.uncertain = None, None, set()
        self.image_label.configure(image="", text="No image loaded")
        for var in self.vars.values():
            var.set("")
        self.show_hints()
        self.save_button.configure(state="disabled")


def main():
    root = tk.Tk()
    ReviewApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
