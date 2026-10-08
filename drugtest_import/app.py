"""Desktop app: pick photos, let Claude read them, review/correct each one, then save."""
import getpass
import os
import queue
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageOps, ImageTk

from .crypto import Cipher, ConfigError
from .extract import ExtractionError, extract_fields
from .forms import DRUG_TEST, FormTemplate, normalize, validate
from .storage import Database

DB_FILE = os.environ.get("DRUGTEST_DB", "drugtest.db")
LEGACY_CSV = os.environ.get("DRUGTEST_CSV", "DrugTestingOrganizer.csv")
ACTOR = f"desktop:{getpass.getuser()}"
PARALLEL_READS = 3
PREVIEW_SIZE = (420, 520)


class ReviewApp:
    def __init__(self, root: tk.Tk, store: Database):
        self.root = root
        self.store = store
        self.root.title("Form Import")
        self.pool = ThreadPoolExecutor(max_workers=PARALLEL_READS)
        self.futures = []
        self.results = queue.Queue()  # worker threads -> UI thread; Tk isn't thread-safe
        self.outcomes = {}  # path -> ((values, uncertain) | None, error | None)
        self.batch, self.index = [], 0
        self.saved = self.skipped = 0
        self.uncertain = set()
        self.preview = None
        self.forms = {f.name: f for f in store.forms()}
        self.form: FormTemplate = store.get_form(DRUG_TEST.key) or next(iter(self.forms.values()))

        toolbar = ttk.Frame(root, padding=8)
        toolbar.pack(fill="x")
        ttk.Label(toolbar, text="Form type").pack(side="left")
        self.form_var = tk.StringVar(value=self.form.name)
        self.form_picker = ttk.Combobox(toolbar, textvariable=self.form_var, values=list(self.forms),
                                        state="readonly", width=24)
        self.form_picker.pack(side="left", padx=(6, 12))
        self.form_picker.bind("<<ComboboxSelected>>", self.change_form)
        self.open_button = ttk.Button(toolbar, text="Open images…", command=self.open_images)
        self.open_button.pack(side="left")
        ttk.Button(toolbar, text="Export to Excel…", command=self.export).pack(side="left", padx=(8, 0))
        self.progress = ttk.Label(toolbar, text="")
        self.progress.pack(side="right")
        self.status = ttk.Label(root, padding=(8, 0), text=f"{store.count()} records in {os.path.abspath(store.path)}")
        self.status.pack(fill="x")

        body = ttk.Frame(root, padding=8)
        body.pack(fill="both", expand=True)
        self.image_label = ttk.Label(body, text="No image loaded", width=50, anchor="center")
        self.image_label.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.form_frame = ttk.Frame(body)
        self.form_frame.grid(row=0, column=1, sticky="n")
        self.build_form()

        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.after(100, self._poll_results)

    def build_form(self):
        """(Re)create the input widgets for the current form type."""
        for child in self.form_frame.winfo_children():
            child.destroy()
        self.vars, self.hints, self.inputs = {}, {}, {}
        for row, field in enumerate(self.form.fields):
            label = field.label + (" (YYYY-MM-DD)" if field.type == "date" else "") + (" *" if field.required else "")
            ttk.Label(self.form_frame, text=label).grid(row=row * 2, column=0, sticky="w")
            var = tk.StringVar()
            if field.type == "choice":
                widget = ttk.Combobox(self.form_frame, textvariable=var, values=field.choices, width=32)
            else:
                widget = ttk.Entry(self.form_frame, textvariable=var, width=35)
            widget.grid(row=row * 2, column=1, sticky="w", pady=(4, 0))
            hint = ttk.Label(self.form_frame, text="", foreground="#b45309")
            hint.grid(row=row * 2 + 1, column=1, sticky="w")
            self.vars[field.key], self.hints[field.key], self.inputs[field.key] = var, hint, widget

        buttons = ttk.Frame(self.form_frame)
        buttons.grid(row=len(self.form.fields) * 2, column=0, columnspan=2, pady=12, sticky="e")
        self.skip_button = ttk.Button(buttons, text="Skip", command=self.skip, state="disabled")
        self.skip_button.pack(side="right")
        self.save_button = ttk.Button(buttons, text="Save", command=self.save, state="disabled")
        self.save_button.pack(side="right", padx=8)

    def change_form(self, _event=None):
        chosen = self.forms[self.form_var.get()]
        if chosen.key == self.form.key:
            return
        if len(self.batch) - self.index:
            messagebox.showinfo("Finish this batch first", "Save or skip the remaining images before changing form type.")
            self.form_var.set(self.form.name)
            return
        self.form = chosen
        self.build_form()
        self.status.configure(text=f"{self.store.count(chosen.key)} {chosen.name} records.")

    # --- batch handling -------------------------------------------------

    @property
    def current(self):
        return self.batch[self.index] if self.index < len(self.batch) else None

    def open_images(self):
        paths = filedialog.askopenfilenames(filetypes=[("Images", "*.jpg *.jpeg *.png *.webp")])
        if not paths:
            return
        remaining = len(self.batch) - self.index
        if remaining and not messagebox.askyesno(
            "Replace batch?", f"{remaining} image(s) in the current batch haven't been saved. Discard them?"
        ):
            return
        for future in self.futures:
            future.cancel()
        self.batch, self.index, self.outcomes = list(dict.fromkeys(paths)), 0, {}
        self.saved = self.skipped = 0
        self.futures = [self.pool.submit(self._extract, path, self.form) for path in self.batch]
        self.show_current()

    def _extract(self, path, form):
        try:
            self.results.put((path, extract_fields(path, form), None))
        except (ExtractionError, OSError) as e:
            self.results.put((path, None, str(e)))

    def _poll_results(self):
        while True:
            try:
                path, result, error = self.results.get_nowait()
            except queue.Empty:
                break
            if path not in self.batch:
                continue  # left over from a batch the user replaced
            self.outcomes[path] = (result, error)
            if path == self.current:
                self.show_outcome()
        self.update_progress()
        self.root.after(100, self._poll_results)

    def advance(self):
        self.index += 1
        if self.current:
            self.show_current()
            return
        self.clear()
        self.status.configure(
            text=f"Batch done: {self.saved} saved, {self.skipped} skipped. {self.store.count(self.form.key)} "
                 f"{self.form.name} records total."
        )
        self.batch, self.index = [], 0
        self.update_progress()

    def update_progress(self):
        if not self.batch:
            self.progress.configure(text="")
            return
        read = sum(1 for path in self.batch if path in self.outcomes)
        self.progress.configure(text=f"Image {self.index + 1} of {len(self.batch)} · {read}/{len(self.batch)} read")

    # --- current image ---------------------------------------------------

    def show_current(self):
        self.clear_form()
        path = self.current
        try:
            with Image.open(path) as img:
                img = ImageOps.exif_transpose(img)
                img.thumbnail(PREVIEW_SIZE)
                self.preview = ImageTk.PhotoImage(img)
            self.image_label.configure(image=self.preview, text="")
        except OSError:
            self.preview = None
            self.image_label.configure(image="", text=f"Can't display {os.path.basename(path)}")
        self.skip_button.configure(state="normal")
        if path in self.outcomes:
            self.show_outcome()
        else:
            self.set_form_enabled(False)
            self.status.configure(text=f"Reading {os.path.basename(path)}…")
        self.update_progress()

    def show_outcome(self):
        result, error = self.outcomes[self.current]
        self.set_form_enabled(True)
        name = os.path.basename(self.current)
        if error:
            self.status.configure(text=f"Couldn't read {name}: {error} Enter the fields manually or skip.")
            return
        values, uncertain = result
        record = normalize(self.form, values)
        for key in self.form.keys:
            self.vars[key].set(record[key] or "")
        self.uncertain = set(uncertain)
        self.status.configure(text=f"{name}: check every field against the photo, then save.")
        self.show_hints()

    def set_form_enabled(self, enabled):
        for widget in self.inputs.values():
            widget.configure(state="normal" if enabled else "disabled")
        self.save_button.configure(state="normal" if enabled else "disabled")
        self.form_picker.configure(state="disabled" if self.batch else "readonly")

    def current_record(self):
        return normalize(self.form, {key: var.get() for key, var in self.vars.items()})

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
        problems = validate(self.form, record)
        self.show_hints(problems)
        if problems:
            self.status.configure(text="Fix the highlighted fields before saving.")
            return
        if self.uncertain:
            names = ", ".join(f.label for f in self.form.fields if f.key in self.uncertain)
            if not messagebox.askyesno("Confirm", f"These fields were hard to read: {names}.\n\nHave you checked them against the photo?"):
                return
        duplicate = self.store.find_duplicate(self.form, record)
        if duplicate:
            summary = ", ".join(f"{self.form.field(k).label}: {duplicate['values'].get(k)}" for k in self.form.duplicate_keys)
            if not messagebox.askyesno("Possible duplicate",
                                       f"A matching record already exists ({summary}).\n\nSave this one anyway?"):
                return
        self.store.add_entry(self.form, record, ACTOR, source_file=self.current)
        self.saved += 1
        self.advance()

    def skip(self):
        self.skipped += 1
        self.advance()

    def export(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=f"{self.form.name.replace(' ', '')}.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        try:
            count = self.store.export_xlsx(path, self.form)
            self.store.log(ACTOR, "records.export", detail=f"{count} rows; {self.form.key}")
        except OSError as e:
            messagebox.showerror("Export failed", f"Could not write {path}: {e}\nIs it open in Excel?")
            return
        self.status.configure(text=f"Exported {count} {self.form.name} records to {path}.")

    def clear_form(self):
        self.uncertain = set()
        self.set_form_enabled(True)
        for var in self.vars.values():
            var.set("")
        self.show_hints()
        self.save_button.configure(state="disabled")

    def clear(self):
        self.preview = None
        self.image_label.configure(image="", text="No image loaded")
        self.clear_form()
        self.skip_button.configure(state="disabled")

    def close(self):
        remaining = len(self.batch) - self.index
        if remaining and not messagebox.askyesno("Quit?", f"{remaining} image(s) haven't been saved. Quit anyway?"):
            return
        self.pool.shutdown(wait=False, cancel_futures=True)
        self.store.close()
        self.root.destroy()


def open_store() -> tuple[Database, str | None]:
    """Open the database, importing the CSV log from earlier versions on first run."""
    store = Database(DB_FILE, Cipher.from_env())
    if store.count() == 0 and os.path.exists(LEGACY_CSV):
        imported, skipped = store.import_csv(LEGACY_CSV, ACTOR)
        message = f"Imported {imported} records from {LEGACY_CSV}"
        return store, message + (f" ({skipped} incomplete rows skipped)." if skipped else ".")
    return store, None


def main():
    try:
        store, message = open_store()
    except ConfigError as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Can't open records", str(e))
        return
    root = tk.Tk()
    app = ReviewApp(root, store)
    if message:
        app.status.configure(text=message)
    root.mainloop()


if __name__ == "__main__":
    main()
