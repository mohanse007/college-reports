"""
College Report Generator - Desktop GUI Application
Provides a user-friendly interface to select raw SIS/ERP Excel exports,
configure dynamic semester/batch metadata, maintain faculty contact directory,
and generate production-ready college reports.
"""
import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, Any

from ..reports import get_available_reports
from ..core.faculty_manager import FacultyManager, DEFAULT_DIRECTORY_FILENAME


class ReportGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("College Report Generator - St. Ann's College for Women")
        self.geometry("820x680")
        self.minsize(760, 620)

        # Set app background & styling
        self.configure(bg="#F4F6F9")
        self._setup_styles()

        # Load available reports from registry
        self.reports_map = get_available_reports()
        self.selected_report_id = tk.StringVar(
            value=list(self.reports_map.keys())[0] if self.reports_map else ""
        )

        # Variables
        self.raw_file_path = tk.StringVar()
        self.college_name_var = tk.StringVar(value="St.Ann's College for Women (A)")
        self.semester_title_var = tk.StringVar(value="4th Semester Course Code, Course name with faculty Details")
        self.admitted_batch_var = tk.StringVar(value="Admitted Batch 2024-2025")
        self.status_var = tk.StringVar(value="Ready. Select a raw Excel file to begin.")
        self.last_generated_file = None

        # Build UI layout
        self._build_header()
        self._build_main_content()
        self._build_footer()

    def _setup_styles(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        # Configure colors and styles
        self.style.configure(".", font=("Segoe UI", 10), background="#F4F6F9")
        self.style.configure("TLabelframe", background="#FFFFFF", relief="solid", borderwidth=1)
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#1F497D", background="#FFFFFF")
        
        self.style.configure("TButton", font=("Segoe UI", 10), padding=6)
        self.style.configure("Accent.TButton", font=("Segoe UI", 11, "bold"), background="#1F497D", foreground="#FFFFFF")
        self.style.map("Accent.TButton",
            background=[("active", "#16365C"), ("disabled", "#CCCCCC")],
            foreground=[("disabled", "#666666")]
        )
        self.style.configure("Secondary.TButton", font=("Segoe UI", 9), padding=5)

    def _build_header(self):
        header_frame = tk.Frame(self, bg="#1F497D", height=80)
        header_frame.pack(fill="x", side="top")

        title_lbl = tk.Label(
            header_frame,
            text="College Report Generator",
            font=("Segoe UI", 16, "bold"),
            fg="#FFFFFF",
            bg="#1F497D"
        )
        title_lbl.pack(anchor="w", padx=20, pady=(12, 2))

        subtitle_lbl = tk.Label(
            header_frame,
            text="Automated SIS / ERP Excel to College Standard Report Converter",
            font=("Segoe UI", 9),
            fg="#D9E2EC",
            bg="#1F497D"
        )
        subtitle_lbl.pack(anchor="w", padx=20, pady=(0, 12))

    def _build_main_content(self):
        container = tk.Frame(self, bg="#F4F6F9", padx=18, pady=12)
        container.pack(fill="both", expand=True)

        # 1. Report Selector Section (Extensible for future reports)
        rep_frame = ttk.LabelFrame(container, text="  1. Report Type  ", padding=12)
        rep_frame.pack(fill="x", pady=(0, 10))

        r_top = tk.Frame(rep_frame, bg="#FFFFFF")
        r_top.pack(fill="x")

        tk.Label(r_top, text="Select Report:", font=("Segoe UI", 10, "bold"), bg="#FFFFFF").pack(side="left", padx=(0, 10))

        report_options = [r.display_name for r in self.reports_map.values()]
        self.report_combo = ttk.Combobox(
            r_top,
            values=report_options,
            state="readonly",
            font=("Segoe UI", 10),
            width=45
        )
        if report_options:
            self.report_combo.current(0)
        self.report_combo.pack(side="left", fill="x", expand=True)
        self.report_combo.bind("<<ComboboxSelected>>", self._on_report_changed)

        self.rep_desc_lbl = tk.Label(
            rep_frame,
            text=list(self.reports_map.values())[0].description if self.reports_map else "",
            font=("Segoe UI", 9, "italic"),
            fg="#555555",
            bg="#FFFFFF",
            wraplength=700,
            justify="left"
        )
        self.rep_desc_lbl.pack(anchor="w", pady=(6, 0))

        # 2. Raw File Input Section
        file_frame = ttk.LabelFrame(container, text="  2. Raw Input Data  ", padding=12)
        file_frame.pack(fill="x", pady=(0, 10))

        f_box = tk.Frame(file_frame, bg="#FFFFFF")
        f_box.pack(fill="x")

        self.file_entry = ttk.Entry(f_box, textvariable=self.raw_file_path, font=("Segoe UI", 10))
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        browse_btn = ttk.Button(f_box, text="Browse File...", command=self._browse_file)
        browse_btn.pack(side="left")

        # 3. Dynamic Metadata Configuration Section
        meta_frame = ttk.LabelFrame(container, text="  3. Report Headers & Semester Settings  ", padding=12)
        meta_frame.pack(fill="x", pady=(0, 10))

        meta_grid = tk.Frame(meta_frame, bg="#FFFFFF")
        meta_grid.pack(fill="x")

        # College Name
        tk.Label(meta_grid, text="College Name:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", anchor="w").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(meta_grid, textvariable=self.college_name_var, width=58).grid(row=0, column=1, sticky="we", padx=(10, 0), pady=4)

        # Semester Title
        tk.Label(meta_grid, text="Semester / Title:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", anchor="w").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(meta_grid, textvariable=self.semester_title_var, width=58).grid(row=1, column=1, sticky="we", padx=(10, 0), pady=4)

        # Admitted Batch
        tk.Label(meta_grid, text="Admitted Batch:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", anchor="w").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(meta_grid, textvariable=self.admitted_batch_var, width=58).grid(row=2, column=1, sticky="we", padx=(10, 0), pady=4)

        meta_grid.columnconfigure(1, weight=1)

        meta_note = tk.Label(
            meta_frame,
            text="* These values are automatically detected from the selected raw file and can be freely adjusted for any semester/batch.",
            font=("Segoe UI", 8, "italic"),
            fg="#777777",
            bg="#FFFFFF"
        )
        meta_note.pack(anchor="w", pady=(4, 0))

        # 4. Action & Tool Buttons
        action_frame = tk.Frame(container, bg="#F4F6F9")
        action_frame.pack(fill="x", pady=(5, 5))

        self.gen_btn = ttk.Button(
            action_frame,
            text="🚀  Generate Report",
            style="Accent.TButton",
            command=self._start_generation
        )
        self.gen_btn.pack(side="left", padx=(0, 15))

        dir_btn = ttk.Button(
            action_frame,
            text="📋 Manage Faculty Phone Directory",
            style="Secondary.TButton",
            command=self._open_faculty_directory_dialog
        )
        dir_btn.pack(side="left", padx=(0, 10))

        open_folder_btn = ttk.Button(
            action_frame,
            text="📂 Open Output Folder",
            style="Secondary.TButton",
            command=self._open_output_folder
        )
        open_folder_btn.pack(side="left")

        # 5. Output Card / Log
        log_frame = ttk.LabelFrame(container, text="  Status & Output  ", padding=10)
        log_frame.pack(fill="both", expand=True, pady=(10, 0))

        self.progress_bar = ttk.Progressbar(log_frame, mode="indeterminate")
        self.progress_bar.pack(fill="x", pady=(0, 8))

        self.status_lbl = tk.Label(
            log_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            fg="#2D3748",
            bg="#FFFFFF",
            anchor="w",
            justify="left"
        )
        self.status_lbl.pack(fill="x")

        self.open_file_btn = tk.Button(
            log_frame,
            text="📄 Click here to open generated Excel report",
            font=("Segoe UI", 9, "underline bold"),
            fg="#1F497D",
            bg="#FFFFFF",
            bd=0,
            cursor="hand2",
            command=self._open_last_generated_file
        )
        # Hidden until generation completes

    def _build_footer(self):
        footer_frame = tk.Frame(self, bg="#E2E8F0", height=24)
        footer_frame.pack(fill="x", side="bottom")

        ftr_lbl = tk.Label(
            footer_frame,
            text="St. Ann's College for Women (Autonomous) | Modular Report System",
            font=("Segoe UI", 8),
            fg="#4A5568",
            bg="#E2E8F0"
        )
        ftr_lbl.pack(side="left", padx=15, pady=3)

    def _on_report_changed(self, event=None):
        sel_name = self.report_combo.get()
        for r_id, r_obj in self.reports_map.items():
            if r_obj.display_name == sel_name:
                self.selected_report_id.set(r_id)
                self.rep_desc_lbl.config(text=r_obj.description)
                break

    def _browse_file(self):
        init_dir = os.path.dirname(os.path.abspath(self.raw_file_path.get() or "."))
        fpath = filedialog.askopenfilename(
            title="Select Raw Course Allocation Report",
            initialdir=init_dir,
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if fpath:
            self.raw_file_path.set(fpath)
            self._auto_detect_metadata(fpath)

    def _auto_detect_metadata(self, fpath: str):
        """Auto-detects semester and batch year when a file is chosen."""
        active_report = self.reports_map.get(self.selected_report_id.get())
        if active_report:
            try:
                meta = active_report.extract_metadata(fpath)
                if meta.get("college_name"):
                    self.college_name_var.set(meta["college_name"])
                if meta.get("semester_title"):
                    self.semester_title_var.set(meta["semester_title"])
                if meta.get("admitted_batch"):
                    self.admitted_batch_var.set(meta["admitted_batch"])
                self.status_var.set(f"File loaded. Auto-detected: {meta.get('semester_title')}")
            except Exception as e:
                self.status_var.set(f"Loaded file, but could not auto-detect headers: {e}")

    def _start_generation(self):
        fpath = self.raw_file_path.get().strip()
        if not fpath or not os.path.exists(fpath):
            messagebox.showerror("Missing File", "Please select a valid raw data Excel file first.")
            return

        self.gen_btn.config(state="disabled")
        self.progress_bar.start(10)
        self.open_file_btn.pack_forget()
        self.status_var.set("Processing raw data and generating report...")

        # Run in thread so GUI doesn't freeze
        t = threading.Thread(target=self._run_generation_thread, args=(fpath,))
        t.daemon = True
        t.start()

    def _run_generation_thread(self, fpath: str):
        try:
            active_report = self.reports_map.get(self.selected_report_id.get())
            if not active_report:
                raise ValueError("Selected report type is not recognized.")

            output_dir = os.path.join(os.path.dirname(fpath), "Generated_Reports")
            options = {
                "college_name": self.college_name_var.get().strip(),
                "semester_title": self.semester_title_var.get().strip(),
                "admitted_batch": self.admitted_batch_var.get().strip(),
            }

            out_path = active_report.generate(fpath, output_dir, options)
            self.last_generated_file = out_path

            self.after(0, self._on_generation_success, out_path)
        except Exception as e:
            self.after(0, self._on_generation_error, str(e))

    def _on_generation_success(self, out_path: str):
        self.progress_bar.stop()
        self.gen_btn.config(state="normal")
        self.status_var.set(f"✅ Report successfully generated!\nSaved at: {out_path}")
        self.open_file_btn.pack(anchor="w", pady=(5, 0))

    def _on_generation_error(self, err_msg: str):
        self.progress_bar.stop()
        self.gen_btn.config(state="normal")
        self.status_var.set(f"❌ Error generating report: {err_msg}")
        messagebox.showerror("Generation Error", f"Failed to generate report:\n{err_msg}")

    def _open_last_generated_file(self):
        if self.last_generated_file and os.path.exists(self.last_generated_file):
            try:
                os.startfile(self.last_generated_file)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file: {e}")

    def _open_output_folder(self):
        raw_p = self.raw_file_path.get().strip()
        base_dir = os.path.dirname(raw_p) if raw_p else os.getcwd()
        out_dir = os.path.join(base_dir, "Generated_Reports")
        os.makedirs(out_dir, exist_ok=True)
        try:
            os.startfile(out_dir)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open output folder: {e}")

    def _open_faculty_directory_dialog(self):
        FacultyDirectoryDialog(self)


class FacultyDirectoryDialog(tk.Toplevel):
    """Interactive dialog to view and update faculty phone numbers."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Faculty Phone Directory")
        self.geometry("700x520")
        self.transient(parent)
        self.grab_set()

        self.mgr = FacultyManager()
        self._build_ui()

    def _build_ui(self):
        top_frame = tk.Frame(self, padx=12, pady=8, bg="#F4F6F9")
        top_frame.pack(fill="x")

        tk.Label(
            top_frame,
            text="Faculty Contact Directory (Used to automatically fill Phone Numbers in reports)",
            font=("Segoe UI", 9, "bold"),
            bg="#F4F6F9"
        ).pack(anchor="w", pady=(0, 4))

        # Search bar
        search_box = tk.Frame(top_frame, bg="#F4F6F9")
        search_box.pack(fill="x")
        tk.Label(search_box, text="Search Name:", bg="#F4F6F9").pack(side="left", padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._filter_table)
        ttk.Entry(search_box, textvariable=self.search_var, width=30).pack(side="left")

        open_excel_btn = ttk.Button(search_box, text="Open in Excel", command=self._open_in_excel)
        open_excel_btn.pack(side="right")

        import_btn = ttk.Button(search_box, text="📥 Import Staff List Report", command=self._import_staff_list)
        import_btn.pack(side="right", padx=(0, 8))

        # Treeview Table
        table_frame = tk.Frame(self, padx=12, pady=4)
        table_frame.pack(fill="both", expand=True)

        cols = ("Name", "Phone", "Notes")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("Name", text="Faculty Name")
        self.tree.heading("Phone", text="Phone Number")
        self.tree.heading("Notes", text="Department / Notes")

        self.tree.column("Name", width=260)
        self.tree.column("Phone", width=140)
        self.tree.column("Notes", width=220)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self._edit_selected)

        # Edit fields below
        edit_box = tk.LabelFrame(self, text=" Edit Contact Details ", padx=10, pady=8)
        edit_box.pack(fill="x", padx=12, pady=8)

        grid = tk.Frame(edit_box)
        grid.pack(fill="x")

        tk.Label(grid, text="Name:").grid(row=0, column=0, sticky="w")
        self.edit_name = ttk.Entry(grid, width=28)
        self.edit_name.grid(row=0, column=1, padx=5, pady=2, sticky="w")

        tk.Label(grid, text="Phone:").grid(row=0, column=2, sticky="w", padx=(10, 0))
        self.edit_phone = ttk.Entry(grid, width=18)
        self.edit_phone.grid(row=0, column=3, padx=5, pady=2, sticky="w")

        save_item_btn = ttk.Button(grid, text="Update / Add", command=self._save_item)
        save_item_btn.grid(row=0, column=4, padx=(10, 0))

        self._refresh_table()

    def _refresh_table(self):
        self.all_records = self.mgr.get_all()
        self._filter_table()

    def _filter_table(self, *args):
        query = self.search_var.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        for r in self.all_records:
            name = r.get("name", "")
            phone = r.get("phone", "")
            notes = r.get("notes", "")
            if query in name.lower() or query in phone.lower():
                self.tree.insert("", "end", values=(name, phone, notes))

    def _edit_selected(self, event=None):
        sel = self.tree.selection()
        if sel:
            vals = self.tree.item(sel[0], "values")
            self.edit_name.delete(0, "end")
            self.edit_name.insert(0, vals[0])
            self.edit_phone.delete(0, "end")
            self.edit_phone.insert(0, vals[1])

    def _save_item(self):
        name = self.edit_name.get().strip()
        phone = self.edit_phone.get().strip()
        if not name:
            messagebox.showwarning("Missing Name", "Faculty name cannot be empty.")
            return

        # Update in all_records
        found = False
        for r in self.all_records:
            if r["name"].strip().lower() == name.lower():
                r["phone"] = phone
                found = True
                break
        if not found:
            self.all_records.append({"name": name, "phone": phone, "notes": ""})

        self.mgr.save_all(self.all_records)
        self._refresh_table()
        messagebox.showinfo("Saved", f"Contact details for {name} updated.")

    def _open_in_excel(self):
        if os.path.exists(self.mgr.directory_path):
            os.startfile(self.mgr.directory_path)

    def _import_staff_list(self):
        fpath = filedialog.askopenfilename(
            title="Select Staff List Report",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if fpath:
            count = self.mgr.import_from_staff_list_report(fpath)
            self._refresh_table()
            messagebox.showinfo("Import Complete", f"Successfully imported/updated staff contact records from Staff List Report.")
