#!/usr/bin/env python3
"""
Schedule Optimizer - GUI
========================

Run:      python schedule_optimizer_gui.py
Needs:    Python 3.10+ (tkinter ships with it) and, for Excel export, `pip install openpyxl`

Workflow
  1. Schedules  - type in (or paste / import) the weekly meetings of every section.
  2. Options    - choose what matters most (priority order), rank your preferred
                  sections for each year level, set optional hard limits, and
                  untick subjects you are not taking.
  3. Press "Find best schedules" - results appear in the Results tab, with a weekly
                  timetable, and can be exported to Excel / CSV.

How the ranking works
  Every subject part (e.g. "DSA Lec", "DSA Lab", or an ALL-in-one subject) is assigned
  to exactly one section. All combinations are searched (branch-and-bound) and ranked
  by your priority list, e.g.  conflicts > days > Year-3 preference > Year-2 preference.
  Section preference: for a year level with N ranked sections, each subject part placed
  in your #1 section earns N points, #2 earns N-1, ... , unranked sections earn 0.
  More points is better.
"""
from __future__ import annotations

import csv
import json
import os
import re
import threading
import time
import tkinter as tk
from collections import defaultdict
from tkinter import filedialog, messagebox, ttk

from engine import (
    DAY_ORDER, PARTS, Row, SAMPLE, SAMPLE_PRIORITY, SAMPLE_RANKINGS, build_components,
    choice_note, comp_label, describe, export_xlsx, fmt, natural_key, overlap_minutes,
    row_sort_key, search, split_label, summary_line, timetable_lines, to_minutes, year_label,
    year_of, year_sort_key,
)

# ----------------------------------------------------------------------------
# Small reusable widgets
# ----------------------------------------------------------------------------
class ScrollFrame(ttk.Frame):
    """A vertically scrollable frame; put children in `.inner`."""

    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_all(seq, self._on_wheel, add="+")

    def _on_wheel(self, e):
        w = self.winfo_containing(e.x_root, e.y_root)
        if isinstance(w, (tk.Listbox, tk.Text)):
            return
        while w is not None:
            if w is self:
                step = -1 if (e.num == 4 or e.delta > 0) else 1
                self.canvas.yview_scroll(step, "units")
                return
            w = getattr(w, "master", None)


class RankPicker(ttk.LabelFrame):
    """Two lists: sections that are not ranked / sections in preference order."""

    def __init__(self, master, year, sections, ranked, on_change):
        super().__init__(master, text=f" {year_label(year)} - rank the sections you prefer ", padding=8)
        self.year, self.sections, self.on_change = year, list(sections), on_change
        self.ranked = [s for s in ranked if s in self.sections]

        left = ttk.Frame(self)
        left.grid(row=0, column=0, sticky="ns")
        ttk.Label(left, text="Not ranked").pack(anchor="w")
        self.lb_free = tk.Listbox(left, height=7, width=10, exportselection=False, selectmode="extended")
        self.lb_free.pack(fill="y", expand=True)

        mid = ttk.Frame(self)
        mid.grid(row=0, column=1, padx=8)
        ttk.Button(mid, text="Rank  →", width=13, command=self.rank).pack(pady=(16, 3))
        ttk.Button(mid, text="←  Remove", width=13, command=self.unrank).pack(pady=3)
        ttk.Button(mid, text="▲  Move up", width=13, command=lambda: self.move(-1)).pack(pady=(14, 3))
        ttk.Button(mid, text="▼  Move down", width=13, command=lambda: self.move(1)).pack(pady=3)

        right = ttk.Frame(self)
        right.grid(row=0, column=2, sticky="ns")
        ttk.Label(right, text="Ranked (top = best)").pack(anchor="w")
        self.lb_rank = tk.Listbox(right, height=7, width=16, exportselection=False, selectmode="extended")
        self.lb_rank.pack(fill="y", expand=True)


        self.lb_free.bind("<Double-Button-1>", lambda e: self.rank())
        self.lb_rank.bind("<Double-Button-1>", lambda e: self.unrank())
        self._fill()

    def _free(self):
        return [s for s in self.sections if s not in self.ranked]

    def _fill(self, select=()):
        self.lb_free.delete(0, "end")
        for s in self._free():
            self.lb_free.insert("end", s)
        self.lb_rank.delete(0, "end")
        for i, s in enumerate(self.ranked, 1):
            self.lb_rank.insert("end", f"{i}.  {s}")
        for i in select:
            self.lb_rank.selection_set(i)

    def _changed(self):
        self.on_change(self.year, list(self.ranked))

    def rank(self):
        free = self._free()
        for i in self.lb_free.curselection():
            self.ranked.append(free[i])
        self._fill()
        self._changed()

    def unrank(self):
        for i in sorted(self.lb_rank.curselection(), reverse=True):
            del self.ranked[i]
        self._fill()
        self._changed()

    def move(self, d):
        sel = list(self.lb_rank.curselection())
        if not sel or (d < 0 and sel[0] == 0) or (d > 0 and sel[-1] == len(self.ranked) - 1):
            return
        for i in (sel if d < 0 else reversed(sel)):
            self.ranked[i], self.ranked[i + d] = self.ranked[i + d], self.ranked[i]
        self._fill(select=[i + d for i in sel])
        self._changed()


# ----------------------------------------------------------------------------
# Main application
# ----------------------------------------------------------------------------
PALETTE = ["#CFE8FF", "#D8F0C8", "#FFE3B8", "#E6D5F5", "#FFD6E0", "#C9F0EC",
           "#F6F0B0", "#DCE3EA", "#F9D3B4", "#CDE7D0"]
CLASH_COLOR = "#F4A6A6"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Schedule Optimizer")
        self.geometry("1260x820")
        self.minsize(1080, 700)
        style = ttk.Style(self)
        if style.theme_use() in ("default", "classic"):
            style.theme_use("clam")

        self.rows: list[Row] = []
        self.next_uid = 1
        self.excluded: set[tuple[str, str]] = set()
        self.rankings: dict[str, list[str]] = {}
        self.priority: list[str] = ["conflict", "days"]
        self.results: list[dict] = []
        self.res_rankings: dict[str, list[str]] = {}
        self.current: dict | None = None
        self.subject_colors: dict[str, str] = {}
        self.worker: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.search_state: dict = {}
        self.search_out: dict = {}
        self._pref_sig = None
        self._comp_keys: dict[str, tuple[str, str]] = {}
        self._redraw_job = None

        self._build_menu()
        self._build_runbar()
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=(8, 0))
        self._build_schedule_tab()
        self._build_options_tab()
        self._build_results_tab()
        self.load_sample()

    # ------------------------------------------------------------------ menu
    def _build_menu(self):
        bar = tk.Menu(self)
        f = tk.Menu(bar, tearoff=0)
        f.add_command(label="New (empty)", command=self.new_project)
        f.add_command(label="Open project…", command=self.open_project)
        f.add_command(label="Save project…", command=self.save_project)
        f.add_separator()
        f.add_command(label="Load sample data", command=self.load_sample)
        f.add_separator()
        f.add_command(label="Quit", command=self.destroy)
        bar.add_cascade(label="File", menu=f)
        h = tk.Menu(bar, tearoff=0)
        h.add_command(label="How the ranking works", command=self.show_help)
        bar.add_cascade(label="Help", menu=h)
        self.config(menu=bar)

    def show_help(self):
        messagebox.showinfo(
            "How the ranking works",
            "Every subject part (e.g. 'DSA Lec', 'DSA Lab', or a single ALL class) is placed in "
            "exactly one section. Lec and Lab of the same subject can come from different sections.\n\n"
            "Schedules are compared using your priority list, top to bottom: the first item decides, "
            "and later items only break ties.\n\n"
            "Section preference: for a year level with N ranked sections, a subject part in your #1 "
            "section earns N points, #2 earns N-1, … and unranked sections earn 0. Higher is better.\n\n"
            "Hard limits (max days / max conflict hours) remove schedules that break them entirely.\n\n"
            "Tip: 'Fewest days first' + 'Max conflict hours = 0' gives the most compact "
            "conflict-free schedule.")

    # ---------------------------------------------------------------- run bar
    def _build_runbar(self):
        bar = ttk.Frame(self, padding=(8, 8))
        bar.pack(side="bottom", fill="x")
        self.run_btn = ttk.Button(bar, text="▶  Find best schedules", command=self.start_search)
        self.run_btn.pack(side="left")
        self.stop_btn = ttk.Button(bar, text="Stop", command=self.stop_search, state="disabled")
        self.stop_btn.pack(side="left", padx=(6, 12))
        self.progress = ttk.Progressbar(bar, maximum=1000, length=240)
        self.progress.pack(side="left")
        self.status = tk.StringVar(value="Ready.")
        ttk.Label(bar, textvariable=self.status).pack(side="left", padx=12)

    # ----------------------------------------------------------- schedules tab
    def _build_schedule_tab(self):
        tab = ttk.Frame(self.nb, padding=8)
        self.nb.add(tab, text="  1 · Schedules  ")

        form = ttk.LabelFrame(tab, text=" Add or edit a class meeting ", padding=8)
        form.pack(fill="x")
        self.f_section, self.f_day = tk.StringVar(), tk.StringVar(value="Mon")
        self.f_subject, self.f_part = tk.StringVar(), tk.StringVar(value="LEC")
        self.f_start, self.f_end = tk.StringVar(), tk.StringVar()
        for c, text in enumerate(["Section", "Day", "Subject", "Type", "Start", "End"]):
            ttk.Label(form, text=text).grid(row=0, column=c, sticky="w", padx=3)
        self.cb_section = ttk.Combobox(form, textvariable=self.f_section, width=9)
        self.cb_day = ttk.Combobox(form, textvariable=self.f_day, values=DAY_ORDER, width=6, state="readonly")
        self.cb_subject = ttk.Combobox(form, textvariable=self.f_subject, width=20)
        self.cb_part = ttk.Combobox(form, textvariable=self.f_part, values=PARTS, width=6, state="readonly")
        e_start = ttk.Entry(form, textvariable=self.f_start, width=10)
        e_end = ttk.Entry(form, textvariable=self.f_end, width=10)
        for c, w in enumerate([self.cb_section, self.cb_day, self.cb_subject, self.cb_part, e_start, e_end]):
            w.grid(row=1, column=c, padx=3, pady=(0, 2), sticky="w")
        for w in (self.cb_section, self.cb_subject, e_start, e_end):
            w.bind("<Return>", lambda e: self.add_row())
        fb = ttk.Frame(form)
        fb.grid(row=2, column=0, columnspan=6, sticky="w", pady=(6, 0))
        ttk.Button(fb, text="Add", width=8, command=self.add_row).pack(side="left", padx=(3, 6))
        ttk.Button(fb, text="Save changes", command=self.update_row).pack(side="left", padx=3)
        ttk.Button(fb, text="Delete", width=8, command=self.delete_rows).pack(side="left", padx=3)
        ttk.Button(fb, text="Clear form", command=self.clear_form).pack(side="left", padx=3)
        ttk.Label(
            form, foreground="#555555", wraplength=900, justify="left",
            text="Times: 7:00AM, 2 PM or 14:00.   Type: LEC and LAB are chosen separately (they may come from "
                 "different sections); ALL = one combined class.   Click a row below to edit it.",
        ).grid(row=3, column=0, columnspan=6, sticky="w", padx=3, pady=(6, 0))

        top = ttk.Frame(tab)
        top.pack(fill="x", pady=(10, 4))
        ttk.Label(top, text="Show:").pack(side="left")
        self.filter_var = tk.StringVar(value="All sections")
        self.cb_filter = ttk.Combobox(top, textvariable=self.filter_var, state="readonly", width=14,
                                      values=["All sections"])
        self.cb_filter.pack(side="left", padx=6)
        self.cb_filter.bind("<<ComboboxSelected>>", lambda e: self.refresh_tree())
        self.count_var = tk.StringVar()
        ttk.Label(top, textvariable=self.count_var, foreground="#555555").pack(side="right")

        mid = ttk.Frame(tab)
        mid.pack(fill="both", expand=True)
        cols = ("section", "day", "subject", "part", "start", "end")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="extended")
        for c, (text, w) in zip(cols, [("Section", 90), ("Day", 70), ("Subject", 220),
                                       ("Type", 70), ("Start", 90), ("End", 90)]):
            self.tree.heading(c, text=text)
            self.tree.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)
        self.tree.bind("<Delete>", lambda e: self.delete_rows())

        btns = ttk.Frame(tab)
        btns.pack(fill="x", pady=(8, 0))
        for text, cmd in [("Paste rows…", self.paste_rows), ("Import CSV…", self.import_csv),
                          ("Export CSV…", self.export_schedule_csv), ("Load sample data", self.load_sample),
                          ("Clear all", self.clear_all)]:
            ttk.Button(btns, text=text, command=cmd).pack(side="left", padx=(0, 6))

    # --- form helpers
    def clear_form(self):
        self.tree.selection_remove(self.tree.selection())
        self.f_subject.set("")
        self.f_start.set("")
        self.f_end.set("")

    def _read_form(self):
        section, subject = self.f_section.get().strip(), self.f_subject.get().strip()
        day, part = self.f_day.get().strip(), self.f_part.get().strip()
        if not section:
            raise ValueError("Enter a section (for example 3-D).")
        if not subject:
            raise ValueError("Enter a subject.")
        if any(ch in section + subject for ch in "[]"):
            raise ValueError("Section and subject names can't contain [ or ].")
        if day not in DAY_ORDER:
            raise ValueError("Pick a day.")
        if part not in PARTS:
            raise ValueError("Pick a type (LEC, LAB or ALL).")
        start, end = to_minutes(self.f_start.get()), to_minutes(self.f_end.get())
        if end <= start:
            raise ValueError("The end time must be after the start time.")
        return section, day, subject, part, start, end

    def _new_row(self, values):
        row = Row(self.next_uid, *values)
        self.next_uid += 1
        return row

    def _exists(self, values, ignore_uid=None):
        return any(r.uid != ignore_uid and (r.section, r.day, r.subject, r.part, r.start, r.end) == values
                   for r in self.rows)

    def add_row(self):
        try:
            values = self._read_form()
        except ValueError as ex:
            messagebox.showwarning("Can't add", str(ex))
            return
        if self._exists(values):
            messagebox.showinfo("Already there", "That exact meeting is already in the list.")
            return
        row = self._new_row(values)
        self.rows.append(row)
        self.refresh_tree(select=[row.uid])
        self.f_subject.set("")
        self.f_start.set("")
        self.f_end.set("")
        self.cb_subject.focus_set()

    def update_row(self):
        sel = self.tree.selection()
        if len(sel) != 1:
            messagebox.showinfo("Select one row", "Click a single row in the table first, edit the form, "
                                                  "then press Save changes.")
            return
        try:
            values = self._read_form()
        except ValueError as ex:
            messagebox.showwarning("Can't save", str(ex))
            return
        uid = int(sel[0])
        if self._exists(values, ignore_uid=uid):
            messagebox.showinfo("Already there", "That exact meeting is already in the list.")
            return
        for r in self.rows:
            if r.uid == uid:
                r.section, r.day, r.subject, r.part, r.start, r.end = values
        self.refresh_tree(select=[uid])

    def delete_rows(self):
        sel = {int(i) for i in self.tree.selection()}
        if not sel:
            return
        self.rows = [r for r in self.rows if r.uid not in sel]
        self.refresh_tree()

    def on_row_select(self, _event=None):
        sel = self.tree.selection()
        if len(sel) != 1:
            return
        uid = int(sel[0])
        for r in self.rows:
            if r.uid == uid:
                self.f_section.set(r.section)
                self.f_day.set(r.day)
                self.f_subject.set(r.subject)
                self.f_part.set(r.part)
                self.f_start.set(fmt(r.start))
                self.f_end.set(fmt(r.end))

    def refresh_tree(self, select=None):
        self.rows.sort(key=row_sort_key)
        sections = sorted({r.section for r in self.rows}, key=natural_key)
        subjects = sorted({r.subject for r in self.rows}, key=str.lower)
        self.cb_section["values"] = sections
        self.cb_subject["values"] = subjects
        self.cb_filter["values"] = ["All sections"] + sections
        if self.filter_var.get() not in self.cb_filter["values"]:
            self.filter_var.set("All sections")
        flt = self.filter_var.get()
        self.tree.delete(*self.tree.get_children())
        for r in self.rows:
            if flt != "All sections" and r.section != flt:
                continue
            self.tree.insert("", "end", iid=str(r.uid),
                             values=(r.section, r.day, r.subject, r.part, fmt(r.start), fmt(r.end)))
        if select:
            keep = [str(u) for u in select if self.tree.exists(str(u))]
            if keep:
                self.tree.selection_set(keep)
                self.tree.see(keep[0])
        self.count_var.set(f"{len(self.rows)} meetings · {len(sections)} sections")
        self.on_data_changed()

    # --- bulk input
    def paste_rows(self):
        win = tk.Toplevel(self)
        win.title("Paste rows")
        win.transient(self)
        win.geometry("620x420")
        ttk.Label(win, padding=8, justify="left",
                  text="One class meeting per line:   Day, Subject, Type, Start, End\n"
                       "(commas, tabs or | all work, so you can paste straight from Excel).\n"
                       "If a line has 6 values, the first one is taken as the section.").pack(anchor="w")
        top = ttk.Frame(win, padding=(8, 0))
        top.pack(fill="x")
        ttk.Label(top, text="Section for lines without one:").pack(side="left")
        sec_var = tk.StringVar(value=self.f_section.get())
        ttk.Combobox(top, textvariable=sec_var, values=list(self.cb_section["values"]), width=10).pack(
            side="left", padx=6)
        txt = tk.Text(win, height=14, wrap="none", font="TkFixedFont")
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        txt.insert("1.0", "Mon, DS and A, LEC, 3:30PM, 5:30PM\nMon, AdvNet, LAB, 6:00PM, 9:00PM\n")

        def do_add():
            section_default = sec_var.get().strip()
            new_rows, errors = [], []
            for ln, line in enumerate(txt.get("1.0", "end").splitlines(), 1):
                if not line.strip():
                    continue
                parts = [p.strip() for p in re.split(r"\t|\||,", line)]
                parts = [p for p in parts if p != ""]
                try:
                    if len(parts) == 6:
                        section, parts = parts[0], parts[1:]
                    elif len(parts) == 5:
                        section = section_default
                    else:
                        raise ValueError("expected 5 values (or 6 with the section first)")
                    if not section:
                        raise ValueError("no section given")
                    day, subject, part, s, e = parts
                    day = day.title()[:3]
                    part = part.upper()
                    if day not in DAY_ORDER:
                        raise ValueError(f"unknown day '{parts[0]}'")
                    if part not in PARTS:
                        raise ValueError(f"type must be LEC, LAB or ALL (got '{parts[2]}')")
                    if any(ch in section + subject for ch in "[]"):
                        raise ValueError("names can't contain [ or ]")
                    start, end = to_minutes(s), to_minutes(e)
                    if end <= start:
                        raise ValueError("end must be after start")
                    new_rows.append((section, day, subject, part, start, end))
                except ValueError as ex:
                    errors.append(f"Line {ln}: {ex}")
            if errors:
                messagebox.showwarning("Fix these lines", "\n".join(errors[:12]), parent=win)
                return
            added = 0
            for values in new_rows:
                if not self._exists(values):
                    self.rows.append(self._new_row(values))
                    added += 1
            self.refresh_tree()
            win.destroy()
            self.status.set(f"Added {added} meeting(s).")

        btns = ttk.Frame(win, padding=8)
        btns.pack(fill="x")
        ttk.Button(btns, text="Add rows", command=do_add).pack(side="right")
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="right", padx=6)

    def import_csv(self):
        path = filedialog.askopenfilename(title="Import schedules from CSV",
                                          filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        added, skipped = 0, []
        try:
            with open(path, newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for ln, rec in enumerate(reader, 2):
                    rec = {(k or "").strip().lower(): (v or "").strip() for k, v in rec.items()}
                    try:
                        section = rec.get("section", "")
                        subject = rec.get("subject", "")
                        part = (rec.get("type") or rec.get("part") or "").upper()
                        day = rec.get("day", "").title()[:3]
                        start, end = to_minutes(rec.get("start", "")), to_minutes(rec.get("end", ""))
                        if not section or not subject or day not in DAY_ORDER or part not in PARTS or end <= start:
                            raise ValueError("bad row")
                        if any(ch in section + subject for ch in "[]"):
                            raise ValueError("bad name")
                        values = (section, day, subject, part, start, end)
                        if not self._exists(values):
                            self.rows.append(self._new_row(values))
                            added += 1
                    except ValueError:
                        skipped.append(ln)
        except OSError as ex:
            messagebox.showerror("Import failed", str(ex))
            return
        self.refresh_tree()
        msg = f"Imported {added} meeting(s)."
        if skipped:
            msg += f"\nSkipped {len(skipped)} unreadable row(s) (file lines: {', '.join(map(str, skipped[:15]))})."
        msg += "\n\nExpected columns: section, day, subject, type, start, end"
        messagebox.showinfo("CSV import", msg)

    def export_schedule_csv(self):
        path = filedialog.asksaveasfilename(title="Export schedules to CSV", defaultextension=".csv",
                                            initialfile="class_schedules.csv",
                                            filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["section", "day", "subject", "type", "start", "end"])
            for r in sorted(self.rows, key=row_sort_key):
                w.writerow([r.section, r.day, r.subject, r.part, fmt(r.start), fmt(r.end)])
        self.status.set(f"Saved {len(self.rows)} meetings to {os.path.basename(path)}")

    def clear_all(self):
        if self.rows and not messagebox.askyesno("Clear all", "Remove every class meeting from the list?"):
            return
        self.rows.clear()
        self.excluded.clear()
        self.rankings.clear()
        self.priority = ["conflict", "days"]
        self._pref_sig = None
        self.refresh_tree()

    def new_project(self):
        self.clear_all()

    def load_sample(self):
        if self.rows and not messagebox.askyesno(
                "Load sample data", "Replace the current schedules and preferences with the sample data?"):
            return
        self.rows = []
        for section, meetings in SAMPLE.items():
            for day, subject, part, s, e in meetings:
                self.rows.append(self._new_row((section, day, subject, part, to_minutes(s), to_minutes(e))))
        self.excluded.clear()
        self.rankings = {y: list(v) for y, v in SAMPLE_RANKINGS.items()}
        self.priority = list(SAMPLE_PRIORITY)
        self.max_days_var.set("No limit")
        self.max_conf_var.set("")
        self._pref_sig = None
        self.refresh_tree()
        self.status.set("Sample data loaded. Press 'Find best schedules'.")

    # ------------------------------------------------------------ project file
    def save_project(self):
        path = filedialog.asksaveasfilename(title="Save project", defaultextension=".json",
                                            initialfile="schedule_project.json",
                                            filetypes=[("Project files", "*.json")])
        if not path:
            return
        data = {
            "version": 1,
            "meetings": [{"section": r.section, "day": r.day, "subject": r.subject, "type": r.part,
                          "start": fmt(r.start), "end": fmt(r.end)} for r in sorted(self.rows, key=row_sort_key)],
            "rankings": self.rankings, "priority": self.priority,
            "excluded": sorted(list(k) for k in self.excluded),
            "max_days": self.max_days_var.get(), "max_conflict_hours": self.max_conf_var.get(),
            "keep": self.top_n_var.get(),
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        self.status.set(f"Project saved to {os.path.basename(path)}")

    def open_project(self):
        path = filedialog.askopenfilename(title="Open project", filetypes=[("Project files", "*.json")])
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            rows = []
            for m in data["meetings"]:
                values = (m["section"], m["day"], m["subject"], m["type"],
                          to_minutes(m["start"]), to_minutes(m["end"]))
                rows.append(self._new_row(values))
            self.rows = rows
            self.excluded = {tuple(k) for k in data.get("excluded", [])}
            self.rankings = {str(y): list(v) for y, v in data.get("rankings", {}).items()}
            self.priority = list(data.get("priority", ["conflict", "days"]))
            self.max_days_var.set(data.get("max_days", "No limit"))
            self.max_conf_var.set(data.get("max_conflict_hours", ""))
            self.top_n_var.set(int(data.get("keep", 15)))
        except (OSError, ValueError, KeyError, TypeError) as ex:
            messagebox.showerror("Can't open project", f"That file couldn't be read:\n{ex}")
            return
        self._pref_sig = None
        self.refresh_tree()
        self.status.set(f"Opened {os.path.basename(path)}")

    # ------------------------------------------------------------ options tab
    def _build_options_tab(self):
        tab = ttk.Frame(self.nb, padding=8)
        self.nb.add(tab, text="  2 · Options  ")
        tab.columnconfigure(2, weight=1)
        tab.rowconfigure(0, weight=1)

        # -- left column: priority + limits
        left = ttk.Frame(tab)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        pri = ttk.LabelFrame(left, text=" What matters most?  (top = most important) ", padding=8)
        pri.pack(fill="x")
        row = ttk.Frame(pri)
        row.pack(fill="x")
        self.pri_list = tk.Listbox(row, height=5, width=31, exportselection=False, activestyle="none")
        self.pri_list.pack(side="left")
        ud = ttk.Frame(row)
        ud.pack(side="left", padx=6, anchor="n")
        ttk.Button(ud, text="▲", width=4, command=lambda: self.move_priority(-1)).pack(pady=(0, 4))
        ttk.Button(ud, text="▼", width=4, command=lambda: self.move_priority(1)).pack()
        ttk.Label(pri, text="Quick presets:").pack(anchor="w", pady=(10, 2))
        ttk.Button(pri, text="Fewest conflicts → fewest days → preferences",
                   command=lambda: self.preset("default")).pack(fill="x", pady=1)
        ttk.Button(pri, text="Fewest DAYS first (then conflicts, preferences)",
                   command=lambda: self.preset("days")).pack(fill="x", pady=1)
        ttk.Button(pri, text="Section preferences first",
                   command=lambda: self.preset("prefs")).pack(fill="x", pady=1)
        ttk.Label(pri, foreground="#555555", wraplength=300, justify="left",
                  text="A lower item only breaks ties between schedules that are equal on every "
                       "item above it.").pack(anchor="w", pady=(8, 0))

        lim = ttk.LabelFrame(left, text=" Hard limits & output ", padding=8)
        lim.pack(fill="x", pady=(10, 0))
        self.max_days_var = tk.StringVar(value="No limit")
        self.max_conf_var = tk.StringVar(value="")
        self.top_n_var = tk.IntVar(value=15)
        ttk.Label(lim, text="Max days at school").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Combobox(lim, textvariable=self.max_days_var, state="readonly", width=9,
                     values=["No limit"] + [str(i) for i in range(1, len(DAY_ORDER) + 1)]
                     ).grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(lim, text="Max conflict hours").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(lim, textvariable=self.max_conf_var, width=11).grid(row=1, column=1, sticky="w", padx=8)
        ttk.Label(lim, text="blank = no limit,  0 = no conflicts allowed", foreground="#555555"
                  ).grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Label(lim, text="Schedules to keep").grid(row=3, column=0, sticky="w", pady=(10, 3))
        ttk.Spinbox(lim, from_=1, to=200, textvariable=self.top_n_var, width=9
                    ).grid(row=3, column=1, sticky="w", padx=8, pady=(10, 3))

        # -- middle column: subjects to include
        mid = ttk.LabelFrame(tab, text=" Subjects to schedule (click to tick / untick) ", padding=8)
        mid.grid(row=0, column=1, sticky="ns", padx=(0, 10))
        cols = ("use", "subject", "part", "n")
        self.sub_tree = ttk.Treeview(mid, columns=cols, show="headings", height=16, selectmode="none")
        for c, (text, w, anchor) in zip(cols, [("", 34, "center"), ("Subject", 140, "w"),
                                               ("Type", 50, "w"), ("Sections", 74, "center")]):
            self.sub_tree.heading(c, text=text)
            self.sub_tree.column(c, width=w, anchor=anchor, stretch=False)
        self.sub_tree.pack(fill="y", expand=True)
        self.sub_tree.bind("<Button-1>", self.toggle_component)
        b = ttk.Frame(mid)
        b.pack(fill="x", pady=(6, 0))
        ttk.Button(b, text="Select all", command=lambda: self.set_all_components(True)).pack(side="left")
        ttk.Button(b, text="Select none", command=lambda: self.set_all_components(False)).pack(side="left", padx=6)

        # -- right column: section preferences per year level
        right = ttk.LabelFrame(tab, text=" Section preferences by year level ", padding=8)
        right.grid(row=0, column=2, sticky="nsew")
        ttk.Label(right, foreground="#555555", wraplength=400, justify="left",
                  text="Rank the sections you'd rather be in; unranked sections mean 'no preference'. "
                       "Double-click a section to rank / un-rank it."
                  ).pack(anchor="w", pady=(0, 6))
        self.pref_scroll = ScrollFrame(right)
        self.pref_scroll.pack(fill="both", expand=True)

    # --- priority list
    def priority_text(self, c):
        if c == "conflict":
            return "Fewest conflict hours"
        if c == "days":
            return "Fewest days at school"
        y = c[5:]
        return f"{year_label(y)} section preference" + ("" if self.rankings.get(y) else "   (nothing ranked)")

    def refresh_priority_list(self, select=None):
        self.pri_list.delete(0, "end")
        for i, c in enumerate(self.priority, 1):
            self.pri_list.insert("end", f"{i}.  {self.priority_text(c)}")
        if select is not None:
            self.pri_list.selection_set(select)

    def move_priority(self, d):
        sel = self.pri_list.curselection()
        if not sel:
            return
        i, j = sel[0], sel[0] + d
        if 0 <= j < len(self.priority):
            self.priority[i], self.priority[j] = self.priority[j], self.priority[i]
            self.refresh_priority_list(select=j)

    def preset(self, name):
        prefs = [c for c in self.priority if c.startswith("pref:")]
        self.priority = {"default": ["conflict", "days"] + prefs,
                         "days": ["days", "conflict"] + prefs,
                         "prefs": prefs + ["conflict", "days"]}[name]
        self.refresh_priority_list()

    def _sync_priority(self, years):
        want = {f"pref:{y}" for y in years} | {"conflict", "days"}
        keep = [c for c in self.priority if c in want]
        for c in ("conflict", "days"):
            if c not in keep:
                keep.insert(0, c) if c == "conflict" else keep.insert(min(1, len(keep)), c)
        numeric = sorted((y for y in years if y.isdigit()), key=int, reverse=True)
        for y in numeric + sorted(y for y in years if not y.isdigit()):
            if f"pref:{y}" not in keep:
                keep.append(f"pref:{y}")
        self.priority = keep
        self.refresh_priority_list()

    # --- subjects panel
    def refresh_components_panel(self):
        sections_of = defaultdict(set)
        for r in self.rows:
            sections_of[(r.subject, r.part)].add(r.section)
        self.excluded &= set(sections_of)
        self.sub_tree.delete(*self.sub_tree.get_children())
        self._comp_keys = {}
        for i, key in enumerate(sorted(sections_of)):
            iid = str(i)
            self._comp_keys[iid] = key
            self.sub_tree.insert("", "end", iid=iid, values=(
                "☐" if key in self.excluded else "☑", key[0], key[1], len(sections_of[key])))

    def toggle_component(self, event):
        iid = self.sub_tree.identify_row(event.y)
        if not iid:
            return "break"
        key = self._comp_keys[iid]
        self.excluded.symmetric_difference_update({key})
        self.sub_tree.set(iid, "use", "☐" if key in self.excluded else "☑")
        return "break"

    def set_all_components(self, on):
        self.excluded = set() if on else set(self._comp_keys.values())
        for iid, key in self._comp_keys.items():
            self.sub_tree.set(iid, "use", "☐" if key in self.excluded else "☑")

    # --- preferences panel
    def sections_by_year(self):
        by_year = defaultdict(set)
        for r in self.rows:
            by_year[year_of(r.section)].add(r.section)
        return {y: sorted(s, key=natural_key) for y, s in by_year.items()}

    def on_data_changed(self):
        by_year = self.sections_by_year()
        # forget ranked sections that no longer exist
        for y in list(self.rankings):
            self.rankings[y] = [s for s in self.rankings[y] if s in by_year.get(y, [])]
            if not self.rankings[y]:
                del self.rankings[y]
        self.refresh_components_panel()
        sig = tuple(sorted((y, tuple(s)) for y, s in by_year.items()))
        if sig != self._pref_sig:
            self._pref_sig = sig
            self._rebuild_prefs(by_year)
            self._sync_priority(list(by_year))
        else:
            self.refresh_priority_list()

    def _rebuild_prefs(self, by_year):
        inner = self.pref_scroll.inner
        for w in inner.winfo_children():
            w.destroy()
        if not by_year:
            ttk.Label(inner, text="Add some schedules first.").pack(anchor="w")
            return
        for y in sorted(by_year, key=year_sort_key):
            RankPicker(inner, y, by_year[y], self.rankings.get(y, []), self._on_rank_change
                       ).pack(fill="x", pady=(0, 8))

    def _on_rank_change(self, year, ranked):
        if ranked:
            self.rankings[year] = ranked
        else:
            self.rankings.pop(year, None)
        sel = self.pri_list.curselection()
        self.refresh_priority_list(select=sel[0] if sel else None)

    # ------------------------------------------------------------ results tab
    def _build_results_tab(self):
        tab = ttk.Frame(self.nb, padding=8)
        self.nb.add(tab, text="  3 · Results  ")
        self.results_tab = tab
        pw = ttk.PanedWindow(tab, orient="horizontal")
        pw.pack(fill="both", expand=True)

        left = ttk.Frame(pw)
        pw.add(left, weight=1)
        sash_done = []

        def place_sash(_e=None):          # runs once, as soon as the pane has a real width
            if not sash_done and pw.winfo_width() > 700:
                sash_done.append(True)
                pw.sashpos(0, 440)
        pw.bind("<Configure>", place_sash)
        self.res_tree = ttk.Treeview(left, show="headings", selectmode="browse", columns=("rank",))
        self.res_tree.heading("rank", text="#")
        sb = ttk.Scrollbar(left, orient="vertical", command=self.res_tree.yview)
        self.res_tree.configure(yscrollcommand=sb.set)
        self.res_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.res_tree.bind("<<TreeviewSelect>>", self.on_result_select)

        right = ttk.Frame(pw)
        pw.add(right, weight=3)
        bar = ttk.Frame(right)
        bar.pack(fill="x", pady=(0, 6))
        self.res_title = tk.StringVar(value="No results yet - press 'Find best schedules'.")
        ttk.Label(bar, textvariable=self.res_title, font=("TkDefaultFont", 10, "bold")).pack(side="left")
        for text, cmd in [("Export CSV…", self.export_results_csv),
                          ("Export selected to Excel…", lambda: self.export_excel(only_selected=True)),
                          ("Export all to Excel…", lambda: self.export_excel(only_selected=False))]:
            ttk.Button(bar, text=text, command=cmd).pack(side="right", padx=(6, 0))

        vp = ttk.PanedWindow(right, orient="vertical")
        vp.pack(fill="both", expand=True)
        cvf = ttk.Frame(vp)
        vp.add(cvf, weight=4)
        self.cv = tk.Canvas(cvf, background="white", highlightthickness=1, highlightbackground="#bbbbbb")
        csb = ttk.Scrollbar(cvf, orient="vertical", command=self.cv.yview)
        self.cv.configure(yscrollcommand=csb.set)
        self.cv.pack(side="left", fill="both", expand=True)
        csb.pack(side="right", fill="y")
        self.cv.bind("<Configure>", lambda e: self._schedule_redraw())
        self.cv.bind("<MouseWheel>", lambda e: self.cv.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        self.cv.bind("<Button-4>", lambda e: self.cv.yview_scroll(-1, "units"))
        self.cv.bind("<Button-5>", lambda e: self.cv.yview_scroll(1, "units"))

        tf = ttk.Frame(vp)
        vp.add(tf, weight=2)
        self.detail = tk.Text(tf, height=10, wrap="none", font="TkFixedFont", state="disabled")
        dsb = ttk.Scrollbar(tf, orient="vertical", command=self.detail.yview)
        self.detail.configure(yscrollcommand=dsb.set)
        self.detail.pack(side="left", fill="both", expand=True)
        dsb.pack(side="right", fill="y")

    def _schedule_redraw(self):
        if self._redraw_job:
            self.after_cancel(self._redraw_job)
        self._redraw_job = self.after(40, self.draw_timetable)

    def draw_timetable(self):
        self._redraw_job = None
        cv = self.cv
        cv.delete("all")
        r = self.current
        if not r:
            return
        meetings = r["meetings"]
        days = DAY_ORDER[:5] + (["Sat"] if any(m[0] == "Sat" for m in meetings) else [])
        lo = min(m[1] for m in meetings) // 60 * 60
        hi = -(-max(m[2] for m in meetings) // 60) * 60
        left, top = 62, 28
        pph = max(32, min(46, (cv.winfo_height() - top - 8) / max(1, (hi - lo) / 60)))
        width = max(cv.winfo_width(), 560)
        colw = (width - left - 6) / len(days)
        height = top + (hi - lo) / 60 * pph + 8

        def y_of(minute):
            return top + (minute - lo) / 60 * pph

        for h in range(lo // 60, hi // 60 + 1):
            y = y_of(h * 60)
            cv.create_line(left, y, width - 6, y, fill="#e3e3e3")
            cv.create_text(left - 6, y, text=fmt(h * 60), anchor="e", fill="#666666", font=("TkDefaultFont", 8))
        for j, day in enumerate(days):
            x = left + j * colw
            cv.create_line(x, top, x, y_of(hi), fill="#d0d0d0")
            cv.create_text(x + colw / 2, top / 2, text=day, font=("TkDefaultFont", 10, "bold"))
        cv.create_line(left + len(days) * colw, top, left + len(days) * colw, y_of(hi), fill="#d0d0d0")

        for j, day in enumerate(days):
            todays = sorted((m for m in meetings if m[0] == day), key=lambda m: (m[1], m[2]))
            clusters = []
            for m in todays:
                if clusters and m[1] < clusters[-1]["end"]:
                    clusters[-1]["items"].append(m)
                    clusters[-1]["end"] = max(clusters[-1]["end"], m[2])
                else:
                    clusters.append({"end": m[2], "items": [m]})
            for cl in clusters:
                lane_ends, placed = [], []
                for m in cl["items"]:
                    for li, le in enumerate(lane_ends):
                        if le <= m[1]:
                            lane_ends[li] = m[2]
                            placed.append((m, li))
                            break
                    else:
                        lane_ends.append(m[2])
                        placed.append((m, len(lane_ends) - 1))
                nl = len(lane_ends)
                for m, li in placed:
                    clash = any(o is not m and overlap_minutes([o], [m]) > 0 for o in cl["items"])
                    name, sec = split_label(m[3])
                    subject = name.rsplit(" ", 1)[0] if name.endswith((" Lec", " Lab")) else name
                    color = CLASH_COLOR if clash else self.subject_colors.get(subject, PALETTE[0])
                    x0 = left + j * colw + li * colw / nl + 1
                    x1 = left + j * colw + (li + 1) * colw / nl - 1
                    y0, y1 = y_of(m[1]) + 1, y_of(m[2]) - 1
                    cv.create_rectangle(x0, y0, x1, y1, fill=color, outline="#888888")
                    h = y1 - y0
                    lines = [name] if h < 30 else [name, f"Section {sec}"] if h < 52 else \
                        [name, f"Section {sec}", f"{fmt(m[1])} - {fmt(m[2])}"]
                    if h < 30:
                        lines = [f"{name} [{sec}]"]
                    cv.create_text(x0 + 4, y0 + 3, text="\n".join(lines), anchor="nw", width=x1 - x0 - 6,
                                   font=("TkDefaultFont", 8, "bold" if clash else "normal"))
        cv.configure(scrollregion=(0, 0, width, height))

    def on_result_select(self, _event=None):
        sel = self.res_tree.selection()
        if not sel:
            return
        self.current = self.results[int(sel[0])]
        self.draw_timetable()
        r, rk = self.current, self.res_rankings
        lines = [summary_line(r), "", "Sections chosen:"]
        for key, sec in r["assignment"]:
            lines.append(f"  {comp_label(key):<20} -> {sec}{choice_note(sec, rk)}")
        lines += ["", "Timetable:"] + timetable_lines(r["meetings"])
        self.res_title.set(f"Schedule #{r['rank']}")
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", "\n".join(lines))
        self.detail.configure(state="disabled")

    def fill_results(self):
        years = sorted({y for r in self.results for y in r["prefs"]}, key=year_sort_key)
        cols = ["rank", "conflict", "days", "used"] + [f"pref{y}" for y in years]
        self.res_tree.configure(columns=cols)
        heads = {"rank": ("#", 34), "conflict": ("Conflict hr", 94), "days": ("Days", 44),
                 "used": ("Days used", 120)}
        for y in years:
            heads[f"pref{y}"] = ((f"Yr {y} pref" if y.isdigit() else "Other pref"), 78)
        for c in cols:
            self.res_tree.heading(c, text=heads[c][0])
            self.res_tree.column(c, width=heads[c][1], anchor="center", stretch=(c == "used"))
        self.res_tree.delete(*self.res_tree.get_children())
        for i, r in enumerate(self.results):
            vals = [r["rank"], f"{r['conflict']:.1f}", r["days"], " ".join(d[:2] if d != "Thu" else "Th" for d in r["days_used"])]
            vals += [f"{r['prefs'][y][0]}/{r['prefs'][y][1]}" if y in r["prefs"] else "-" for y in years]
            self.res_tree.insert("", "end", iid=str(i), values=vals)
        if self.results:
            self.res_tree.selection_set("0")
        else:
            self.current = None
            self.draw_timetable()

    # ----------------------------------------------------------------- search
    def start_search(self):
        if self.worker and self.worker.is_alive():
            return
        keys, comps = build_components(self.rows, self.excluded)
        if not comps:
            messagebox.showwarning("Nothing to schedule",
                                   "Add some class meetings (tab 1) and tick at least one subject (tab 2).")
            return
        try:
            top_n = max(1, min(500, int(self.top_n_var.get())))
        except (tk.TclError, ValueError):
            messagebox.showwarning("Schedules to keep", "Enter a whole number for 'Schedules to keep'.")
            return
        max_days = None if self.max_days_var.get() == "No limit" else int(self.max_days_var.get())
        max_conf = None
        if self.max_conf_var.get().strip():
            try:
                max_conf = int(round(float(self.max_conf_var.get()) * 60))
                if max_conf < 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Max conflict hours", "Enter a number such as 0 or 1.5 (or leave it blank).")
                return

        by_year = self.sections_by_year()
        rankings = {y: [s for s in r if s in by_year.get(y, [])] for y, r in self.rankings.items()}
        rankings = {y: r for y, r in rankings.items() if r}
        priority = list(self.priority)
        subjects = sorted({k[0] for k in keys})
        self.subject_colors = {s: PALETTE[i % len(PALETTE)] for i, s in enumerate(subjects)}

        total = 1
        for c in comps:
            total *= len(c)
        self.stop_event = threading.Event()
        self.search_state = {"nodes": 0, "progress": 0.0}
        self.search_out = {}
        state, stop, out = self.search_state, self.stop_event, self.search_out

        def work():
            try:
                out["best"], out["nodes"], out["cancelled"] = search(
                    comps, rankings, priority, top_n, max_conf, max_days, stop, state)
            except Exception as ex:  # report in the GUI instead of dying silently
                out["error"] = f"{type(ex).__name__}: {ex}"
            finally:
                out["done"] = True

        self.ctx = {"keys": keys, "comps": comps, "rankings": rankings, "total": total}
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.progress["value"] = 0
        self.t0 = time.perf_counter()
        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()
        self.after(100, self._poll)

    def stop_search(self):
        self.stop_event.set()
        self.status.set("Stopping…")

    def _poll(self):
        out, state = self.search_out, self.search_state
        elapsed = time.perf_counter() - self.t0
        self.progress["value"] = state.get("progress", 0) * 1000
        if not out.get("done"):
            self.status.set(f"Searching {self.ctx['total']:,} combinations…  {elapsed:.1f}s  ·  "
                            f"{state.get('nodes', 0):,} steps  ·  {state.get('progress', 0) * 100:.0f}%")
            self.after(150, self._poll)
            return
        self.run_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        if "error" in out:
            self.status.set("Search failed.")
            messagebox.showerror("Search failed", out["error"])
            return
        ctx = self.ctx
        self.res_rankings = ctx["rankings"]
        self.results = []
        for rank, entry in enumerate(out["best"], 1):
            r = describe(entry, ctx["keys"], ctx["comps"], ctx["rankings"])
            r["rank"] = rank
            self.results.append(r)
        self.fill_results()
        self.progress["value"] = 1000 if not out["cancelled"] else self.progress["value"]
        if not self.results:
            self.status.set("No schedule satisfies your hard limits.")
            messagebox.showinfo("No schedules found",
                                "No combination satisfies your hard limits (max days / max conflict hours). "
                                "Try relaxing them.")
            return
        note = "  (stopped early - results may not be the very best)" if out["cancelled"] else ""
        self.status.set(f"Done in {elapsed:.2f}s · {len(self.results)} schedule(s) from "
                        f"{ctx['total']:,} combinations.{note}")
        self.nb.select(self.results_tab)

    # ----------------------------------------------------------------- export
    def export_excel(self, only_selected):
        if not self.results:
            messagebox.showinfo("Nothing to export", "Run the search first.")
            return
        chosen = self.results
        if only_selected:
            sel = self.res_tree.selection()
            chosen = [self.results[int(sel[0])]] if sel else []
            if not chosen:
                messagebox.showinfo("Nothing selected", "Select a schedule in the list first.")
                return
        path = filedialog.asksaveasfilename(title="Export to Excel", defaultextension=".xlsx",
                                            initialfile="schedules.xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            export_xlsx(chosen, path)
        except (RuntimeError, OSError) as ex:
            messagebox.showerror("Export failed", str(ex))
            return
        self.status.set(f"Saved {len(chosen)} sheet(s) to {os.path.basename(path)}")

    def export_results_csv(self):
        if not self.results:
            messagebox.showinfo("Nothing to export", "Run the search first.")
            return
        path = filedialog.asksaveasfilename(title="Export results to CSV", defaultextension=".csv",
                                            initialfile="schedules.csv", filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        years = sorted({y for r in self.results for y in r["prefs"]}, key=year_sort_key)
        comp_keys = [k for k, _ in self.results[0]["assignment"]]
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["rank", "conflict_hours", "days_at_school", "days_used"]
                       + [f"{year_label(y)} preference" for y in years] + [comp_label(k) for k in comp_keys])
            for r in self.results:
                sec = dict(r["assignment"])
                w.writerow([r["rank"], f"{r['conflict']:.2f}", r["days"], " ".join(r["days_used"])]
                           + [f"{r['prefs'][y][0]}/{r['prefs'][y][1]}" if y in r["prefs"] else "" for y in years]
                           + [sec[k] for k in comp_keys])
        self.status.set(f"Saved {len(self.results)} schedule(s) to {os.path.basename(path)}")


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
