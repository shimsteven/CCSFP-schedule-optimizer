#!/usr/bin/env python3
"""
Schedule Optimizer - GUI (customtkinter edition)
================================================

Run:      python schedule_optimizer_ctk.py
Needs:    Python 3.10+, `pip install customtkinter`
          and, for Excel export, `pip install openpyxl`

Workflow
  1. Schedules  - type in (or paste / import) the weekly meetings of every section.
  2. Options    - choose what matters most (priority order), rank your preferred
                  sections for each year level, set optional hard limits, and
                  untick subjects you are not taking.
  3. Press "Find best schedules" - results appear in the Results page, with a weekly
                  timetable, and can be exported to Excel / CSV.

How the ranking works
  Every subject part (e.g. "DSA Lec", "DSA Lab", or an ALL-in-one subject) is assigned
  to exactly one section. All combinations are searched (branch-and-bound) and ranked
  by your priority list, e.g.  conflicts > days > Year-3 preference > Year-2 preference.
  Section preference: for a year level with N ranked sections, each subject part placed
  in your #1 section earns N points, #2 earns N-1, ... , unranked sections earn 0.
  More points is better.

Note: customtkinter has no table or canvas widget, so the tables (ttk.Treeview) and the
timetable (tk.Canvas) are plain tkinter widgets styled to follow the active theme.
"""
from __future__ import annotations

import csv
import json
import os
import re
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
from collections import defaultdict
from tkinter import filedialog, ttk

try:
    import customtkinter as ctk
except ImportError:
    raise SystemExit("This app needs customtkinter.  Install it with:  pip install customtkinter")

from engine import (
    DAY_ORDER, PARTS, Row, SAMPLE, SAMPLE_PRIORITY, SAMPLE_RANKINGS, build_components,
    choice_note, comp_label, describe, export_xlsx, fmt, natural_key, overlap_minutes,
    row_sort_key, search, split_label, summary_line, timetable_lines, to_minutes, year_label,
    year_of, year_sort_key,
)

# ----------------------------------------------------------------------------
# Look & feel  (every colour is a (light, dark) pair - customtkinter picks one)
# ----------------------------------------------------------------------------
ACCENT, ACCENT_HOVER = "#6366F1", "#4F46E5"
DANGER, DANGER_HOVER = "#EF4444", "#DC2626"
BG = ("#F3F4F8", "#0F1117")
SIDEBAR = ("#E8EAF2", "#161922")
CARD = ("#FFFFFF", "#1B1F2A")
CARD_ALT = ("#F1F3F9", "#252A38")
BORDER = ("#DDE1EB", "#2B3140")
MUTED = ("#6B7280", "#8B93A7")
TEXT = ("#111827", "#E5E7EB")
SECONDARY = ("#E3E6F0", "#2A3040")
SECONDARY_HOVER = ("#D3D8E6", "#363E54")

BUTTON_KINDS = {
    "primary": dict(fg_color=ACCENT, hover_color=ACCENT_HOVER, text_color="white"),
    "danger": dict(fg_color=DANGER, hover_color=DANGER_HOVER, text_color="white"),
    "secondary": dict(fg_color=SECONDARY, hover_color=SECONDARY_HOVER, text_color=TEXT),
    "ghost": dict(fg_color="transparent", hover_color=SECONDARY, text_color=TEXT,
                  border_width=1, border_color=BORDER),
}

# timetable block colours (pastels read well on both themes)
PALETTE = ["#BFDBFE", "#BBF7D0", "#FDE68A", "#DDD6FE", "#FBCFE8", "#99F6E4",
           "#FEF08A", "#CBD5E1", "#FED7AA", "#D9F99D"]
CLASH_COLOR = "#FCA5A5"
BLOCK_TEXT = "#1F2937"


def F(size=13, weight="normal", **kw):
    return ctk.CTkFont(size=size, weight=weight, **kw)


def card(master, **kw):
    kw.setdefault("corner_radius", 14)
    return ctk.CTkFrame(master, fg_color=CARD, border_width=1, border_color=BORDER, **kw)


def button(master, text, command, kind="secondary", **kw):
    opts = dict(corner_radius=9, height=34, font=F(13))
    opts.update(BUTTON_KINDS[kind])
    opts.update(kw)
    return ctk.CTkButton(master, text=text, command=command, **opts)


def mini_button(master, text, command, **kw):
    opts = dict(width=28, height=26, corner_radius=7, font=F(11), fg_color=SECONDARY,
                hover_color=SECONDARY_HOVER, text_color=TEXT, text_color_disabled=MUTED)
    opts.update(kw)
    return ctk.CTkButton(master, text=text, command=command, **opts)


def label(master, text, size=13, weight="normal", muted=False, **kw):
    kw.setdefault("anchor", "w")
    return ctk.CTkLabel(master, text=text, font=F(size, weight),
                        text_color=MUTED if muted else TEXT, **kw)


def entry(master, var, width):
    return ctk.CTkEntry(master, textvariable=var, width=width, height=34, corner_radius=8,
                        border_color=BORDER, fg_color=CARD_ALT, text_color=TEXT)


def combo(master, var, values, width):
    return ctk.CTkComboBox(master, variable=var, values=values, width=width, height=34,
                           corner_radius=8, border_color=BORDER, fg_color=CARD_ALT,
                           button_color=SECONDARY, button_hover_color=SECONDARY_HOVER,
                           text_color=TEXT, dropdown_fg_color=CARD, dropdown_text_color=TEXT,
                           dropdown_hover_color=SECONDARY)


def optionmenu(master, var, values, width, command=None):
    return ctk.CTkOptionMenu(master, variable=var, values=values, width=width, height=34,
                             corner_radius=8, fg_color=SECONDARY, button_color=SECONDARY_HOVER,
                             button_hover_color=BORDER, text_color=TEXT, dropdown_fg_color=CARD,
                             dropdown_text_color=TEXT, dropdown_hover_color=SECONDARY,
                             command=command)


def segmented(master, var, values):
    return ctk.CTkSegmentedButton(master, values=values, variable=var, height=32,
                                  selected_color=ACCENT, selected_hover_color=ACCENT_HOVER,
                                  unselected_color=CARD_ALT, unselected_hover_color=SECONDARY_HOVER,
                                  text_color=("#111827", "#E5E7EB"))


def page_header(master, title, subtitle):
    f = ctk.CTkFrame(master, fg_color="transparent")
    label(f, title, 26, "bold").pack(anchor="w")
    label(f, subtitle, 13, muted=True).pack(anchor="w")
    return f


def round_rect(cv, x0, y0, x1, y1, r=8, **kw):
    r = max(0, min(r, (x1 - x0) / 2, (y1 - y0) / 2))
    pts = [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1, x1 - r, y1,
           x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]
    return cv.create_polygon(pts, smooth=True, **kw)


class Popup(ctk.CTkToplevel):
    """Modal message / yes-no dialog that follows the theme."""
    KINDS = {"info": ("i", "#3B82F6"), "warning": ("!", "#F59E0B"),
             "error": ("×", DANGER), "question": ("?", ACCENT)}

    def __init__(self, master, title, message, kind="info", yes_no=False):
        super().__init__(master)
        self.withdraw()
        self.title(title)
        self.resizable(False, False)
        self.configure(fg_color=CARD)
        self.result = False
        glyph, color = self.KINDS.get(kind, self.KINDS["info"])

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=(22, 18))
        body.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(body, text=glyph, width=42, height=42, corner_radius=21, fg_color=color,
                     text_color="white", font=F(20, "bold")).grid(
            row=0, column=0, rowspan=2, sticky="n", padx=(0, 16))
        label(body, title, 16, "bold").grid(row=0, column=1, sticky="w")
        label(body, message, 13, wraplength=400, justify="left").grid(
            row=1, column=1, sticky="w", pady=(6, 0))

        btns = ctk.CTkFrame(body, fg_color="transparent")
        btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=(20, 0))
        if yes_no:
            button(btns, "No", self._no, "ghost", width=84).pack(side="left", padx=(0, 8))
            button(btns, "Yes", self._yes, "primary", width=84).pack(side="left")
        else:
            button(btns, "OK", self._yes, "primary", width=84).pack(side="left")
            self.bind("<Return>", lambda e: self._yes())
        self.bind("<Escape>", lambda e: self._no())
        self.protocol("WM_DELETE_WINDOW", self._no)

        self.update_idletasks()
        try:
            root = master.winfo_toplevel()
            x = root.winfo_rootx() + (root.winfo_width() - self.winfo_reqwidth()) // 2
            y = root.winfo_rooty() + (root.winfo_height() - self.winfo_reqheight()) // 3
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except tk.TclError:
            pass
        self.transient(master.winfo_toplevel())
        self.deiconify()
        self.after(100, self._grab)

    def _grab(self):
        try:
            self.grab_set()
            self.focus_force()
        except tk.TclError:
            pass

    def _yes(self):
        self.result = True
        self.destroy()

    def _no(self):
        self.result = False
        self.destroy()


def popup(master, title, message, kind="info", yes_no=False):
    p = Popup(master, title, message, kind, yes_no)
    master.wait_window(p)
    return p.result


class RankPicker(ctk.CTkFrame):
    """Rank sections of one year level: click a chip to rank it, then reorder with the arrows."""
    CHIP_COLS = 4

    def __init__(self, master, year, sections, ranked, on_change):
        super().__init__(master, corner_radius=12, fg_color=CARD_ALT)
        self.year, self.sections, self.on_change = year, list(sections), on_change
        self.ranked = [s for s in ranked if s in self.sections]

        label(self, year_label(year), 15, "bold").pack(fill="x", padx=14, pady=(12, 0))
        label(self, "Click a section to rank it. Top of the list is your favourite.",
              12, muted=True).pack(fill="x", padx=14)
        self.ranked_box = ctk.CTkFrame(self, fg_color="transparent")
        self.ranked_box.pack(fill="x", padx=10, pady=(8, 0))
        label(self, "Not ranked", 12, "bold", muted=True).pack(fill="x", padx=14, pady=(8, 0))
        self.free_box = ctk.CTkFrame(self, fg_color="transparent")
        self.free_box.pack(fill="x", padx=10, pady=(4, 12))
        self._fill()

    def _free(self):
        return [s for s in self.sections if s not in self.ranked]

    def _fill(self):
        if not self.winfo_exists():
            return
        for w in self.ranked_box.winfo_children():
            w.destroy()
        for w in self.free_box.winfo_children():
            w.destroy()
        if not self.ranked:
            label(self.ranked_box, "Nothing ranked - no preference.", 12, muted=True).pack(
                fill="x", padx=4, pady=2)
        for i, s in enumerate(self.ranked):
            row = ctk.CTkFrame(self.ranked_box, corner_radius=8, fg_color=CARD)
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=str(i + 1), width=24, height=24, corner_radius=12,
                         fg_color=ACCENT, text_color="white", font=F(11, "bold")).pack(
                side="left", padx=(8, 8), pady=6)
            label(row, s, 13).pack(side="left", fill="x", expand=True)
            mini_button(row, "✕", lambda s=s: self.unrank(s), hover_color=DANGER).pack(
                side="right", padx=(0, 6))
            mini_button(row, "▼", lambda i=i: self.move(i, 1),
                        state="normal" if i < len(self.ranked) - 1 else "disabled").pack(side="right", padx=2)
            mini_button(row, "▲", lambda i=i: self.move(i, -1),
                        state="normal" if i > 0 else "disabled").pack(side="right", padx=2)
        for n, s in enumerate(self._free()):
            button(self.free_box, s, lambda s=s: self.rank(s), "ghost", width=70, height=30,
                   corner_radius=15).grid(row=n // self.CHIP_COLS, column=n % self.CHIP_COLS,
                                          padx=3, pady=3, sticky="w")

    def _changed(self):
        self.on_change(self.year, list(self.ranked))
        self.after(10, self._fill)       # rebuild after the clicked button's callback has returned

    def rank(self, s):
        if s not in self.ranked:
            self.ranked.append(s)
        self._changed()

    def unrank(self, s):
        if s in self.ranked:
            self.ranked.remove(s)
        self._changed()

    def move(self, i, d):
        j = i + d
        if 0 <= j < len(self.ranked):
            self.ranked[i], self.ranked[j] = self.ranked[j], self.ranked[i]
            self._changed()


# ----------------------------------------------------------------------------
# Main application
# ----------------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Schedule Optimizer")
        self.geometry("1340x860")
        self.minsize(1140, 720)
        self.configure(fg_color=BG)
        self.mono = tkfont.nametofont("TkFixedFont").actual("family")
        self.ui_font = tkfont.nametofont("TkDefaultFont").actual("family")
        self.apply_tree_style()

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
        self._comp_vars: dict[tuple[str, str], tk.BooleanVar] = {}
        self._section_values: list[str] = []
        self._filter_values = ["All sections"]
        self._redraw_job = None
        self.current_page = "schedules"

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(0, weight=1)
        host = ctk.CTkFrame(main, fg_color="transparent")
        host.grid(row=0, column=0, sticky="nsew", padx=24, pady=(22, 0))
        host.grid_columnconfigure(0, weight=1)
        host.grid_rowconfigure(0, weight=1)
        self.pages = {}
        for key in ("schedules", "options", "results"):
            p = ctk.CTkFrame(host, fg_color="transparent")
            p.grid(row=0, column=0, sticky="nsew")
            self.pages[key] = p

        self._build_runbar(main)
        self._build_schedule_page()
        self._build_options_page()
        self._build_results_page()
        self.load_sample()
        self.show_page("schedules")

    # ------------------------------------------------------------ theme helpers
    def _scale(self):
        try:
            return float(ctk.ScalingTracker.get_window_scaling(self))
        except Exception:
            return 1.0

    def apply_tree_style(self):
        dark = ctk.get_appearance_mode() == "Dark"
        bg, fg = ("#1B1F2A", "#E5E7EB") if dark else ("#FFFFFF", "#111827")
        head = "#252A38" if dark else "#EEF0F6"
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", background=bg, fieldbackground=bg, foreground=fg, borderwidth=0,
                        rowheight=round(30 * self._scale()), font=(self.ui_font, 10))
        style.map("Treeview", background=[("selected", ACCENT)], foreground=[("selected", "#FFFFFF")])
        style.configure("Treeview.Heading", background=head, foreground=fg, relief="flat",
                        borderwidth=0, padding=(8, 7), font=(self.ui_font, 10, "bold"))
        style.map("Treeview.Heading", background=[("active", head)])

    def toggle_theme(self):
        ctk.set_appearance_mode("dark" if self.theme_switch.get() else "light")
        self.after(80, self._restyle)

    def _restyle(self):
        self.apply_tree_style()
        for pane in (self.opane, self.lpane, self.hpane, self.vpane):
            pane.configure(background=self._pane_bg())
        self._schedule_redraw()

    def msg(self, title, text, kind="info", parent=None):
        popup(parent or self, title, text, kind)

    def confirm(self, title, text):
        return popup(self, title, text, "question", yes_no=True)

    # ---------------------------------------------------------------- sidebar
    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color=SIDEBAR)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_rowconfigure(5, weight=1)

        head = ctk.CTkFrame(sb, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=20, pady=(26, 22))
        label(head, "Schedule", 22, "bold").pack(anchor="w")
        ctk.CTkLabel(head, text="Optimizer", font=F(22, "bold"), text_color=ACCENT, anchor="w").pack(anchor="w")
        label(head, "find your best timetable", 12, muted=True).pack(anchor="w", pady=(2, 0))

        self.nav_btns = {}
        for r, (key, text) in enumerate([("schedules", "1   Schedules"), ("options", "2   Options"),
                                         ("results", "3   Results")], start=1):
            b = ctk.CTkButton(sb, text=text, anchor="w", height=42, corner_radius=10, font=F(14, "bold"),
                              fg_color="transparent", hover_color=SECONDARY, text_color=TEXT,
                              command=lambda k=key: self.show_page(k))
            b.grid(row=r, column=0, sticky="ew", padx=14, pady=3)
            self.nav_btns[key] = b

        tools = ctk.CTkFrame(sb, fg_color="transparent")
        tools.grid(row=6, column=0, sticky="ew", padx=14, pady=(0, 10))
        label(tools, "PROJECT", 11, "bold", muted=True).pack(fill="x", padx=6, pady=(0, 4))
        for text, cmd in [("New project", self.new_project), ("Open project…", self.open_project),
                          ("Save project…", self.save_project), ("How ranking works", self.show_help)]:
            button(tools, text, cmd, "ghost", anchor="w", height=32).pack(fill="x", pady=2)
        self.theme_switch = ctk.CTkSwitch(tools, text="Dark mode", command=self.toggle_theme,
                                          progress_color=ACCENT, font=F(13), text_color=TEXT)
        self.theme_switch.pack(anchor="w", padx=8, pady=(14, 4))
        self.theme_switch.select()

    def show_page(self, key):
        self.pages[key].tkraise()
        self.current_page = key
        for k, b in self.nav_btns.items():
            if k == key:
                b.configure(fg_color=ACCENT, hover_color=ACCENT_HOVER, text_color="white")
            else:
                b.configure(fg_color="transparent", hover_color=SECONDARY, text_color=TEXT)

    def show_help(self):
        self.msg("How the ranking works",
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
    def _build_runbar(self, parent):
        bar = card(parent)
        bar.grid(row=1, column=0, sticky="ew", padx=24, pady=(12, 18))
        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=12)
        self.run_btn = button(inner, "▶   Find best schedules", self.start_search, "primary",
                              height=40, width=200, font=F(14, "bold"))
        self.run_btn.pack(side="left")
        self.stop_btn = button(inner, "Stop", self.stop_search, "ghost", height=40, width=80, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 16))
        self.progress = ctk.CTkProgressBar(inner, width=260, height=10, progress_color=ACCENT, fg_color=CARD_ALT)
        self.progress.set(0)
        self.progress.pack(side="left")
        self.status = tk.StringVar(value="Ready.")
        ctk.CTkLabel(inner, textvariable=self.status, font=F(13), text_color=MUTED, anchor="w").pack(
            side="left", padx=16)

    # ----------------------------------------------------------- schedules page
    def _build_schedule_page(self):
        page = self.pages["schedules"]
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(3, weight=1)
        page_header(page, "Class schedules",
                    "Enter every section's weekly meetings, paste them from Excel, or import a CSV."
                    ).grid(row=0, column=0, sticky="w")

        form = card(page)
        form.grid(row=1, column=0, sticky="ew", pady=(16, 10))
        self.f_section, self.f_day = tk.StringVar(), tk.StringVar(value="Mon")
        self.f_subject, self.f_part = tk.StringVar(), tk.StringVar(value="LEC")
        self.f_start, self.f_end = tk.StringVar(), tk.StringVar()
        pad = dict(padx=(16, 0), sticky="w")
        for c, text in enumerate(["Section", "Subject", "Start", "End"]):
            label(form, text, 12, "bold", muted=True).grid(row=0, column=c, pady=(14, 2), **pad)
        self.cb_section = combo(form, self.f_section, [], 120)
        self.cb_subject = combo(form, self.f_subject, [], 250)
        e_start, e_end = entry(form, self.f_start, 110), entry(form, self.f_end, 110)
        for c, w in enumerate([self.cb_section, self.cb_subject, e_start, e_end]):
            w.grid(row=1, column=c, **pad)
        for w in (self.cb_section, self.cb_subject, e_start, e_end):
            w.bind("<Return>", lambda e: self.add_row())
        label(form, "Day", 12, "bold", muted=True).grid(row=2, column=0, pady=(12, 2), **pad)
        label(form, "Type", 12, "bold", muted=True).grid(row=2, column=2, pady=(12, 2), **pad)
        segmented(form, self.f_day, DAY_ORDER).grid(row=3, column=0, columnspan=2, **pad)
        segmented(form, self.f_part, PARTS).grid(row=3, column=2, columnspan=2, **pad)

        fb = ctk.CTkFrame(form, fg_color="transparent")
        fb.grid(row=4, column=0, columnspan=5, sticky="w", padx=16, pady=(14, 4))
        button(fb, "Add", self.add_row, "primary", width=90).pack(side="left", padx=(0, 8))
        button(fb, "Save changes", self.update_row, width=120).pack(side="left", padx=4)
        button(fb, "Delete", self.delete_rows, "danger", width=90).pack(side="left", padx=4)
        button(fb, "Clear form", self.clear_form, "ghost", width=100).pack(side="left", padx=4)
        label(form, "Times: 7:00AM, 2 PM or 14:00.   LEC and LAB are chosen separately (they may come from "
                    "different sections); ALL = one combined class.   Click a row below to edit it.",
              12, muted=True, wraplength=900, justify="left").grid(
            row=5, column=0, columnspan=5, sticky="w", padx=16, pady=(0, 14))

        top = ctk.CTkFrame(page, fg_color="transparent")
        top.grid(row=2, column=0, sticky="ew", pady=(4, 6))
        label(top, "Show", 13, muted=True).pack(side="left")
        self.filter_var = tk.StringVar(value="All sections")
        self.cb_filter = optionmenu(top, self.filter_var, ["All sections"], 150,
                                    command=lambda _v: self.refresh_tree())
        self.cb_filter.pack(side="left", padx=10)
        self.count_var = tk.StringVar()
        ctk.CTkLabel(top, textvariable=self.count_var, font=F(13), text_color=MUTED).pack(side="right")

        tcard = card(page)
        tcard.grid(row=3, column=0, sticky="nsew")
        cols = ("section", "day", "subject", "part", "start", "end")
        self.tree = ttk.Treeview(tcard, columns=cols, show="headings", selectmode="extended")
        for c, (text, w) in zip(cols, [("Section", 100), ("Day", 80), ("Subject", 260),
                                       ("Type", 80), ("Start", 100), ("End", 100)]):
            self.tree.heading(c, text=text)
            self.tree.column(c, width=w, anchor="w")
        sb = ctk.CTkScrollbar(tcard, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        sb.pack(side="right", fill="y", padx=(4, 6), pady=10)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)
        self.tree.bind("<Delete>", lambda e: self.delete_rows())

        btns = ctk.CTkFrame(page, fg_color="transparent")
        btns.grid(row=4, column=0, sticky="w", pady=(10, 0))
        for text, cmd in [("Paste rows…", self.paste_rows), ("Import CSV…", self.import_csv),
                          ("Export CSV…", self.export_schedule_csv), ("Load sample data", self.load_sample),
                          ("Clear all", self.clear_all)]:
            button(btns, text, cmd, "ghost").pack(side="left", padx=(0, 8))

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
            self.msg("Can't add", str(ex), "warning")
            return
        if self._exists(values):
            self.msg("Already there", "That exact meeting is already in the list.")
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
            self.msg("Select one row", "Click a single row in the table first, edit the form, "
                                       "then press Save changes.")
            return
        try:
            values = self._read_form()
        except ValueError as ex:
            self.msg("Can't save", str(ex), "warning")
            return
        uid = int(sel[0])
        if self._exists(values, ignore_uid=uid):
            self.msg("Already there", "That exact meeting is already in the list.")
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
        self._section_values = sections
        self.cb_section.configure(values=sections)
        self.cb_subject.configure(values=subjects)
        self._filter_values = ["All sections"] + sections
        self.cb_filter.configure(values=self._filter_values)
        if self.filter_var.get() not in self._filter_values:
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
        win = ctk.CTkToplevel(self)
        win.title("Paste rows")
        win.geometry("660x500")
        win.configure(fg_color=BG)
        win.transient(self)
        win.after(150, win.focus_force)
        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=22, pady=18)
        label(body, "Paste rows", 20, "bold").pack(anchor="w")
        label(body, "One class meeting per line:   Day, Subject, Type, Start, End\n"
                    "(commas, tabs or | all work, so you can paste straight from Excel).\n"
                    "If a line has 6 values, the first one is taken as the section.",
              12, muted=True, justify="left").pack(anchor="w", pady=(4, 10))
        top = ctk.CTkFrame(body, fg_color="transparent")
        top.pack(fill="x")
        label(top, "Section for lines without one", 13, muted=True).pack(side="left")
        sec_var = tk.StringVar(value=self.f_section.get())
        combo(top, sec_var, list(self._section_values), 130).pack(side="left", padx=10)
        txt = ctk.CTkTextbox(body, wrap="none", font=ctk.CTkFont(family=self.mono, size=13),
                             corner_radius=10, fg_color=CARD, border_width=1, border_color=BORDER)
        txt.pack(fill="both", expand=True, pady=12)
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
                self.msg("Fix these lines", "\n".join(errors[:12]), "warning", parent=win)
                return
            added = 0
            for values in new_rows:
                if not self._exists(values):
                    self.rows.append(self._new_row(values))
                    added += 1
            self.refresh_tree()
            win.destroy()
            self.status.set(f"Added {added} meeting(s).")

        btns = ctk.CTkFrame(body, fg_color="transparent")
        btns.pack(fill="x")
        button(btns, "Add rows", do_add, "primary", width=110).pack(side="right")
        button(btns, "Cancel", win.destroy, "ghost", width=90).pack(side="right", padx=8)

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
            self.msg("Import failed", str(ex), "error")
            return
        self.refresh_tree()
        text = f"Imported {added} meeting(s)."
        if skipped:
            text += f"\nSkipped {len(skipped)} unreadable row(s) (file lines: {', '.join(map(str, skipped[:15]))})."
        text += "\n\nExpected columns: section, day, subject, type, start, end"
        self.msg("CSV import", text)

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
        if self.rows and not self.confirm("Clear all", "Remove every class meeting from the list?"):
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
        if self.rows and not self.confirm(
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
            "keep": self._top_n() or 15,
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
            self.top_n_var.set(str(int(data.get("keep", 15))))
        except (OSError, ValueError, KeyError, TypeError) as ex:
            self.msg("Can't open project", f"That file couldn't be read:\n{ex}", "error")
            return
        self._pref_sig = None
        self.refresh_tree()
        self.status.set(f"Opened {os.path.basename(path)}")

    # ------------------------------------------------------------ options page
    def _build_options_page(self):
        page = self.pages["options"]
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)
        page_header(page, "Options", "Decide what matters most, which subjects to take, and which sections you like. "
                                     "Drag the dividers to resize the panels."
                    ).grid(row=0, column=0, sticky="w", pady=(0, 14))

        # draggable dividers between the three panels
        self.opane = tk.PanedWindow(page, orient="horizontal", sashwidth=14, sashrelief="flat",
                                    borderwidth=0, showhandle=False, opaqueresize=False,
                                    background=self._pane_bg())
        self.opane.grid(row=1, column=0, sticky="nsew")

        # -- left column: priority + limits
        # stacked vertically with its own draggable divider (priority on top, limits below)
        self.lpane = tk.PanedWindow(self.opane, orient="vertical", sashwidth=14, sashrelief="flat",
                                    borderwidth=0, showhandle=False, opaqueresize=False,
                                    background=self._pane_bg())
        self.opane.add(self.lpane, minsize=330, width=350, stretch="never")
        pri = card(self.lpane)
        label(pri, "What matters most?", 16, "bold").pack(fill="x", padx=16, pady=(14, 0))
        label(pri, "The top item decides; lower items only break ties.", 12, muted=True).pack(
            fill="x", padx=16, pady=(0, 6))
        pscroll = ctk.CTkScrollableFrame(pri, fg_color="transparent")   # scrolls when the panel is short
        pscroll.pack(fill="both", expand=True, padx=6, pady=(0, 8))
        self.pri_box = ctk.CTkFrame(pscroll, fg_color="transparent")
        self.pri_box.pack(fill="x", padx=4)
        label(pscroll, "Quick presets", 12, "bold", muted=True).pack(fill="x", padx=8, pady=(12, 4))
        for text, name in [("Conflicts → days → preferences", "default"),
                           ("Fewest days first", "days"), ("Preferences first", "prefs")]:
            button(pscroll, text, lambda n=name: self.preset(n), "ghost", anchor="w").pack(
                fill="x", padx=8, pady=2)

        lim = card(self.lpane)
        self.lpane.add(pri, minsize=200, stretch="always")
        self.lpane.add(lim, minsize=215, height=240, stretch="never")
        self.max_days_var = tk.StringVar(value="No limit")
        self.max_conf_var = tk.StringVar(value="")
        self.top_n_var = tk.StringVar(value="15")
        label(lim, "Hard limits & output", 16, "bold").grid(row=0, column=0, columnspan=2, sticky="w",
                                                            padx=16, pady=(14, 8))
        label(lim, "Max days at school", 13).grid(row=1, column=0, sticky="w", padx=(16, 8), pady=4)
        optionmenu(lim, self.max_days_var, ["No limit"] + [str(i) for i in range(1, len(DAY_ORDER) + 1)],
                   110).grid(row=1, column=1, sticky="e", padx=(0, 16), pady=4)
        label(lim, "Max conflict hours", 13).grid(row=2, column=0, sticky="w", padx=(16, 8), pady=4)
        entry(lim, self.max_conf_var, 110).grid(row=2, column=1, sticky="e", padx=(0, 16), pady=4)
        label(lim, "blank = no limit,  0 = no conflicts allowed", 12, muted=True).grid(
            row=3, column=0, columnspan=2, sticky="w", padx=16)
        label(lim, "Schedules to keep", 13).grid(row=4, column=0, sticky="w", padx=(16, 8), pady=(12, 14))
        entry(lim, self.top_n_var, 110).grid(row=4, column=1, sticky="e", padx=(0, 16), pady=(12, 14))

        # -- middle column: subjects to include
        mid = card(self.opane)
        self.opane.add(mid, minsize=230, width=290, stretch="never")
        label(mid, "Subjects to schedule", 16, "bold").pack(fill="x", padx=16, pady=(14, 0))
        label(mid, "Untick what you are not taking.", 12, muted=True).pack(fill="x", padx=16, pady=(0, 6))
        self.sub_scroll = ctk.CTkScrollableFrame(mid, fg_color="transparent", width=250)
        self.sub_scroll.pack(fill="both", expand=True, padx=6)
        b = ctk.CTkFrame(mid, fg_color="transparent")
        b.pack(fill="x", padx=16, pady=(6, 14))
        button(b, "Select all", lambda: self.set_all_components(True), "ghost", width=100).pack(side="left")
        button(b, "Select none", lambda: self.set_all_components(False), "ghost", width=100).pack(
            side="left", padx=8)

        # -- right column: section preferences per year level
        right = card(self.opane)
        self.opane.add(right, minsize=320, stretch="always")
        label(right, "Section preferences by year level", 16, "bold").pack(fill="x", padx=16, pady=(14, 0))
        label(right, "Unranked sections mean 'no preference'.", 12, muted=True).pack(
            fill="x", padx=16, pady=(0, 6))
        self.pref_scroll = ctk.CTkScrollableFrame(right, fg_color="transparent")
        self.pref_scroll.pack(fill="both", expand=True, padx=6, pady=(0, 8))

    # --- priority list
    def priority_text(self, c):
        if c == "conflict":
            return "Fewest conflict hours"
        if c == "days":
            return "Fewest days at school"
        y = c[5:]
        return f"{year_label(y)} preference" + ("" if self.rankings.get(y) else "  (none ranked)")

    def refresh_priority_list(self):
        for w in self.pri_box.winfo_children():
            w.destroy()
        for i, c in enumerate(self.priority):
            row = ctk.CTkFrame(self.pri_box, corner_radius=10, fg_color=CARD_ALT)
            row.pack(fill="x", pady=3)
            ctk.CTkLabel(row, text=str(i + 1), width=26, height=26, corner_radius=13, fg_color=ACCENT,
                         text_color="white", font=F(12, "bold")).pack(side="left", padx=(8, 8), pady=7)
            label(row, self.priority_text(c), 13).pack(side="left", fill="x", expand=True)
            mini_button(row, "▼", lambda i=i: self.move_priority(i, 1),
                        state="normal" if i < len(self.priority) - 1 else "disabled",
                        fg_color=CARD).pack(side="right", padx=(2, 8))
            mini_button(row, "▲", lambda i=i: self.move_priority(i, -1),
                        state="normal" if i > 0 else "disabled", fg_color=CARD).pack(side="right", padx=2)

    def move_priority(self, i, d):
        j = i + d
        if 0 <= j < len(self.priority):
            self.priority[i], self.priority[j] = self.priority[j], self.priority[i]
            self.after(10, self.refresh_priority_list)

    def preset(self, name):
        prefs = [c for c in self.priority if c.startswith("pref:")]
        self.priority = {"default": ["conflict", "days"] + prefs,
                         "days": ["days", "conflict"] + prefs,
                         "prefs": prefs + ["conflict", "days"]}[name]
        self.after(10, self.refresh_priority_list)

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
        for w in self.sub_scroll.winfo_children():
            w.destroy()
        self._comp_vars = {}
        for key in sorted(sections_of):
            var = tk.BooleanVar(value=key not in self.excluded)
            self._comp_vars[key] = var
            row = ctk.CTkFrame(self.sub_scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkCheckBox(row, text=comp_label(key), variable=var, font=F(13), text_color=TEXT,
                            checkbox_width=20, checkbox_height=20, corner_radius=6,
                            fg_color=ACCENT, hover_color=ACCENT_HOVER,
                            command=lambda k=key: self._toggle_component(k)).pack(side="left")
            label(row, f"{len(sections_of[key])} sec", 11, muted=True).pack(side="right", padx=(8, 4))

    def _toggle_component(self, key):
        if self._comp_vars[key].get():
            self.excluded.discard(key)
        else:
            self.excluded.add(key)

    def set_all_components(self, on):
        self.excluded = set() if on else set(self._comp_vars)
        for key, var in self._comp_vars.items():
            var.set(key not in self.excluded)

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
        for w in self.pref_scroll.winfo_children():
            w.destroy()
        if not by_year:
            label(self.pref_scroll, "Add some schedules first.", 13, muted=True).pack(anchor="w", padx=8, pady=8)
            return
        for y in sorted(by_year, key=year_sort_key):
            RankPicker(self.pref_scroll, y, by_year[y], self.rankings.get(y, []), self._on_rank_change
                       ).pack(fill="x", pady=(0, 10), padx=4)

    def _on_rank_change(self, year, ranked):
        if ranked:
            self.rankings[year] = ranked
        else:
            self.rankings.pop(year, None)
        self.refresh_priority_list()

    # ------------------------------------------------------------ results page
    def _build_results_page(self):
        page = self.pages["results"]
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)
        page_header(page, "Results", "Pick a schedule on the left to see its weekly timetable. "
                                     "Drag the dividers to resize the panels."
                    ).grid(row=0, column=0, sticky="w", pady=(0, 14))

        # draggable dividers: list | (timetable / details)
        self.hpane = tk.PanedWindow(page, orient="horizontal", sashwidth=14, sashrelief="flat",
                                    borderwidth=0, showhandle=False, opaqueresize=False,
                                    background=self._pane_bg())
        self.hpane.grid(row=1, column=0, sticky="nsew")

        lcard = card(self.hpane)
        self.res_tree = ttk.Treeview(lcard, show="headings", selectmode="browse", columns=("rank",))
        self.res_tree.heading("rank", text="#")
        sb = ctk.CTkScrollbar(lcard, command=self.res_tree.yview)
        self.res_tree.configure(yscrollcommand=sb.set)
        self.res_tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        sb.pack(side="right", fill="y", padx=(4, 6), pady=10)
        self.res_tree.bind("<<TreeviewSelect>>", self.on_result_select)

        right = ctk.CTkFrame(self.hpane, fg_color=BG)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)
        self.hpane.add(lcard, minsize=300, width=440, stretch="never")
        self.hpane.add(right, minsize=520, stretch="always")
        bar = ctk.CTkFrame(right, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.res_title = tk.StringVar(value="No results yet")
        ctk.CTkLabel(bar, textvariable=self.res_title, font=F(16, "bold"), text_color=TEXT).pack(side="left")
        for text, cmd in [("Export all to Excel…", lambda: self.export_excel(only_selected=False)),
                          ("Export selected to Excel…", lambda: self.export_excel(only_selected=True)),
                          ("Export CSV…", self.export_results_csv)]:
            button(bar, text, cmd, "ghost", height=32).pack(side="right", padx=(8, 0))

        self.vpane = tk.PanedWindow(right, orient="vertical", sashwidth=14, sashrelief="flat",
                                    borderwidth=0, showhandle=False, opaqueresize=False,
                                    background=self._pane_bg())
        self.vpane.grid(row=1, column=0, sticky="nsew")
        cvcard = card(self.vpane)
        self.cv = tk.Canvas(cvcard, highlightthickness=0, borderwidth=0)
        csb = ctk.CTkScrollbar(cvcard, command=self.cv.yview)
        self.cv.configure(yscrollcommand=csb.set)
        self.cv.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        csb.pack(side="right", fill="y", padx=(4, 6), pady=10)
        self.cv.bind("<Configure>", lambda e: self._schedule_redraw())
        self.cv.bind("<MouseWheel>", lambda e: self.cv.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        self.cv.bind("<Button-4>", lambda e: self.cv.yview_scroll(-1, "units"))
        self.cv.bind("<Button-5>", lambda e: self.cv.yview_scroll(1, "units"))

        self.detail = ctk.CTkTextbox(self.vpane, wrap="none", height=150, corner_radius=14, fg_color=CARD,
                                     border_width=1, border_color=BORDER, text_color=TEXT,
                                     font=ctk.CTkFont(family=self.mono, size=12), state="disabled")
        self.vpane.add(cvcard, minsize=180, stretch="always")
        self.vpane.add(self.detail, minsize=90, height=190, stretch="never")

    def _pane_bg(self):
        return BG[1] if ctk.get_appearance_mode() == "Dark" else BG[0]

    def _schedule_redraw(self):
        if self._redraw_job:
            self.after_cancel(self._redraw_job)
        self._redraw_job = self.after(120, self.draw_timetable)

    def draw_timetable(self):
        self._redraw_job = None
        cv = self.cv
        dark = ctk.get_appearance_mode() == "Dark"
        c_bg = "#1B1F2A" if dark else "#FFFFFF"
        c_grid = "#2E3445" if dark else "#E8EAF0"
        c_axis = "#8B93A7" if dark else "#6B7280"
        c_head = "#E5E7EB" if dark else "#111827"
        cv.configure(background=c_bg)
        cv.delete("all")
        r = self.current
        if not r:
            cv.create_text(max(cv.winfo_width(), 100) / 2, max(cv.winfo_height(), 60) / 2,
                           text="Your timetable will show up here", fill=c_axis, font=(self.ui_font, 12))
            return
        sc = self._scale()
        meetings = r["meetings"]
        days = DAY_ORDER[:5] + (["Sat"] if any(m[0] == "Sat" for m in meetings) else [])
        lo = min(m[1] for m in meetings) // 60 * 60
        hi = -(-max(m[2] for m in meetings) // 60) * 60
        left, top = round(66 * sc), round(34 * sc)
        pph = max(34 * sc, min(48 * sc, (cv.winfo_height() - top - 8) / max(1, (hi - lo) / 60)))
        width = max(cv.winfo_width(), 560)
        colw = (width - left - 6) / len(days)
        height = top + (hi - lo) / 60 * pph + 8

        def y_of(minute):
            return top + (minute - lo) / 60 * pph

        for h in range(lo // 60, hi // 60 + 1):
            y = y_of(h * 60)
            cv.create_line(left, y, width - 6, y, fill=c_grid)
            cv.create_text(left - 8, y, text=fmt(h * 60), anchor="e", fill=c_axis, font=(self.ui_font, 8))
        for j, day in enumerate(days):
            x = left + j * colw
            cv.create_line(x, top, x, y_of(hi), fill=c_grid)
            cv.create_text(x + colw / 2, top / 2, text=day, fill=c_head, font=(self.ui_font, 10, "bold"))
        cv.create_line(left + len(days) * colw, top, left + len(days) * colw, y_of(hi), fill=c_grid)

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
                    x0 = left + j * colw + li * colw / nl + 3
                    x1 = left + j * colw + (li + 1) * colw / nl - 3
                    y0, y1 = y_of(m[1]) + 2, y_of(m[2]) - 2
                    round_rect(cv, x0, y0, x1, y1, r=8, fill=color, outline="#DC2626" if clash else "")
                    h = y1 - y0
                    if h < 30 * sc:
                        lines = [f"{name} [{sec}]"]
                    elif h < 52 * sc:
                        lines = [name, f"Section {sec}"]
                    else:
                        lines = [name, f"Section {sec}", f"{fmt(m[1])} - {fmt(m[2])}"]
                    cv.create_text(x0 + 7, y0 + 5, text="\n".join(lines), anchor="nw", width=x1 - x0 - 12,
                                   fill=BLOCK_TEXT, font=(self.ui_font, 8, "bold" if clash else "normal"))
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
            vals = [r["rank"], f"{r['conflict']:.1f}", r["days"],
                    " ".join(d[:2] if d != "Thu" else "Th" for d in r["days_used"])]
            vals += [f"{r['prefs'][y][0]}/{r['prefs'][y][1]}" if y in r["prefs"] else "-" for y in years]
            self.res_tree.insert("", "end", iid=str(i), values=vals)
        if self.results:
            self.res_tree.selection_set("0")
        else:
            self.current = None
            self.res_title.set("No results yet")
            self.draw_timetable()

    # ----------------------------------------------------------------- search
    def _top_n(self):
        try:
            return max(1, min(500, int(self.top_n_var.get())))
        except ValueError:
            return None

    def start_search(self):
        if self.worker and self.worker.is_alive():
            return
        keys, comps = build_components(self.rows, self.excluded)
        if not comps:
            self.msg("Nothing to schedule",
                     "Add some class meetings (page 1) and tick at least one subject (page 2).", "warning")
            return
        top_n = self._top_n()
        if top_n is None:
            self.msg("Schedules to keep", "Enter a whole number for 'Schedules to keep'.", "warning")
            return
        max_days = None if self.max_days_var.get() == "No limit" else int(self.max_days_var.get())
        max_conf = None
        if self.max_conf_var.get().strip():
            try:
                max_conf = int(round(float(self.max_conf_var.get()) * 60))
                if max_conf < 0:
                    raise ValueError
            except ValueError:
                self.msg("Max conflict hours", "Enter a number such as 0 or 1.5 (or leave it blank).", "warning")
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
        self.progress.set(0)
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
        self.progress.set(min(1.0, max(0.0, state.get("progress", 0))))
        if not out.get("done"):
            self.status.set(f"Searching {self.ctx['total']:,} combinations…  {elapsed:.1f}s  ·  "
                            f"{state.get('nodes', 0):,} steps  ·  {state.get('progress', 0) * 100:.0f}%")
            self.after(150, self._poll)
            return
        self.run_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        if "error" in out:
            self.status.set("Search failed.")
            self.msg("Search failed", out["error"], "error")
            return
        ctx = self.ctx
        self.res_rankings = ctx["rankings"]
        self.results = []
        for rank, entry_ in enumerate(out["best"], 1):
            r = describe(entry_, ctx["keys"], ctx["comps"], ctx["rankings"])
            r["rank"] = rank
            self.results.append(r)
        self.fill_results()
        if not out["cancelled"]:
            self.progress.set(1)
        if not self.results:
            self.status.set("No schedule satisfies your hard limits.")
            self.msg("No schedules found",
                     "No combination satisfies your hard limits (max days / max conflict hours). "
                     "Try relaxing them.")
            return
        note = "  (stopped early - results may not be the very best)" if out["cancelled"] else ""
        self.status.set(f"Done in {elapsed:.2f}s · {len(self.results)} schedule(s) from "
                        f"{ctx['total']:,} combinations.{note}")
        self.show_page("results")

    # ----------------------------------------------------------------- export
    def export_excel(self, only_selected):
        if not self.results:
            self.msg("Nothing to export", "Run the search first.")
            return
        chosen = self.results
        if only_selected:
            sel = self.res_tree.selection()
            chosen = [self.results[int(sel[0])]] if sel else []
            if not chosen:
                self.msg("Nothing selected", "Select a schedule in the list first.")
                return
        path = filedialog.asksaveasfilename(title="Export to Excel", defaultextension=".xlsx",
                                            initialfile="schedules.xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            export_xlsx(chosen, path)
        except (RuntimeError, OSError) as ex:
            self.msg("Export failed", str(ex), "error")
            return
        self.status.set(f"Saved {len(chosen)} sheet(s) to {os.path.basename(path)}")

    def export_results_csv(self):
        if not self.results:
            self.msg("Nothing to export", "Run the search first.")
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
    ctk.set_appearance_mode("dark")          # "dark", "light" or "system"; the sidebar switch changes it live
    ctk.set_default_color_theme("blue")
    App().mainloop()


if __name__ == "__main__":
    main()
