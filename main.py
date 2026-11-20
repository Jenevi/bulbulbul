# -*- coding: utf-8 -*-
from dataclasses import dataclass
from tkinter import Tk, Frame, Label, Listbox, Entry, Button, Checkbutton, IntVar, StringVar, Radiobutton, Scrollbar, Toplevel, END, messagebox, PhotoImage, Canvas, filedialog
from tkinter.font import families, Font
import os
import tkinter as tk
import json
from pathlib import Path
import sys
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font as XLFont, Alignment
    OPENPYXL_AVAILABLE = True
except Exception:
    Workbook = None
    XLFont = None
    Alignment = None
    OPENPYXL_AVAILABLE = False
try:
    from PIL import Image, ImageTk, ImageChops
    PIL_AVAILABLE = True
except Exception:
    Image = None
    ImageTk = None
    ImageChops = None
    PIL_AVAILABLE = False

ui_text = {
    "ru": {
        "state1": "Состояние 1",
        "state2": "Состояние 2",
        "state3": "Состояние 3",
        "enter": "Ввести",
        "insert": "Вставить",
        "delete": "Удалить",
        "edit": "Правка",
        "controls": "Управление",
        "start_state": "Начальное состояние",
        "step_back": "Шаг назад",
        "clear": "Очистить",
        "seconds_input": "Ввод в сек",
        "enter_all": "Ввести все",
        "end_time_mss": "Конец (M.SS)",
        "end_time_sec": "Конец (сек)",
        "calculate": "Подсчёт",
        "total": "Итого",
        "global": "Общий",
        "lang_en": "English",
        "lang_ru": "Русский",
        "image_pos": "Позиция картинки",
    },
    "en": {
        "state1": "State 1",
        "state2": "State 2",
        "state3": "State 3",
        "enter": "Enter",
        "insert": "Insert",
        "delete": "Delete",
        "edit": "Edit",
        "controls": "Controls",
        "start_state": "Start state",
        "step_back": "Step back",
        "clear": "Clear",
        "seconds_input": "Input in seconds",
        "enter_all": "Enter all",
        "end_time_mss": "End time (M.SS)",
        "end_time_sec": "End time (sec)",
        "calculate": "Calculate",
        "total": "Total",
        "global": "Global",
        "lang_en": "English",
        "lang_ru": "Русский",
        "image_pos": "Image position",
    },
}


@dataclass
class Event:
    t_seconds: int
    state: str
    order: int


class BulBulApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulbulbul")
        self.root.geometry("1100x630")
        self.root.configure(bg="#e6e2dc")
        self.root.resizable(False, False)
        self.lang = "ru"
        try:
            fams = set(families())
        except Exception:
            fams = set()
        if "DejaVu Sans" in fams:
            self.title_font = ("DejaVu Sans", 14, "bold")
            self.ui_font = ("DejaVu Sans", 11)
        else:
            self.title_font = ("Arial", 14, "bold")
            self.ui_font = ("Arial", 11)
        self.states = [("A", "#c62828"), ("P", "#1565c0"), ("I", "#616161")]
        self.order_counter = 0
        self.undo_stack = []
        self.current_state = "A"
        self.initial_state = "A"
        self.initial_state_var = StringVar(value=self.initial_state)
        self.search_var = StringVar(value="")
        self.seconds_mode = IntVar(value=0)
        self.columns = {}
        self.totals = {"A": 0, "P": 0, "I": 0}
        self.search_dialog = None
        self.search_matches = []
        self.search_match_index = -1
        self.search_last_value_seconds = None
        self.search_img_tk = None
        self.initial_marker_img_tk = None
        self.copy_img_tk = None
        self.save_icon_tk = {}
        self.save_state = "unsaved"
        self.save_btn = None
        self._build_ui()

    def _build_ui(self):
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        banner_frame = Frame(self.root, bg="#e6e2dc")
        banner_frame.grid(row=0, column=0, sticky="ew")
        banner_frame.grid_propagate(False)
        banner_frame.grid_columnconfigure(0, weight=0)
        banner_frame.grid_columnconfigure(1, weight=1)
        banner_frame.grid_columnconfigure(2, weight=0)
        BANNER_H = 80
        banner_frame.configure(height=BANNER_H)
        self.banner_frame = banner_frame
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            title_path = os.path.join(base_dir, "assets", "название.png")
            if PIL_AVAILABLE:
                _img = Image.open(title_path)
                _img.thumbnail((2000, BANNER_H - 10), Image.LANCZOS)
                self.banner_img_tk = ImageTk.PhotoImage(_img, master=self.root)
            else:
                self.banner_img_tk = PhotoImage(file=title_path, master=self.root)
        except Exception:
            self.banner_img_tk = None
        self.save_btn = Button(banner_frame, text="Сохранить", command=self._save_results, font=self.ui_font, bd=1, relief="groove", padx=2, pady=2)
        self.save_btn.grid(row=0, column=0, sticky="nw", padx=(10, 0), pady=(0, 0))
        self._load_save_button_icons()
        self._set_save_state("unsaved")

        self.banner_label = Label(
            banner_frame,
            image=self.banner_img_tk if self.banner_img_tk is not None else None,
            bd=0,
            highlightthickness=0,
            bg=banner_frame.cget("background"),
        )
        self.banner_label.grid(row=0, column=1, sticky="nsew")

        self.main_container = Frame(self.root, bg="#e6e2dc")
        self.main_container.grid(row=1, column=0, sticky="nsew")
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(1, weight=0)
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(1, weight=0)
        self.main_container.grid_columnconfigure(2, weight=1)

        main = Frame(self.main_container, bg="#e6e2dc")
        main.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        self.cols_wrap = main

        for i in range(3):
            main.grid_columnconfigure(i, weight=0)
        main.grid_rowconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=0)
        for ci, (sid, color) in enumerate(self.states):
            col = Frame(main, bg="#e6e2dc")
            col.grid(row=0, column=ci, sticky="n", padx=(12 if ci > 0 else 4, 18))
            title_wrap = Frame(col, bg="#e6e2dc")
            title_wrap.grid(row=0, column=0, columnspan=2, sticky="n", pady=(0, 4))
            title_lbl = Label(title_wrap, text="", fg=color, bg="#e6e2dc", font=self.title_font, cursor="hand2")
            title_lbl.pack(side="left")
            marker_wrap = Frame(title_wrap, width=22, height=28, bg="#e6e2dc")
            marker_wrap.pack(side="left", padx=(2, 0))
            marker_wrap.pack_propagate(False)
            title_marker = Button(marker_wrap, text="", relief="flat", bd=0, highlightthickness=0, command=lambda s=sid: self._set_initial_state(s), font=self.ui_font, padx=0, pady=0)
            title_marker.pack(fill="both", expand=True)
            copy_btn = Button(title_wrap, text="⧉", relief="groove", bd=1, command=lambda s=sid: self.copy_column_to_clipboard(s), font=self.ui_font, padx=1, pady=1)
            copy_btn.pack(side="left", padx=(2, 0))
            title_lbl.bind("<Double-Button-1>", lambda e, s=sid: self._start_title_edit(s))
            scrollbar = Scrollbar(col, orient="vertical")
            listbox = Listbox(col, height=15, width=16, yscrollcommand=scrollbar.set, font=self.ui_font)
            scrollbar.config(command=listbox.yview)
            listbox.grid(row=1, column=0, sticky="n")
            scrollbar.grid(row=1, column=1, sticky="nsw", padx=(2, 0))
            total_lbl = Label(col, text="", bg="#e6e2dc", font=self.ui_font)
            total_lbl.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 2))
            entry = Entry(col, width=16, font=self.ui_font)
            entry.grid(row=3, column=0, columnspan=2, sticky="we", pady=(2, 2))
            btn = Button(col, text="", width=16, command=lambda s=sid: self.add_time_from_entry(s), font=self.ui_font)
            btn.grid(row=4, column=0, columnspan=2, sticky="we", pady=(2, 2))
            btns_frame = Frame(col, bg="#e6e2dc")
            btns_frame.grid(row=5, column=0, columnspan=2, sticky="we", pady=(2, 6))
            btns_frame.grid_columnconfigure(0, weight=1)
            btns_frame.grid_columnconfigure(1, weight=1)
            btns_frame.grid_columnconfigure(2, weight=1)
            insert_btn = Button(btns_frame, text="", command=lambda s=sid: self.insert_time_from_entry(s), font=self.ui_font)
            delete_btn = Button(btns_frame, text="", command=lambda s=sid: self.delete_selected(s), font=self.ui_font)
            edit_btn = Button(btns_frame, text="", command=lambda s=sid: self.edit_selected_from_entry(s), font=self.ui_font)
            insert_btn.grid(row=0, column=0, padx=0, sticky="we")
            delete_btn.grid(row=0, column=1, padx=0, sticky="we")
            edit_btn.grid(row=0, column=2, padx=0, sticky="we")
            entry.bind("<FocusIn>", lambda e, s=sid: self._focus_state(s))
            entry.bind("<Return>", lambda e, s=sid: (self.add_time_from_entry(s), "break")[1])
            listbox.bind("<FocusIn>", lambda e, s=sid: self._focus_state(s))
            listbox.bind("<Button-1>", lambda e, s=sid: self._focus_entry_from_listbox_click(e, s))
            listbox.bind("<Double-Button-1>", lambda e, s=sid: self.start_edit_from_listbox(e, s))
            listbox.bind("<Delete>", lambda e, s=sid: (self.delete_selected(s), "break")[1])
            self.columns[sid] = {
                "frame": col,
                "listbox": listbox,
                "scroll": scrollbar,
                "entry": entry,
                "total": total_lbl,
                "title": title_lbl,
                "title_wrap": title_wrap,
                "marker_wrap": marker_wrap,
                "initial_marker": title_marker,
                "copy_btn": copy_btn,
                "events": [],
                "enter_btn": btn,
                "insert_btn": insert_btn,
                "delete_btn": delete_btn,
                "edit_btn": edit_btn,
            }
        try:
            marker_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "assets",
                "сыр1.png"
            )
            if PIL_AVAILABLE:
                _img = Image.open(marker_path).convert("RGBA")
                if ImageChops is not None:
                    bg = Image.new("RGBA", _img.size, _img.getpixel((0, 0)))
                    diff = ImageChops.difference(_img, bg)
                    bbox = diff.getbbox()
                    if bbox:
                        _img = _img.crop(bbox)
                else:
                    bbox = _img.getbbox()
                    if bbox:
                        _img = _img.crop(bbox)
                _img.thumbnail((20, 26), Image.LANCZOS)
                self.initial_marker_img_tk = ImageTk.PhotoImage(_img, master=self.root)
            else:
                self.initial_marker_img_tk = PhotoImage(file=marker_path, master=self.root)
            for sid in ("A", "P", "I"):
                self.columns[sid]["initial_marker"].configure(image=self.initial_marker_img_tk, text="", compound="center", padx=2, pady=2)
            copy_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "assets",
                "копи.png"
            )
            if PIL_AVAILABLE:
                _img = Image.open(copy_path).convert("RGBA")
                bbox = _img.getbbox()
                if bbox:
                    _img = _img.crop(bbox)
                _img.thumbnail((22, 22), Image.LANCZOS)
                self.copy_img_tk = ImageTk.PhotoImage(_img, master=self.root)
            else:
                self.copy_img_tk = PhotoImage(file=copy_path, master=self.root)
            for sid in ("A", "P", "I"):
                self.columns[sid]["copy_btn"].configure(image=self.copy_img_tk, text="", compound="center", padx=1, pady=1)
        except Exception:
            self.initial_marker_img_tk = None
            self.copy_img_tk = None
        controls = Frame(self.main_container, bg="#e6e2dc")
        controls.grid(row=1, column=0, columnspan=3, sticky="ew", padx=8, pady=(8, 8))
        self.controls_frame = controls
        controls.grid_columnconfigure(0, weight=1)
        controls.grid_columnconfigure(1, weight=0)
        controls.grid_columnconfigure(2, weight=1)
        left_controls = Frame(controls, bg="#e6e2dc")
        left_controls.grid(row=0, column=0, sticky="w")
        right_controls = Frame(controls, bg="#e6e2dc")
        right_controls.grid(row=0, column=2, sticky="e")
        self.step_back_btn = Button(left_controls, text="", width=11, command=self.step_back, font=self.ui_font)
        self.step_back_btn.grid(row=0, column=0, padx=4, pady=2, sticky="w")
        self.clear_btn = Button(left_controls, text="", width=11, command=self.clear_all, font=self.ui_font)
        self.clear_btn.grid(row=0, column=1, padx=4, pady=2, sticky="w")
        self.chk = Checkbutton(left_controls, text="", variable=self.seconds_mode, bg="#e6e2dc", command=self._toggle_units, font=self.ui_font, anchor="w")
        self.chk.grid(row=0, column=2, padx=8, pady=2, sticky="w")
        self.enter_all_btn = Button(left_controls, text="", width=11, command=self.enter_all_from_clipboard, font=self.ui_font)
        self.enter_all_btn.grid(row=0, column=3, padx=4, pady=2, sticky="w")
        self.end_label = Label(left_controls, text="", bg="#e6e2dc", font=self.ui_font)
        self.end_label.grid(row=1, column=0, padx=(4, 4), pady=(2, 4), sticky="w")
        self.end_entry = Entry(left_controls, width=12, font=self.ui_font)
        self.end_entry.grid(row=1, column=1, padx=(0, 8), pady=(2, 4), sticky="w")
        self.end_entry.bind("<KeyRelease>", lambda e: self._mark_dirty())
        self.calc_btn = Button(left_controls, text="", width=11, command=self.calculate, font=self.ui_font)
        self.calc_btn.grid(row=1, column=2, padx=4, pady=(2, 4), sticky="w")
        self.global_total_lbl = Label(left_controls, text="", bg="#e6e2dc", font=self.ui_font)
        self.global_total_lbl.grid(row=1, column=3, padx=8, pady=(2, 4), sticky="w")
        self.search_btn = Button(
            right_controls,
            text="Поиск",
            command=self._open_search_dialog,
            font=self.ui_font,
            bd=1,
            relief="raised",
            padx=2,
            pady=2,
        )
        self.search_btn.grid(row=0, column=0, padx=(4, 6), pady=(2, 4), sticky="e")

        try:
            search_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "assets",
                "глазкриски.png"
            )
            if PIL_AVAILABLE:
                _img = Image.open(search_path).convert("RGBA")
                _img.thumbnail((28, 48), Image.LANCZOS)
                self.search_img_tk = ImageTk.PhotoImage(_img, master=self.root)
            else:
                self.search_img_tk = PhotoImage(file=search_path, master=self.root)

            self.search_btn.configure(
                image=self.search_img_tk,
                text="",
                compound="center",
                padx=1,
                pady=1,
            )
        except Exception:
            self.search_img_tk = None
        self.lang_btn = Button(right_controls, text="", width=12, command=self.toggle_language, font=self.ui_font)
        self.lang_btn.grid(row=0, column=1, padx=(4, 0), pady=(2, 4), sticky="e")
        self.image_pos_btn = Button(controls, text="", width=14, command=self.toggle_image_position_mode, font=self.ui_font)
        self.image_pos_btn.grid_remove()
        self.rat_img_pil = None
        self.rat_img_tk = None
        self.rat_label = None
        self.rat_x = 0
        self.rat_y = 0
        self.rat_target_max_h = 170
        self.rat_target_max_w = 220
        self.position_mode = False
        self._drag_active = False
        self._drag_offset_x = 0
        self._drag_offset_y = 0
        self._click_marker = None
        self._coord_label = None
        self.rat_x_rel = None
        self.rat_y_rel = None
        self._legacy_abs = None
        self._title_edit_entry = None
        self._title_edit_sid = None
        self._edit_mode = False
        self._edit_sid = None
        self._edit_index = None
        self.custom_titles = {}
        if sys.platform == "win32":
            config_root = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
        elif sys.platform == "darwin":
            config_root = Path.home() / "Library" / "Application Support"
        else:
            config_root = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
        self.ui_config_path = config_root / "bulbulbul" / "ui_config.json"
        self.legacy_config_path = Path(__file__).with_name("ui_config.json")
        self._load_custom_titles()
        self._load_initial_state()
        self.update_static_texts()
        self.root.bind_all("<Control-z>", lambda e: (self.step_back(), "break")[1], add="+")
        self.root.bind_all("<Control-f>", lambda e: (self._open_search_dialog(), "break")[1], add="+")
        self.root.bind_all("<Control-KeyPress>", self._on_ctrl_keypress, add="+")
        self.root.update_idletasks()
        req_w = max(self.root.winfo_width(), main.winfo_reqwidth() + 40)
        req_h = max(self.root.winfo_height(), main.winfo_reqheight() + 60)
        req_h += 80
        self.root.minsize(req_w, req_h)
        self.root.after(300, self._center_current_window_once)
        self._place_gap_decor()
        try:
            self._rat_placed_once = False
            self.root.after(300, self._apply_startup_rat_position_once)
            self.main_container.bind("<Configure>", lambda e: self._apply_saved_rat_position() if self._rat_placed_once else None, add="+")
        except Exception:
            pass
        try:
            self.root.bind("<Configure>", lambda e: self._ensure_rat_visible(), add="+")
        except Exception:
            pass

    def _place_gap_decor(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        img_path = os.path.join(base_dir, "assets", "rat.png")
        self.root.update_idletasks()
        if self.rat_img_tk is None:
            self._ensure_rat_image(img_path, self.rat_target_max_w, self.rat_target_max_h)
        saved = self._load_image_position()
        if self.rat_x_rel is not None and self.rat_y_rel is not None:
            return
        if self._legacy_abs is not None:
            return
        if not saved:
            self._set_default_rat_position()

    def _set_default_rat_position(self):
        self.root.update_idletasks()
        cw = max(1, self.main_container.winfo_width())
        ch = max(1, self.main_container.winfo_height())
        self.rat_x = max(0, cw - 300)
        self.rat_y = max(0, ch - 70)

    def _ensure_rat_label(self):
        if self.rat_img_tk is None:
            return None
        if self.rat_label is None:
            self.rat_label = tk.Label(self.main_container, image=self.rat_img_tk, bg="#e6e2dc", bd=0)
        else:
            try:
                self.rat_label.configure(image=self.rat_img_tk)
            except Exception:
                pass
        return self.rat_label

    def _apply_startup_rat_position_once(self):
        if self._rat_placed_once:
            return
        self.root.update_idletasks()
        if self.rat_x_rel is not None or self._legacy_abs is not None:
            self._apply_saved_rat_position()
        elif self.rat_img_tk is not None:
            self._set_default_rat_position()
            self._ensure_rat_label()
            self._place_rat_safely(self.rat_x, self.rat_y)
            self._rat_placed_once = True

    def _place_rat_from_rel(self):
        try:
            if self.rat_x_rel is None or self.rat_y_rel is None:
                return
            ow = self.main_container.winfo_width()
            oh = self.main_container.winfo_height()
            if ow <= 0 or oh <= 0:
                return
            x = int(self.rat_x_rel * float(ow))
            y = int(self.rat_y_rel * float(oh))
            self._place_rat_safely(x, y)
        except Exception:
            pass

    def _convert_legacy_abs(self):
        try:
            if self._legacy_abs is None:
                return
            ow = self.main_container.winfo_width()
            oh = self.main_container.winfo_height()
            if ow <= 0 or oh <= 0:
                return
            lx, ly = self._legacy_abs
            self._legacy_abs = None
            self.rat_x_rel = float(lx) / float(max(1, ow))
            self.rat_y_rel = float(ly) / float(max(1, oh))
            self._place_rat_safely(lx, ly)
        except Exception:
            pass

    def _apply_saved_rat_position(self):
        try:
            ow = self.main_container.winfo_width()
            oh = self.main_container.winfo_height()
            if ow < 10 or oh < 10:
                self.root.after(20, self._apply_saved_rat_position)
                return
            x = y = None
            if self.rat_x_rel is not None and self.rat_y_rel is not None:
                x = float(self.rat_x_rel) * float(ow)
                y = float(self.rat_y_rel) * float(oh)
            elif self._legacy_abs is not None:
                x, y = self._legacy_abs
            if x is None or y is None:
                self._set_default_rat_position()
                x = self.rat_x
                y = self.rat_y
            self._ensure_rat_label()
            if self.rat_label is not None:
                self.rat_x = int(x)
                self.rat_y = int(y)
                self.rat_label.place(x=self.rat_x, y=self.rat_y, anchor="center")
                self.rat_label.lift()
                self._rat_placed_once = True
        except Exception:
            pass

    def _ensure_rat_image(self, img_path, max_w, max_h, min_h=120):
        if PIL_AVAILABLE:
            try:
                pil_img = Image.open(img_path).convert("RGBA")
                w0, h0 = pil_img.size
                scale = min(max_w / w0, max_h / h0)
                new_w = max(1, int(w0 * scale))
                new_h = max(1, int(h0 * scale))
                if new_h < min_h:
                    adj = min_h / max(1, new_h)
                    new_w = max(1, int(new_w * adj))
                    new_h = min_h
                if (new_w, new_h) != (w0, h0):
                    try:
                        pil_img = pil_img.resize((new_w, new_h), Image.LANCZOS)
                    except Exception:
                        pil_img = pil_img.resize((new_w, new_h))
                self.rat_img_pil = pil_img
                self.rat_img_tk = ImageTk.PhotoImage(self.rat_img_pil, master=self.root)
                return
            except Exception:
                self.rat_img_tk = None
        if self.rat_img_tk is None:
            try:
                raw = PhotoImage(file=img_path, master=self.root)
                w0, h0 = raw.width(), raw.height()
                factor_w = max(1, (w0 + max_w - 1) // max_w)
                factor_h = max(1, (h0 + max_h - 1) // max_h)
                factor = max(factor_w, factor_h)
                if factor > 1:
                    raw = raw.subsample(factor, factor)
                self.rat_img_tk = raw
            except Exception:
                self.rat_img_tk = None

    def _apply_non_overlap_rules(self):
        if self.rat_img_tk is None:
            return
        c1 = self.columns["A"]["frame"]
        c2 = self.columns["P"]["frame"]
        lb1 = self.columns["A"]["listbox"]
        sb1 = self.columns["A"]["scroll"]
        lb2 = self.columns["P"]["listbox"]
        sb2 = self.columns["P"]["scroll"]
        def rect_of(frame, child):
            return (
                frame.winfo_x() + child.winfo_x(),
                frame.winfo_y() + child.winfo_y(),
                child.winfo_width(),
                child.winfo_height(),
            )
        forb = [
            rect_of(c1, lb1), rect_of(c1, sb1),
            rect_of(c2, lb2), rect_of(c2, sb2),
        ]
        def intersects(ax, ay, aw, ah, bx, by, bw, bh):
            return not (ax + aw <= bx or bx + bw <= ax or ay + ah <= by or by + bh <= ay)
        def any_overlap(x, y, w, h):
            for rx, ry, rw, rh in forb:
                if intersects(x, y, w, h, rx, ry, rw, rh):
                    return True
            return False
        x, y = self.rat_x, self.rat_y
        sizes = [self.rat_target_max_h, 150, 130, 120]
        sizes = [s for s in sizes if s >= 120]
        base_max_w = self.rat_target_max_w
        placed = False
        for target_h in sizes:
            self._ensure_rat_image(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "rat.png"), base_max_w, target_h, min_h=120)
            if self.rat_img_tk is None:
                continue
            w, h = self.rat_img_tk.width(), self.rat_img_tk.height()
            ok = not any_overlap(x, y, w, h)
            if not ok:
                found = False
                for dx in range(0, 61, 5):
                    if not any_overlap(x - dx, y, w, h):
                        x = x - dx
                        found = True
                        break
                if not found:
                    for dy in range(0, 61, 5):
                        if not any_overlap(x, y + dy, w, h):
                            y = y + dy
                            found = True
                            break
                ok = found
            if ok:
                placed = True
                self.rat_target_max_h = target_h
                break
        self.rat_x, self.rat_y = x, y
        if self.rat_label is None:
            self.rat_label = tk.Label(self.main_container, image=self.rat_img_tk, bg="#e6e2dc", bd=0)
        else:
            try:
                self.rat_label.configure(image=self.rat_img_tk)
            except Exception:
                pass
        self.rat_label.place(x=self.rat_x, y=self.rat_y, anchor="nw")
        self.rat_label.lift()

    def toggle_image_position_mode(self):
        self.position_mode = not self.position_mode
        if self.position_mode:
            label = self._ui_text("Позиция: ВКЛ" if self.lang == "ru" else "Position: ON")
            self.root.configure(cursor="crosshair")
            try:
                self.root.bind_all("<Button-1>", self._on_position_click_global, add="+")
            except Exception:
                pass
            self.image_pos_btn.configure(text=label, relief="sunken")
        else:
            self.root.configure(cursor="")
            try:
                self.root.unbind_all("<Button-1>")
            except Exception:
                pass
            self.image_pos_btn.configure(text=self._ui_text("image_pos"), relief="raised")
            try:
                if self._click_marker is not None:
                    self._click_marker.destroy()
            except Exception:
                pass
            self._click_marker = None
            try:
                if self._coord_label is not None:
                    self._coord_label.destroy()
            except Exception:
                pass
            self._coord_label = None
            try:
                ow = max(1, self.main_container.winfo_width())
                oh = max(1, self.main_container.winfo_height())
                self.rat_x_rel = float(self.rat_x) / float(ow)
                self.rat_y_rel = float(self.rat_y) / float(oh)
                self._save_image_position()
            except Exception:
                pass

    def _on_any_click(self, event):
        if not self.position_mode:
            return
        if not hasattr(self, "main_container"):
            return
        target = self.root.winfo_containing(event.x_root, event.y_root)
        if target is None:
            return
        if target is self.rat_label:
            return
        # Walk up to see if click is within overlay (main_container) and not on white panels
        cur = target
        inside_overlay = False
        white_hit = False
        while cur is not None:
            if cur == self.main_container:
                inside_overlay = True
                break
            try:
                bg = None
                try:
                    bg = cur.cget("bg")
                except Exception:
                    try:
                        bg = cur.cget("background")
                    except Exception:
                        bg = None
                if isinstance(bg, str) and bg.lower() in ("white", "#ffffff"):
                    white_hit = True
            except Exception:
                pass
            cur = getattr(cur, "master", None)
        if not inside_overlay or white_hit or self._is_in_white_panel(target):
            return
        rx = self.root.winfo_pointerx()
        ry = self.root.winfo_pointery()
        mcx = self.main_container.winfo_rootx()
        mcy = self.main_container.winfo_rooty()
        cx = rx - mcx
        cy = ry - mcy
        if self.rat_img_tk is None or self.rat_label is None:
            return
        w = self.rat_img_tk.width()
        h = self.rat_img_tk.height()
        center_x = self.rat_x + w / 2
        center_y = self.rat_y + h / 2
        self._drag_offset_x = center_x - cx
        self._drag_offset_y = center_y - cy
        self._drag_active = True
        new_center_x = cx + self._drag_offset_x
        new_center_y = cy + self._drag_offset_y
        new_x = int(new_center_x - w / 2)
        new_y = int(new_center_y - h / 2)
        self._place_rat_safely(new_x, new_y)
        self._ensure_rat_visible()
        self._drag_active = False

    def _on_position_click_global(self, event):
        if not self.position_mode or not hasattr(self, "main_container"):
            return
        w = self.root.winfo_containing(event.x_root, event.y_root)
        if w is None or (self.rat_label is not None and w is self.rat_label):
            return
        ox0 = self.main_container.winfo_rootx()
        oy0 = self.main_container.winfo_rooty()
        ow = self.main_container.winfo_width()
        oh = self.main_container.winfo_height()
        if ow <= 2 or oh <= 2:
            return
        ox1 = ox0 + ow
        oy1 = oy0 + oh
        xr = event.x_root
        yr = event.y_root
        if not (ox0 <= xr < ox1 and oy0 <= yr < oy1):
            return
        try:
            cur = w
            interactive = {"Button","TButton","Entry","TEntry","Text","Checkbutton","TCheckbutton","Combobox","TCombobox","Scrollbar","TScrollbar","Listbox","Spinbox"}
            while cur is not None:
                cls = cur.winfo_class()
                if cls in interactive:
                    return
                cur = getattr(cur,"master",None)
        except Exception:
            pass
        x = xr - ox0
        y = yr - oy0
        self._place_rat_safely(int(x), int(y))

    def _on_any_drag(self, event):
        if not self.position_mode or not self._drag_active:
            return
        if not hasattr(self, "main_container"):
            return
        target = self.root.winfo_containing(event.x_root, event.y_root)
        if target is None or target is self.rat_label:
            return
        cur = target
        inside_overlay = False
        white_hit = False
        while cur is not None:
            if cur == self.main_container:
                inside_overlay = True
                break
            try:
                bg = None
                try:
                    bg = cur.cget("bg")
                except Exception:
                    try:
                        bg = cur.cget("background")
                    except Exception:
                        bg = None
                if isinstance(bg, str) and bg.lower() in ("white", "#ffffff"):
                    white_hit = True
            except Exception:
                pass
            cur = getattr(cur, "master", None)
        if not inside_overlay or white_hit or self._is_in_white_panel(target):
            return
        rx = self.root.winfo_pointerx()
        ry = self.root.winfo_pointery()
        mcx = self.main_container.winfo_rootx()
        mcy = self.main_container.winfo_rooty()
        cx = rx - mcx
        cy = ry - mcy
        if self.rat_img_tk is None:
            return
        w = self.rat_img_tk.width()
        h = self.rat_img_tk.height()
        new_center_x = cx + self._drag_offset_x
        new_center_y = cy + self._drag_offset_y
        new_x = int(new_center_x - w / 2)
        new_y = int(new_center_y - h / 2)
        self._place_rat_safely(new_x, new_y)

    def _on_any_release(self, event):
        if not self.position_mode:
            return
        if self._drag_active:
            self._drag_active = False

    def _read_config(self):
        path = self.ui_config_path
        if not path.exists():
            path = self.legacy_config_path
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}
        return data if isinstance(data, dict) else {}

    def _write_config(self, data):
        self.ui_config_path.parent.mkdir(parents=True, exist_ok=True)
        self.ui_config_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    def _load_image_position(self):
        try:
            data = self._read_config()
            custom_titles = data.get("custom_titles")
            if isinstance(custom_titles, dict):
                self.custom_titles = {
                    sid: val for sid, val in custom_titles.items()
                    if sid in ("A", "P", "I") and isinstance(val, str) and val.strip()
                }
            pos_rel = data.get("rat_pos_rel")
            if isinstance(pos_rel, dict):
                try:
                    self.rat_x_rel = float(pos_rel.get("x_rel", 0.0))
                    self.rat_y_rel = float(pos_rel.get("y_rel", 0.0))
                except Exception:
                    self.rat_x_rel = self.rat_y_rel = None
                self.rat_target_max_h = int(data.get("rat_max_h", 170))
                return True
            pos = data.get("rat_pos")
            if isinstance(pos, dict):
                try:
                    self._legacy_abs = (int(pos.get("x", 0)), int(pos.get("y", 0)))
                except Exception:
                    self._legacy_abs = None
                self.rat_target_max_h = int(data.get("rat_max_h", 170))
                return True
        except Exception:
            return False
        return False

    def _load_custom_titles(self):
        try:
            data = self._read_config()
            custom_titles = data.get("custom_titles")
            if isinstance(custom_titles, dict):
                self.custom_titles = {
                    sid: val for sid, val in custom_titles.items()
                    if sid in ("A", "P", "I") and isinstance(val, str) and val.strip()
                }
        except Exception:
            self.custom_titles = {}

    def _load_initial_state(self):
        try:
            data = self._read_config()
            state = data.get("initial_state")
            if state in ("A", "P", "I"):
                self.initial_state = state
                self.initial_state_var.set(state)
        except Exception:
            pass

    def _load_save_button_icons(self):
        self.save_icon_tk = {}
        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_files = {
            "unsaved": ["разбитыйпиксель.png"],
            "saved": ["целый пиксель.png"],
        }
        for state, files in icon_files.items():
            for fname in files:
                path = os.path.join(base_dir, "assets", fname)
                try:
                    if PIL_AVAILABLE:
                        _img = Image.open(path).convert("RGBA")
                        bbox = _img.getbbox()
                        if bbox:
                            _img = _img.crop(bbox)
                        _img.thumbnail((33, 33), Image.LANCZOS)
                        self.save_icon_tk[state] = ImageTk.PhotoImage(_img, master=self.root)
                    else:
                        self.save_icon_tk[state] = PhotoImage(file=path, master=self.root)
                    break
                except Exception:
                    continue

    def _refresh_save_button_icon(self):
        if self.save_btn is None:
            return
        img = self.save_icon_tk.get(self.save_state)
        if img is not None:
            self.save_btn.configure(image=img, text="", compound="center", padx=2, pady=2)
        else:
            self.save_btn.configure(image="", text="Сохранить", compound="center", padx=2, pady=2)

    def _set_save_state(self, state):
        if state not in ("unsaved", "saved"):
            return
        self.save_state = state
        self._refresh_save_button_icon()

    def _mark_dirty(self):
        self._set_save_state("unsaved")

    def _save_image_position(self):
        try:
            data = self._read_config()
            ow = max(1, self.main_container.winfo_width())
            oh = max(1, self.main_container.winfo_height())
            self.rat_x_rel = float(self.rat_x) / float(ow)
            self.rat_y_rel = float(self.rat_y) / float(oh)
            if self.custom_titles:
                data["custom_titles"] = self.custom_titles
            data["initial_state"] = self.initial_state
            data["rat_pos_rel"] = {"x_rel": self.rat_x_rel, "y_rel": self.rat_y_rel}
            data["rat_max_h"] = int(self.rat_target_max_h)
            self._write_config(data)
        except Exception:
            pass

    def _save_custom_titles(self):
        try:
            data = self._read_config()
            if self.custom_titles:
                data["custom_titles"] = self.custom_titles
            else:
                data.pop("custom_titles", None)
            data["initial_state"] = self.initial_state
            self._write_config(data)
        except Exception:
            pass

    def _save_initial_state(self):
        try:
            data = self._read_config()
            data["initial_state"] = self.initial_state
            self._write_config(data)
        except Exception:
            pass

    def _save_rat_position_rel(self):
        self._save_image_position()

    def _is_in_white_panel(self, widget):
        try:
            cur = widget
            while cur is not None:
                cls = cur.winfo_class()
                if cls in ("Listbox", "Scrollbar", "TScrollbar"):
                    return True
                if cur == self.main_container or cur == self.root:
                    break
                cur = cur.master
        except Exception:
            return False
        return False

    def _is_descendant_of(self, w, ancestor):
        t = w
        while t is not None:
            if t is ancestor:
                return True
            t = getattr(t, "master", None)
        return False

    def _ensure_rat_visible(self):
        if self.rat_img_tk is None or self.rat_label is None:
            return
        cw = self.main_container.winfo_width()
        ch = self.main_container.winfo_height()
        w = self.rat_img_tk.width()
        h = self.rat_img_tk.height()
        if cw <= 0 or ch <= 0:
            return
        left = self.rat_x - w // 2
        top = self.rat_y - h // 2
        right = left + w
        bottom = top + h
        stuck_bottom = top >= max(0, ch - 2)
        oob = right < 0 or left > cw - 1 or bottom < 0 or top > ch - 1
        if stuck_bottom or oob:
            self._set_default_rat_position()
            self.rat_label.place(x=self.rat_x, y=self.rat_y, anchor="center")
            self.rat_label.lift()
        else:
            pass

    def _place_rat_safely(self, x, y):
        if self.rat_img_tk is None or self.rat_label is None:
            return
        self.rat_x, self.rat_y = x, y
        self.rat_label.place(x=self.rat_x, y=self.rat_y, anchor="center")
        self.rat_label.lift()

    def _on_configure(self, event):
        return

    def _center_on_screen(self, w, h):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _center_current_window_once(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _init_title_banner(self):
        return

    def _focus_state(self, sid):
        self.current_state = sid

    def _toggle_units(self):
        self._mark_dirty()
        self.update_static_texts()

    def _on_ctrl_keypress(self, event):
        ks = (getattr(event, "keysym", "") or "").lower()
        ch = (getattr(event, "char", "") or "").lower()
        if ks == "z" or ch in ("z", "я"):
            self.step_back()
            return "break"
        if ks == "f" or ch in ("f", "а"):
            self._open_search_dialog()
            return "break"

    def _on_initial_state_change(self):
        self._mark_dirty()
        self._set_initial_state(self.initial_state_var.get())

    def _set_initial_state(self, sid):
        if sid not in ("A", "P", "I"):
            return
        if sid != self.initial_state:
            self._mark_dirty()
        self.initial_state = sid
        self.initial_state_var.set(sid)
        self._save_initial_state()
        self._refresh_initial_state_markers()
        self._calculate(silent=True)

    def _refresh_initial_state_markers(self):
        for sid in ("A", "P", "I"):
            wrap = self.columns[sid].get("marker_wrap")
            marker = self.columns[sid].get("initial_marker")
            if wrap is not None:
                wrap.configure(relief="sunken" if sid == self.initial_state else "groove", bd=1)
            if marker is not None:
                if self.initial_marker_img_tk is not None:
                    if sid == self.initial_state:
                        marker.configure(image=self.initial_marker_img_tk, text="")
                    else:
                        marker.configure(image="", text="", relief="groove")
                else:
                    marker.configure(text="1" if sid == self.initial_state else "")

    def t(self, key):
        return ui_text[self.lang][key]

    def _fix_text(self, s):
        if isinstance(s, str) and "\\u" in s:
            try:
                return s.encode("utf-8").decode("unicode_escape")
            except Exception:
                return s
        return s

    def _ui_text(self, key_or_text):
        if isinstance(key_or_text, str) and key_or_text in ui_text.get(self.lang, {}):
            return self.t(key_or_text)
        return key_or_text

    def update_static_texts(self):
        for idx, sid in enumerate(("A", "P", "I")):
            title = self.custom_titles.get(sid) or [self.t("state1"), self.t("state2"), self.t("state3")][idx]
            self.columns[sid]["title"].configure(text=self._ui_text(title))
        for sid in ("A", "P", "I"):
            self.columns[sid]["enter_btn"].configure(text=self._ui_text("enter"))
            self.columns[sid]["insert_btn"].configure(text=self._ui_text("insert"))
            self.columns[sid]["delete_btn"].configure(text=self._ui_text("delete"))
            self.columns[sid]["edit_btn"].configure(text=self._ui_text("edit"))
        self.step_back_btn.configure(text=self._ui_text("step_back"))
        self.clear_btn.configure(text=self._ui_text("clear"))
        self.chk.configure(text=self._ui_text("seconds_input"))
        self.enter_all_btn.configure(text=self._ui_text("enter_all"))
        if self.seconds_mode.get() == 1:
            self.end_label.configure(text=self._ui_text("end_time_sec"))
        else:
            self.end_label.configure(text=self._ui_text("end_time_mss"))
        self.calc_btn.configure(text=self._ui_text("calculate"))
        self._calculate(silent=True)
        self._refresh_initial_state_markers()
        self.lang_btn.configure(text=self._ui_text(self.t("lang_en") if self.lang == "ru" else self.t("lang_ru")))
        self.image_pos_btn.configure(text=self._ui_text("image_pos"))
        if self.search_img_tk is None:
            self.search_btn.configure(text="🔍", width=10, height=4, compound="center")

    def _start_title_edit(self, sid):
        if self._title_edit_entry is not None:
            return
        lbl = self.columns[sid]["title"]
        text = lbl.cget("text")
        info = lbl.pack_info()
        lbl.pack_forget()
        entry = Entry(lbl.master, width=max(8, len(text) + 2), font=self.title_font)
        entry.insert(0, text)
        entry.select_range(0, END)
        entry.focus_set()
        entry.pack(side=info.get("side", "left"), padx=info.get("padx", 0), pady=info.get("pady", 0))
        entry.bind("<Return>", lambda e, s=sid: self._commit_title_edit(s))
        entry.bind("<FocusOut>", lambda e, s=sid: self._commit_title_edit(s))
        entry.bind("<Escape>", lambda e, s=sid: self._cancel_title_edit(s))
        self._title_edit_entry = entry
        self._title_edit_sid = sid

    def _normalize_search_query(self, text):
        s = text.strip()
        if not s:
            return None
        if self.seconds_mode.get() == 1:
            return self.parse_time(s)
        parsed = self._parse_mss_text(s)
        if not parsed:
            return None
        return parsed[0][0]

    def _collect_search_matches(self, value_seconds):
        matches = []
        for sid in ("A", "P", "I"):
            for idx, ev in enumerate(self.columns[sid]["events"]):
                if ev.t_seconds == value_seconds:
                    matches.append((sid, idx))
        return matches

    def _focus_search_match(self, sid, idx):
        lb = self.columns[sid]["listbox"]
        lb.selection_clear(0, END)
        lb.selection_set(idx)
        lb.see(idx)
        lb.focus_set()
        self._focus_state(sid)

    def copy_column_to_clipboard(self, sid):
        try:
            values = list(self.columns[sid]["listbox"].get(0, END))
            text = "\n".join(values)
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update()
            btn = self.columns[sid].get("copy_btn")
            if btn is not None:
                try:
                    btn.flash()
                except Exception:
                    pass
        except Exception:
            messagebox.showwarning("Copy", "Unable to copy column")

    def _get_export_data(self):
        if not self._calculate():
            return None
        titles = {sid: self.columns[sid]["title"].cget("text") for sid in ("A", "P", "I")}
        columns = {}
        for sid in ("A", "P", "I"):
            values = list(self.columns[sid]["listbox"].get(0, END))
            if sid == self.initial_state:
                values = ["0.00"] + values
            columns[sid] = values
        end_text = self.end_entry.get().strip()
        if end_text:
            try:
                end_seconds = self.parse_time(end_text)
                end_display = self.seconds_to_mdotss(end_seconds)
            except Exception:
                end_display = end_text
        else:
            end_display = ""
        totals = {}
        for sid in ("A", "P", "I"):
            lb_text = self.columns[sid]["total"].cget("text")
            totals[sid] = lb_text
        global_total = self.global_total_lbl.cget("text")
        return {
            "titles": titles,
            "columns": columns,
            "initial_state": self.initial_state,
            "end_time": end_display,
            "totals": totals,
            "global_total": global_total,
        }

    def _save_results(self):
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Сохранить результат",
            defaultextension=".txt",
            filetypes=[
                ("Text files", "*.txt"),
                ("Excel workbook", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        try:
            lower = path.lower()
            if lower.endswith(".xlsx"):
                self._export_xlsx(path)
            else:
                if not lower.endswith(".txt"):
                    path += ".txt"
                self._export_txt(path)
        except OSError as error:
            messagebox.showerror("Save", str(error), parent=self.root)

    def _export_txt(self, path):
        data = self._get_export_data()
        if data is None:
            return
        lines = []
        lines.append("Save export")
        lines.append(f"Initial state: {data['initial_state']}")
        lines.append(f"End time: {data['end_time']}")
        lines.append("")
        for sid in ("A", "P", "I"):
            lines.append(f"[{data['titles'][sid]}]")
            for value in data["columns"][sid]:
                lines.append(value)
            lines.append("")
        lines.append("[Totals]")
        lines.append(data["totals"]["A"])
        lines.append(data["totals"]["P"])
        lines.append(data["totals"]["I"])
        lines.append(data["global_total"])
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines).rstrip() + "\n")
        self._set_save_state("saved")

    def _export_xlsx(self, path):
        if not OPENPYXL_AVAILABLE:
            messagebox.showwarning("Save", "openpyxl is not available")
            return
        data = self._get_export_data()
        if data is None:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Results"
        hdr = XLFont(bold=True) if XLFont is not None else None
        center = Alignment(horizontal="center") if Alignment is not None else None
        ws["A1"] = "Initial state"
        ws["B1"] = data["initial_state"]
        ws["A2"] = "End time"
        ws["B2"] = data["end_time"]
        row = 4
        for idx, sid in enumerate(("A", "P", "I"), start=1):
            cell = ws.cell(row=row, column=idx, value=data["titles"][sid])
            cell.data_type = "s"
            if hdr is not None:
                cell.font = hdr
            if center is not None:
                cell.alignment = center
            for r, value in enumerate(data["columns"][sid], start=row + 1):
                ws.cell(row=r, column=idx, value=value)
        summary_row = row + max(len(data["columns"][sid]) for sid in ("A", "P", "I")) + 2
        ws.cell(row=summary_row, column=1, value="Totals")
        if hdr is not None:
            ws.cell(row=summary_row, column=1).font = hdr
        ws.cell(row=summary_row + 1, column=1, value="State 1")
        ws.cell(row=summary_row + 1, column=2, value=data["totals"]["A"])
        ws.cell(row=summary_row + 2, column=1, value="State 2")
        ws.cell(row=summary_row + 2, column=2, value=data["totals"]["P"])
        ws.cell(row=summary_row + 3, column=1, value="State 3")
        ws.cell(row=summary_row + 3, column=2, value=data["totals"]["I"])
        ws.cell(row=summary_row + 4, column=1, value="Global total")
        ws.cell(row=summary_row + 4, column=2, value=data["global_total"])
        for col in ("A", "B", "C"):
            ws.column_dimensions[col].auto_size = True
        for col in ("A", "B", "C"):
            max_len = 0
            for cell in ws[col]:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col].width = max(12, min(28, max_len + 2))
        wb.save(path)
        self._set_save_state("saved")

    def _perform_search(self, next_match=False):
        query = self.search_var.get()
        try:
            value_seconds = self._normalize_search_query(query)
        except ValueError as e:
            messagebox.showwarning("Search", str(e))
            return
        if value_seconds is None:
            messagebox.showinfo("Search", "Enter a time to search")
            return
        if value_seconds != self.search_last_value_seconds or not self.search_matches:
            self.search_matches = self._collect_search_matches(value_seconds)
            self.search_match_index = -1
            self.search_last_value_seconds = value_seconds
        if not self.search_matches:
            messagebox.showinfo("Search", "Nothing found")
            return
        self.search_match_index = (self.search_match_index + 1) % len(self.search_matches)
        sid, idx = self.search_matches[self.search_match_index]
        self._focus_search_match(sid, idx)

    def _open_search_dialog(self):
        if self.search_dialog is not None and self.search_dialog.winfo_exists():
            self.search_dialog.lift()
            self.search_dialog.focus_force()
            return
        dlg = Toplevel(self.root)
        dlg.title("Search")
        dlg.resizable(False, False)
        dlg.configure(bg="#e6e2dc")
        dlg.transient(self.root)
        self.search_dialog = dlg
        row = Frame(dlg, bg="#e6e2dc")
        row.pack(fill="both", expand=True, padx=6, pady=6)
        entry = Entry(row, textvariable=self.search_var, width=18, font=self.ui_font)
        entry.grid(row=0, column=0, padx=(0, 4), pady=2, sticky="we")
        btn = Button(row, text="Find Next", command=lambda: self._perform_search(next_match=True), font=self.ui_font)
        btn.grid(row=0, column=1, pady=2, sticky="e")
        row.grid_columnconfigure(0, weight=1)
        entry.bind("<Return>", lambda e: (self._perform_search(next_match=True), "break")[1])
        dlg.bind("<Return>", lambda e: (self._perform_search(next_match=True), "break")[1])
        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        self.root.update_idletasks()
        dlg.update_idletasks()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()
        rx = self.root.winfo_x()
        ry = self.root.winfo_y()
        dw = dlg.winfo_width()
        dh = dlg.winfo_height()
        x = rx + max(0, (rw - dw) // 2)
        y = ry + max(0, (rh - dh) // 2)
        dlg.geometry(f"{dw}x{dh}+{x}+{y}")
        entry.focus_set()
        entry.select_range(0, END)

    def _commit_title_edit(self, sid):
        entry = self._title_edit_entry
        if entry is None or self._title_edit_sid != sid:
            return "break"
        text = entry.get().strip()
        if text:
            self.custom_titles[sid] = text
        else:
            self.custom_titles.pop(sid, None)
        self._save_custom_titles()
        self._mark_dirty()
        self._title_edit_entry = None
        self._title_edit_sid = None
        entry.destroy()
        self.columns[sid]["title"].pack(side="left")
        self.update_static_texts()
        return "break"

    def _cancel_title_edit(self, sid):
        entry = self._title_edit_entry
        if entry is None or self._title_edit_sid != sid:
            return "break"
        self._title_edit_entry = None
        self._title_edit_sid = None
        entry.destroy()
        self.columns[sid]["title"].pack(side="left")
        self.update_static_texts()
        return "break"

    def parse_time(self, text):
        s = text.strip()
        if self.seconds_mode.get() == 1:
            if not s:
                raise ValueError("Enter seconds as a non-negative integer")
            if not s.isdecimal():
                raise ValueError("Seconds must be an integer number >= 0")
            return int(s)
        return self._parse_mss_token(s)[0]

    def _parse_mss_token(self, token):
        s = token.strip()
        if not s:
            raise ValueError("Empty time token")
        for sep in (".", ",", "/"):
            if sep in s:
                parts = s.split(sep)
                if len(parts) != 2:
                    raise ValueError(f"Invalid time token '{token}'")
                m_txt, sec_txt = parts[0].strip(), parts[1].strip()
                if not m_txt.isdecimal() or not sec_txt.isdecimal():
                    raise ValueError(f"Invalid time token '{token}'")
                if len(sec_txt) == 1:
                    sec_txt = "0" + sec_txt
                if len(sec_txt) != 2:
                    raise ValueError(f"Invalid time token '{token}'")
                minutes = int(m_txt)
                seconds = int(sec_txt)
                if seconds > 59:
                    raise ValueError(f"Invalid time token '{token}': seconds must be 00-59")
                return minutes * 60 + seconds, f"{minutes}.{seconds:02d}"
        return self._parse_compact_mss_token(s, token)

    def _parse_compact_mss_token(self, token, raw_token=None):
        # 5 -> 0.05
        # 45 -> 0.45
        # 007 -> 0.07
        # 123 -> 1.23
        digits = token
        if not digits.isdecimal():
            raise ValueError(f"Invalid time token '{raw_token or token}'")
        if len(digits) == 1:
            minutes = 0
            seconds = int(digits)
        else:
            minutes = int(digits[:-2] or "0")
            seconds = int(digits[-2:])
        if not (0 <= seconds <= 59):
            raise ValueError(f"Invalid time token '{raw_token or token}': seconds must be 00-59")
        return minutes * 60 + seconds, f"{minutes}.{seconds:02d}"

    def _parse_mss_text(self, text):
        tokens = [tok for tok in text.split() if tok.strip()]
        if not tokens:
            return []
        parsed = []
        for tok in tokens:
            t, canon = self._parse_mss_token(tok)
            parsed.append((t, canon))
        return parsed

    def _parse_mss_entry_tokens(self, sid, text):
        tokens = [tok for tok in text.split() if tok.strip()]
        if not tokens:
            return [], False
        parsed = []
        zero_seen = False
        for tok in tokens:
            t, canon = self._parse_mss_token(tok)
            if t == 0 and canon == "0.00":
                zero_seen = True
                continue
            parsed.append((t, canon))
        if zero_seen:
            self._set_initial_state(sid)
        return parsed, zero_seen

    def seconds_to_mdotss(self, t):
        m = t // 60
        s = t % 60
        return f"{m}.{s:02d}"

    def seconds_to_m_colon_ss(self, t):
        m = t // 60
        s = t % 60
        return f"{m}:{s:02d}"

    def add_time_from_entry(self, sid):
        entry = self.columns[sid]["entry"]
        txt = entry.get()
        if self._edit_mode and self._edit_sid == sid:
            self.edit_selected_from_entry(sid)
            return
        zero_seen = False
        try:
            if self.seconds_mode.get() == 1:
                t = self.parse_time(txt)
                times = [(t, self.seconds_to_mdotss(t))]
            else:
                times, zero_seen = self._parse_mss_entry_tokens(sid, txt)
        except ValueError as e:
            messagebox.showwarning("Invalid time", str(e))
            entry.focus_set()
            return
        if not times and not zero_seen:
            entry.focus_set()
            return
        if not times and zero_seen:
            self._mark_dirty()
            entry.delete(0, END)
            entry.focus_set()
            self._calculate(silent=True)
            return
        for t, _canon in times:
            self._add_event_at_index(sid, t, None, focus_entry=None)
        self._mark_dirty()
        entry.delete(0, END)
        entry.focus_set()
        self._calculate(silent=True)

    def _add_event_at_index(self, sid, t, index=None, focus_entry=None):
        ev = Event(t_seconds=t, state=sid, order=self.order_counter)
        self.order_counter += 1
        lb = self.columns[sid]["listbox"]
        ev_list = self.columns[sid]["events"]
        if index is None or index >= lb.size():
            lb.insert(END, self.seconds_to_mdotss(t))
            ev_list.append(ev)
            idx = lb.size() - 1
        else:
            lb.insert(index, self.seconds_to_mdotss(t))
            ev_list.insert(index, ev)
            idx = index
        self.undo_stack.append({"op": "add", "sid": sid, "index": idx})
        if focus_entry is not None:
            focus_entry.delete(0, END)
            focus_entry.focus_set()
        return True

    def insert_time_from_entry(self, sid):
        entry = self.columns[sid]["entry"]
        txt = entry.get()
        zero_seen = False
        try:
            if self.seconds_mode.get() == 1:
                t = self.parse_time(txt)
                times = [(t, self.seconds_to_mdotss(t))]
            else:
                times, zero_seen = self._parse_mss_entry_tokens(sid, txt)
        except ValueError as e:
            messagebox.showwarning("Invalid time", str(e))
            entry.focus_set()
            return
        if not times and not zero_seen:
            entry.focus_set()
            return
        if not times and zero_seen:
            lb = self.columns[sid]["listbox"]
            sel = lb.curselection()
            if sel:
                lb.selection_clear(0, END)
                lb.selection_set(sel[0])
            self._mark_dirty()
            entry.delete(0, END)
            entry.focus_set()
            self._calculate(silent=True)
            return
        lb = self.columns[sid]["listbox"]
        sel = lb.curselection()
        index = sel[0] if sel else None
        for offset, (t, _canon) in enumerate(times):
            self._add_event_at_index(sid, t, index + offset if index is not None else None, focus_entry=None)
        if index is None:
            lb.selection_clear(0, END)
            lb.selection_set(lb.size() - 1)
        else:
            lb.selection_clear(0, END)
            lb.selection_set(index)
        entry.delete(0, END)
        entry.focus_set()
        self._mark_dirty()
        self._calculate(silent=True)

    def delete_selected(self, sid):
        lb = self.columns[sid]["listbox"]
        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("Delete", "Select a row to delete")
            return
        idx = sel[0]
        ev_list = self.columns[sid]["events"]
        if not (0 <= idx < len(ev_list)):
            return
        ev = ev_list.pop(idx)
        lb.delete(idx)
        self.undo_stack.append({"op": "delete", "sid": sid, "index": idx, "event": ev})
        self._mark_dirty()
        self._calculate(silent=True)

    def edit_selected_from_entry(self, sid):
        lb = self.columns[sid]["listbox"]
        idx = self._edit_index if self._edit_mode and self._edit_sid == sid else None
        if idx is None:
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("Edit", "Select a row to edit")
                return
            idx = sel[0]
        if not (0 <= idx < lb.size()):
            messagebox.showwarning("Edit", "Select a row to edit")
            return
        entry = self.columns[sid]["entry"]
        txt = entry.get()
        try:
            t = self.parse_time(txt)
        except ValueError as e:
            messagebox.showwarning("Invalid time", str(e))
            entry.focus_set()
            return
        ev_list = self.columns[sid]["events"]
        if not (0 <= idx < len(ev_list)):
            return
        old_t = ev_list[idx].t_seconds
        ev_list[idx].t_seconds = t
        lb.delete(idx)
        lb.insert(idx, self.seconds_to_mdotss(t))
        lb.selection_clear(0, END)
        lb.selection_set(idx)
        self._edit_mode = False
        self._edit_sid = None
        self._edit_index = None
        self.undo_stack.append({"op": "edit", "sid": sid, "index": idx, "old_t": old_t, "new_t": t})
        entry.delete(0, END)
        entry.focus_set()
        self._mark_dirty()
        self._calculate(silent=True)

    def start_edit_from_listbox(self, event, sid):
        lb = self.columns[sid]["listbox"]
        idxs = lb.curselection()
        if not idxs:
            try:
                idx = lb.nearest(event.y)
            except Exception:
                return
        else:
            idx = idxs[0]
        events = self.columns[sid]["events"]
        if not 0 <= idx < len(events):
            return
        value = str(events[idx].t_seconds) if self.seconds_mode.get() else lb.get(idx)
        lb.selection_clear(0, END)
        lb.selection_set(idx)
        self._edit_mode = True
        self._edit_sid = sid
        self._edit_index = idx
        entry = self.columns[sid]["entry"]
        entry.delete(0, END)
        entry.insert(0, value)
        entry.focus_set()
        entry.selection_range(0, END)

    def copy_row_to_entry(self, event, sid):
        self.start_edit_from_listbox(event, sid)

    def _focus_entry_from_listbox_click(self, event, sid):
        lb = self.columns[sid]["listbox"]
        entry = self.columns[sid]["entry"]
        if lb.size() == 0:
            lb.selection_clear(0, END)
            self._focus_state(sid)
            entry.focus_set()
            entry.icursor(END)
            return "break"
        idx = lb.nearest(event.y)
        bbox = lb.bbox(idx)
        if not bbox:
            lb.selection_clear(0, END)
            self._focus_state(sid)
            entry.focus_set()
            entry.icursor(END)
            return "break"
        x, y, w, h = bbox
        if not (y <= event.y < y + h):
            lb.selection_clear(0, END)
            self._focus_state(sid)
            entry.focus_set()
            entry.icursor(END)
            return "break"
        self._focus_state(sid)
        return None

    def step_back(self):
        if not self.undo_stack:
            return
        op = self.undo_stack.pop()
        kind = op.get("op")
        sid = op.get("sid")
        lb = self.columns[sid]["listbox"]
        ev_list = self.columns[sid]["events"]
        if kind == "add":
            idx = op.get("index")
            if 0 <= idx < lb.size():
                lb.delete(idx)
            if 0 <= idx < len(ev_list):
                ev_list.pop(idx)
        elif kind == "delete":
            idx = op.get("index")
            ev = op.get("event")
            if idx is None:
                idx = lb.size()
            if idx > lb.size():
                idx = lb.size()
            lb.insert(idx, self.seconds_to_mdotss(ev.t_seconds))
            ev_list.insert(idx, ev)
            lb.selection_clear(0, END)
            lb.selection_set(idx)
        elif kind == "edit":
            idx = op.get("index")
            old_t = op.get("old_t")
            if 0 <= idx < len(ev_list):
                ev_list[idx].t_seconds = old_t
                lb.delete(idx)
                lb.insert(idx, self.seconds_to_mdotss(old_t))
                lb.selection_clear(0, END)
                lb.selection_set(idx)
        self._mark_dirty()
        self._calculate(silent=True)

    def clear_all(self):
        for data in self.columns.values():
            data["listbox"].delete(0, END)
            data["events"] = []
        self.undo_stack = []
        self.order_counter = 0
        self._edit_mode = False
        self._edit_sid = None
        self._edit_index = None
        self._calculate(silent=True)
        self._set_save_state("unsaved")

    def get_all_events(self):
        all_ev = []
        for sid in ("A", "P", "I"):
            all_ev.extend(self.columns[sid]["events"])
        return all_ev

    def _calculate(self, silent=False):
        events = sorted(self.get_all_events(), key=lambda event: (event.t_seconds, event.order))
        latest = events[-1].t_seconds if events else 0
        end_text = self.end_entry.get().strip()
        end_time = latest
        if end_text:
            try:
                end_time = self.parse_time(end_text)
                if end_time < latest:
                    raise ValueError("End time is smaller than the latest entered time")
            except ValueError as error:
                if not silent:
                    messagebox.showwarning("Invalid end time", str(error), parent=self.root)
                    self.end_entry.focus_set()
                    return False
                end_time = latest

        totals = {"A": 0, "P": 0, "I": 0}
        previous_time = 0
        state = self.initial_state
        for event in events:
            totals[state] += event.t_seconds - previous_time
            previous_time = event.t_seconds
            state = event.state
        totals[state] += end_time - previous_time
        for sid in ("A", "P", "I"):
            s = totals[sid]
            self.columns[sid]["total"].configure(text=self._ui_text(f"{self.t('total')}: {s} s ({self.seconds_to_m_colon_ss(s)})"))
        self.global_total_lbl.configure(text=self._ui_text(f"{self.t('global')}: {end_time} s ({self.seconds_to_m_colon_ss(end_time)})"))
        return True

    def calculate(self):
        self._calculate(silent=False)

    def toggle_language(self):
        self.lang = "en" if self.lang == "ru" else "ru"
        self.update_static_texts()

    def enter_all_from_clipboard(self):
        try:
            text = self.root.clipboard_get()
        except Exception:
            messagebox.showwarning("Clipboard", "Clipboard is empty or not accessible")
            return
        pending = []
        for line_no, raw in enumerate(text.splitlines(), start=1):
            line = raw.strip()
            if not line:
                continue
            sid = self.current_state
            if line[:2].upper() in ("A:", "P:", "I:"):
                sid = line[0].upper()
                line = line[2:]
            try:
                if self.seconds_mode.get() == 1:
                    times = [self.parse_time(token) for token in line.split()]
                else:
                    times = [time for time, _ in self._parse_mss_text(line)]
            except ValueError as error:
                messagebox.showwarning("Invalid time", f"Line {line_no}: {error}", parent=self.root)
                self.columns[sid]["entry"].focus_set()
                return
            pending.extend((sid, time) for time in times)
        for sid, time in pending:
            self._add_event_at_index(sid, time)
        if pending:
            self._mark_dirty()
        self._calculate(silent=True)


if __name__ == "__main__":
    root = Tk()
    app = BulBulApp(root)
    root.mainloop()
