import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser, font as tkfont
import json
import os
import ctypes
from datetime import datetime, date, timedelta
import threading
from PIL import Image, ImageDraw
import pystray
import logging

# ── Logging Configuration ──────────────────────────────────────────────────────
logging.basicConfig(
    filename='app.log',
    level=logging.ERROR,  # 僅在發生 Exception 時記錄
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    encoding='utf-8'
)
logger = logging.getLogger(__name__)

# ── Windows API ────────────────────────────────────────────────────────────────
WS_EX_TRANSPARENT = 0x00000020
WS_EX_LAYERED     = 0x00080000
GWL_EXSTYLE       = -20
HWND_BOTTOM       = 1
HWND_TOPMOST      = -1
SWP_NOSIZE        = 0x0001
SWP_NOMOVE        = 0x0002
SWP_NOACTIVATE    = 0x0010
user32            = ctypes.windll.user32


# ── Note format helper ────────────────────────────────────────────────────────
def parse_note_lines(note):
    """
    將不同版本的記事本格式統一轉換為單行字典串列 (List of dicts)。
    
    每個字典代表一行文字及其樣式設定，包含以下鍵值：
    - text: 該行的純文字內容
    - bold: 是否粗體 (布林值)
    - size: 字體大小 ("小", "中", "大")
    - align: 文字對齊 ("left", "center", "right")
    - color: 文字顏色 (Hex 色碼，例如 "#ffffff")
    
    支援解析的舊有格式包含：
    - 純字串 (str)
    - 舊版單一格式字典 (dict with "text" key)
    - 新版多行格式字典 (dict with "lines" key)
    """
    d = {"bold": False, "size": "中", "align": "left", "color": "#ffffff"}
    if isinstance(note, dict):
        if "lines" in note:
            return [{**d, **ln} for ln in note["lines"]]
        text = note.get("text", "")
        fmt  = {k: note.get(k, d[k]) for k in d}
        return [{**fmt, "text": ln} for ln in text.splitlines() if ln.strip()]
    if isinstance(note, str):
        return [{**d, "text": ln} for ln in note.splitlines() if ln.strip()]
    return []


# ── Rich Note Dialog ───────────────────────────────────────────────────────────
class NoteDialog(tk.Toplevel):
    """
    提供豐富文字格式編輯的獨立對話方塊。
    允許使用者針對逐行文字設定：
    - 字體大小 (小/中/大)
    - 粗體切換
    - 對齊方式 (左/置中/右)
    - 文字顏色設定
    具有懸浮置頂特性，修改完成後可立即儲存回主日曆。
    """
    SIZES = {"小": 9, "中": 11, "大": 14}

    def __init__(self, parent, date_str, current_notes, theme):
        """
        初始化編輯對話方塊。
        
        Args:
            parent: 父視窗物件 (通常為 main application root)
            date_str: 目前正在編輯的日期字串 (例如 "2024-03-12")
            current_notes: 當天既有的記事資料 (字串或字典)
            theme: 目前的主題配色字典
        """
        super().__init__(parent)
        self.result  = None
        self.bg      = theme.get("bg", "#3c7ea1")
        self.hbg     = theme.get("header_bg", "#4a8cb0")
        self._bold   = False
        self._size   = "中"
        self._align  = "left"
        self._color  = "#ffffff"

        self.title(f"記事 – {date_str}")
        self.geometry("480x460")
        self.minsize(360, 340)
        self.resizable(True, True)
        self.configure(bg=self.hbg)
        self.grab_set()
        self.lift()

        # ── Header ────────────────────────────────────────────────────────────
        tk.Label(self, text=f"📅  {date_str}",
                 font=("Microsoft JhengHei", 12, "bold"),
                 bg=self.hbg, fg="white").pack(pady=(10, 2))
        tk.Label(self, text="💡 選取文字後點格式按鈕套用，或先設定後再輸入",
                 font=("Arial", 8), bg=self.hbg, fg="#c8e6f5").pack()

        # ── Formatting toolbar ─────────────────────────────────────────────────
        toolbar = tk.Frame(self, bg=self.hbg)
        toolbar.pack(fill=tk.X, padx=12, pady=(6, 2))

        tk.Label(toolbar, text="大小:", bg=self.hbg, fg="white",
                 font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 4))
        self.size_btns = {}
        for sz in ["小", "中", "大"]:
            b = tk.Button(toolbar, text=sz,
                          bg="#27ae60" if sz == self._size else self.bg,
                          fg="white", bd=0, padx=7, pady=2,
                          font=("Microsoft JhengHei", 9),
                          command=lambda s=sz: self._set_size(s))
            b.pack(side=tk.LEFT, padx=1)
            self.size_btns[sz] = b

        tk.Label(toolbar, text=" | ", bg=self.hbg, fg="#a0c4d9").pack(side=tk.LEFT)
        self.bold_btn = tk.Button(toolbar, text="粗體 B",
                                  bg="#27ae60" if self._bold else self.bg,
                                  fg="white", bd=0, padx=7, pady=2,
                                  font=("Arial", 9, "bold"),
                                  command=self._toggle_bold)
        self.bold_btn.pack(side=tk.LEFT, padx=2)

        tk.Label(toolbar, text=" | ", bg=self.hbg, fg="#a0c4d9").pack(side=tk.LEFT)
        self.align_btns = {}
        for label, align in [("≡左", "left"), ("≡中", "center"), ("≡右", "right")]:
            b = tk.Button(toolbar, text=label,
                          bg="#27ae60" if align == self._align else self.bg,
                          fg="white", bd=0, padx=7, pady=2,
                          font=("Arial", 9),
                          command=lambda a=align: self._set_align(a))
            b.pack(side=tk.LEFT, padx=1)
            self.align_btns[align] = b

        tk.Label(toolbar, text=" | ", bg=self.hbg, fg="#a0c4d9").pack(side=tk.LEFT)
        tk.Label(toolbar, text="色:", bg=self.hbg, fg="white",
                 font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 3))
        # Color swatch button — shows current color, opens picker on click
        self.color_swatch_btn = tk.Button(
            toolbar, text="  ", width=2,
            bg=self._color, relief="solid", bd=1,
            command=self._pick_color)
        self.color_swatch_btn.pack(side=tk.LEFT, padx=2)

        # ── Bottom buttons (packed BEFORE text so always visible) ──────────────
        btn_bar = tk.Frame(self, bg=self.hbg)
        btn_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=12, pady=10)
        tk.Button(btn_bar, text="🗑 清除",
                  command=lambda: self.text.delete("1.0", tk.END),
                  bg="#7a3c3c", fg="white", font=("Microsoft JhengHei", 10),
                  bd=0, padx=10, pady=5).pack(side=tk.LEFT)
        tk.Button(btn_bar, text="✕ 取消", command=self.destroy,
                  bg="#4a8cb0", fg="white", font=("Microsoft JhengHei", 10),
                  bd=0, padx=10, pady=5).pack(side=tk.RIGHT, padx=(6, 0))
        tk.Button(btn_bar, text="✔ 儲存", command=self.on_save,
                  bg="#27ae60", fg="white", font=("Microsoft JhengHei", 11, "bold"),
                  bd=0, padx=16, pady=6).pack(side=tk.RIGHT)

        # ── Text area (主要文字編輯區) ────────────────────────────────────────────────────────
        text_frame = tk.Frame(self, bg=self.hbg)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 0))
        sb = tk.Scrollbar(text_frame)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        # 啟用 undo 功能，支援復原操作
        self.text = tk.Text(text_frame, font=("Microsoft JhengHei", 11),
                            wrap=tk.WORD, bg=self.bg, fg="white",
                            insertbackground="white", relief=tk.FLAT,
                            padx=8, pady=8, undo=True, yscrollcommand=sb.set)
        sb.config(command=self.text.yview)
        self.text.pack(fill=tk.BOTH, expand=True)

        # 設定 Text Widget 標籤 (Tags) 以支援多重樣式渲染
        self._configure_tags()
        # 載入目前的記事內容，並逐行套用對應標籤
        self._load_content(current_notes)

        self.text.focus_set()
        # 綁定 ctrl+enter 快速鍵直接儲存
        self.text.bind("<Control-Return>", lambda e: self.on_save())
        # 監聽游標移動與點擊，同步更新工具列的按鈕狀態顯示
        self.text.bind("<ButtonRelease-1>", lambda e: self._sync_toolbar_to_cursor())
        self.text.bind("<KeyRelease>",      lambda e: self._sync_toolbar_to_cursor())
        
        # 覆寫內建的 insert 方法，確保使用者輸入新文字時，會自動繼承當前游標的樣式標籤
        self._orig_insert = self.text.insert
        self.text.insert = self._tagged_insert

        # 點擊對話框關閉 (X) 時的行為
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        # 設定為視窗模式並強制鎖定焦點直到關閉
        self.transient(parent)
        self.wait_window()

    def _configure_tags(self):
        """
        設定 Tkinter Text widget 內建的樣式標籤 (Tags)。
        這些標籤之後可套用至特定文字段落，用以即時改變字體大小、粗細或對齊方式。
        """
        for name, size in self.SIZES.items():
            self.text.tag_configure(f"size_{name}",
                                    font=("Microsoft JhengHei", size))
            self.text.tag_configure(f"size_{name}_bold",
                                    font=("Microsoft JhengHei", size, "bold"))
        for align in ("left", "center", "right"):
            self.text.tag_configure(f"align_{align}", justify=align)

    def _ensure_color_tag(self, color):
        """
        (動態生成) 確保文字顏色標籤存在。
        因為自訂顏色的數量不可全數預先建立，所以當使用者選擇新顏色時，才動態生成該標籤。
        """
        tag = f"color_{color.lstrip('#')}"
        try:
            self.text.tag_cget(tag, "foreground")   # raises if tag absent
        except tk.TclError:
            self.text.tag_configure(tag, foreground=color)
        return tag

    def _load_content(self, note):
        """
        載入既有的記事資料並反序列化渲染。
        將串列中的每一行文字，逐行插入 Text 區塊，並重新套用其對應的大小、粗體、對齊與顏色標籤。
        """
        lines = parse_note_lines(note)
        for i, ln in enumerate(lines):
            if i > 0:
                self.text.insert(tk.END, "\n")
            start = self.text.index(f"{tk.END}-1c")
            self.text.insert(tk.END, ln["text"])
            end = self.text.index(f"{tk.END}-1c")
            
            # 組裝該行文字所需的標籤名稱
            sz_tag  = f"size_{ln['size']}" + ("_bold" if ln["bold"] else "")
            col_tag = self._ensure_color_tag(ln.get("color", "#ffffff"))
            
            self.text.tag_add(sz_tag,                   start, end)
            self.text.tag_add(f"align_{ln['align']}",  start, end)
            self.text.tag_add(col_tag,                  start, end)
        
        if not lines:
            # 如果是空記事 (新建)，直接為第一行套用預設款式
            col_tag = self._ensure_color_tag("#ffffff")
            self.text.tag_add("size_中",    "1.0", tk.END)
            self.text.tag_add("align_left", "1.0", tk.END)
            self.text.tag_add(col_tag,      "1.0", tk.END)

    def _get_line_format(self, line_num):
        """
        反向解析單行格式。
        給定一個以 1 為基準的行號，掃描該行具備的所有標籤，並還原成字典結構。
        """
        s = f"{line_num}.0"
        e = f"{line_num}.end"
        size, bold, align, color = "中", False, "left", "#ffffff"
        
        for name in self.SIZES:
            if self.text.tag_nextrange(f"size_{name}_bold", s, e):
                size, bold = name, True
                break
            if self.text.tag_nextrange(f"size_{name}", s, e):
                size = name
                break
                
        for a in ("left", "center", "right"):
            if self.text.tag_nextrange(f"align_{a}", s, e):
                align = a
                break
                
        # 偵測顏色標籤 (格式: color_RRGGBB)
        for tag in self.text.tag_names(s):
            if tag.startswith("color_"):
                color = "#" + tag[6:]
                break
        return {"bold": bold, "size": size, "align": align, "color": color}

    def _sync_toolbar_to_cursor(self):
        """
        工具列連動機制。
        當滑鼠點擊或鍵盤移動到另一行時，讀取該行的樣式並自動切換上方工具列的按鈕高亮狀態。
        """
        cursor = self.text.index(tk.INSERT)
        line_num = int(cursor.split(".")[0])
        fmt = self._get_line_format(line_num)
        
        self._bold  = fmt["bold"]
        self._size  = fmt["size"]
        self._align = fmt["align"]
        self._color = fmt["color"]
        
        for s, b in self.size_btns.items():
            b.config(bg="#27ae60" if s == self._size else self.bg)
        self.bold_btn.config(bg="#27ae60" if self._bold else self.bg)
        for a, b in self.align_btns.items():
            b.config(bg="#27ae60" if a == self._align else self.bg)
        self.color_swatch_btn.config(bg=self._color)

    def _current_size_tag(self):
        return f"size_{self._size}" + ("_bold" if self._bold else "")

    def _tagged_insert(self, index, chars, *args):
        """
        覆寫 Text 元件的預設插入行為。
        當使用者敲擊鍵盤輸入新字元時，確保新字串會自動繼承當前游標所在處的樣式標籤 (粗體/大小/顏色等)。
        """
        self._orig_insert(index, chars, *args)
        
        # 定位剛剛插入的文字範圍
        end = self.text.index(f"{index}+{len(chars)}c")
        
        # 準備要套用的當下樣式標籤
        size_tag  = self._current_size_tag()
        align_tag = f"align_{self._align}"
        col_tag   = self._ensure_color_tag(self._color)
        
        # 套用新標籤前，先清除該範圍內舊有的衝突標籤
        for n in self.SIZES:
            self.text.tag_remove(f"size_{n}",      index, end)
            self.text.tag_remove(f"size_{n}_bold", index, end)
        for a in ("left", "center", "right"):
            self.text.tag_remove(f"align_{a}", index, end)
        for tag in list(self.text.tag_names(index)):
            if tag.startswith("color_"):
                self.text.tag_remove(tag, index, end)
                
        # 掛上新的標籤
        self.text.tag_add(size_tag,  index, end)
        self.text.tag_add(align_tag, index, end)
        self.text.tag_add(col_tag,   index, end)

    def _apply_to_selection(self, size_tag=None, align_tag=None, col_tag=None):
        """
        將指定的格式標籤套用到目前使用者框選的文字範圍。
        如果使用者沒有框選任何文字，則針對游標所在的「整行」進行格式套用。
        """
        try:
            s = self.text.index(tk.SEL_FIRST)
            e = self.text.index(tk.SEL_LAST)
        except tk.TclError:
            # 發生例外代表目前無選取範圍 → 自動鎖定游標所在的一整行
            cursor   = self.text.index(tk.INSERT)
            line_num = cursor.split(".")[0]
            s = f"{line_num}.0"
            e = f"{line_num}.end"

        if size_tag:
            # 清除舊尺寸，套用新尺寸
            for n in self.SIZES:
                self.text.tag_remove(f"size_{n}",      s, e)
                self.text.tag_remove(f"size_{n}_bold", s, e)
            self.text.tag_add(size_tag, s, e)
        if align_tag:
            # 清除舊對齊，套用新對齊
            for a in ("left", "center", "right"):
                self.text.tag_remove(f"align_{a}", s, e)
            self.text.tag_add(align_tag, s, e)
        if col_tag:
            # 清除舊顏色，套用新顏色色碼
            for tag in list(self.text.tag_names(s)):
                if tag.startswith("color_"):
                    self.text.tag_remove(tag, s, e)
            self.text.tag_add(col_tag, s, e)

    def _set_size(self, size):
        self._size = size
        for s, b in self.size_btns.items():
            b.config(bg="#27ae60" if s == size else self.bg)
        self._apply_to_selection(size_tag=self._current_size_tag())

    def _toggle_bold(self):
        self._bold = not self._bold
        self.bold_btn.config(bg="#27ae60" if self._bold else self.bg)
        self._apply_to_selection(size_tag=self._current_size_tag())

    def _set_align(self, align):
        self._align = align
        for a, b in self.align_btns.items():
            b.config(bg="#27ae60" if a == align else self.bg)
        self._apply_to_selection(align_tag=f"align_{align}")

    def _pick_color(self):
        result = colorchooser.askcolor(color=self._color, title="選擇文字顏色", parent=self)
        if result and result[1]:
            self._color = result[1]
            self.color_swatch_btn.config(bg=self._color)
            col_tag = self._ensure_color_tag(self._color)
            self._apply_to_selection(col_tag=col_tag)

    def on_save(self):
        """
        儲存邏輯。
        將 Text 編輯器裡的內容逐行掃描，連同其附帶的格式標籤，打包組裝為 {"lines": [...]} 字典。
        將結果賦予 self.result 後銷毀視窗，讓呼叫方接手處理。
        """
        total = int(self.text.index(tk.END).split(".")[0])
        lines_data = []
        for i in range(1, total + 1):
            txt = self.text.get(f"{i}.0", f"{i}.end").strip()
            if not txt:
                continue
            fmt = self._get_line_format(i)
            lines_data.append({"text": txt, **fmt})
        self.result = {"lines": lines_data}
        self.destroy()


# ── Settings Dialog ────────────────────────────────────────────────────────────
class SettingsDialog(tk.Toplevel):
    """
    提供透明度調整、日曆底色配置，以及行事曆 JSON 資料的匯入/匯出介面。
    調整色彩與透明度時支援即時預覽 (Live Preview)。
    """
    def __init__(self, parent, settings, on_live_preview, on_apply, on_cancel,
                 import_fn, export_fn):
        super().__init__(parent)
        self.settings        = settings
        self.on_live_preview = on_live_preview
        self.on_apply_cb     = on_apply
        self.on_cancel_cb    = on_cancel
        self.import_fn       = import_fn
        self.export_fn       = export_fn
        self._orig_edit_opacity = settings.get("edit_opacity", 0.9)
        self._orig_bg           = settings.get("bg_color",     "#3c7ea1")
        self._preview_color     = self._orig_bg

        self.title("設定")
        self.geometry("380x450")
        self.resizable(False, False)
        self.configure(bg="#2e5f7a")
        self.grab_set()
        self.lift()

        self._section("透明度", "30%~100%")
        self.edit_var = tk.DoubleVar(value=self._orig_edit_opacity)
        self._slider_row(self.edit_var, 0.3, 1.0, self._on_change)

        self._section("日曆底色", "")
        color_row = tk.Frame(self, bg="#2e5f7a")
        color_row.pack(fill=tk.X, padx=20, pady=4)
        self.color_swatch = tk.Label(color_row, bg=self._preview_color,
                                     width=5, relief="solid", bd=1)
        self.color_swatch.pack(side=tk.LEFT, padx=(0, 8))
        self.color_hex_lbl = tk.Label(color_row, text=self._preview_color,
                                      bg="#2e5f7a", fg="white", font=("Courier", 10))
        self.color_hex_lbl.pack(side=tk.LEFT, expand=True, anchor="w")
        tk.Button(color_row, text="選擇顏色…", command=self._pick_color,
                  bg="#4a8cb0", fg="white", bd=0, font=("Arial", 9),
                  padx=8, pady=3).pack(side=tk.RIGHT)

        tk.Frame(self, bg="#4a8cb0", height=1).pack(fill=tk.X, padx=16, pady=10)

        self._section("資料管理", "")
        ie_row = tk.Frame(self, bg="#2e5f7a")
        ie_row.pack(fill=tk.X, padx=20, pady=4)
        tk.Button(ie_row, text="📥 匯入行事曆", command=self._do_import,
                  bg="#4a8cb0", fg="white", font=("Microsoft JhengHei", 9),
                  bd=0, padx=10, pady=5).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(ie_row, text="📤 匯出行事曆", command=self._do_export,
                  bg="#4a8cb0", fg="white", font=("Microsoft JhengHei", 9),
                  bd=0, padx=10, pady=5).pack(side=tk.LEFT)

        tk.Frame(self, bg="#4a8cb0", height=1).pack(fill=tk.X, padx=16, pady=(12, 0))
        btn_bar = tk.Frame(self, bg="#2e5f7a")
        btn_bar.pack(fill=tk.X, padx=16, pady=10)
        tk.Button(btn_bar, text="✕ 取消", command=self._cancel,
                  bg="#7a3c3c", fg="white", font=("Microsoft JhengHei", 10),
                  bd=0, padx=10, pady=5).pack(side=tk.LEFT)
        tk.Button(btn_bar, text="✔ 套用並關閉", command=self._apply,
                  bg="#27ae60", fg="white", font=("Microsoft JhengHei", 10, "bold"),
                  bd=0, padx=14, pady=5).pack(side=tk.RIGHT)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.transient(parent)
        self.wait_window()

    def _section(self, title, subtitle):
        tk.Label(self, text=title, font=("Microsoft JhengHei", 10, "bold"),
                 bg="#2e5f7a", fg="white").pack(anchor="w", padx=20, pady=(14, 0))
        if subtitle:
            tk.Label(self, text=subtitle, font=("Arial", 8),
                     bg="#2e5f7a", fg="#a0c4d9").pack(anchor="w", padx=20, pady=(0, 2))

    def _slider_row(self, var, from_, to, cmd):
        row = tk.Frame(self, bg="#2e5f7a")
        row.pack(fill=tk.X, padx=20, pady=2)
        pct = tk.Label(row, text=f"{int(var.get()*100)}%",
                       width=5, bg="#2e5f7a", fg="white", font=("Arial", 10))
        pct.pack(side=tk.RIGHT)
        def on_change(v):
            pct.config(text=f"{int(float(v)*100)}%")
            cmd()
        tk.Scale(row, variable=var, from_=from_, to=to, resolution=0.05,
                 orient=tk.HORIZONTAL, bg="#2e5f7a", fg="white",
                 troughcolor="#3c7ea1", highlightthickness=0,
                 showvalue=False, sliderlength=18, command=on_change
                 ).pack(fill=tk.X, expand=True)

    def _on_change(self):
        self.on_live_preview(self.edit_var.get(), self._preview_color)

    def _pick_color(self):
        result = colorchooser.askcolor(color=self._preview_color,
                                       title="選擇日曆底色", parent=self)
        if result and result[1]:
            self._preview_color = result[1]
            self.color_swatch.config(bg=self._preview_color)
            self.color_hex_lbl.config(text=self._preview_color)
            self._on_change()

    def _do_import(self):
        self.grab_release()
        self.import_fn()
        self.grab_set()

    def _do_export(self):
        self.grab_release()
        self.export_fn()
        self.grab_set()

    def _apply(self):
        self.settings["edit_opacity"] = round(self.edit_var.get(), 2)
        self.settings["bg_color"]     = self._preview_color
        self.on_apply_cb()
        self.destroy()

    def _cancel(self):
        self.settings["edit_opacity"] = self._orig_edit_opacity
        self.settings["bg_color"]     = self._orig_bg
        self.on_cancel_cb()
        self.destroy()


# ── Main App ──────────────────────────────────────────────────────────────────
class DesktopCalendar:
    """
    主應用程式類別。
    負責生成主視窗、建立日曆格狀 UI、綁定游標操作，
    以及驅動 Windows API 層確保日曆視窗「釘選」在桌面上且在所有視窗最底層。
    """
    def __init__(self, root):
        """
        初始化主程式系統狀態並建立 UI 元件。
        """
        self.root = root
        self.root.title("DesktopCalendar")
        # 將視窗屬性設定為 toolwindow，這會隱藏掉 Taskbar 上的圖示
        self.root.wm_attributes("-toolwindow", True)  

        # 資料路徑改為 .config 以降低防毒軟體誤報
        self.data_path = "tasks.config"
        self.old_data_path = "tasks.json"
        # 依照 JSON 檔案載入上次保存的位置或記事資料
        self.load_data()
        # 基於載入的設定，計算與初始化目前的 UI 色票組合
        self._rebuild_theme()
        # 立即設定根視窗底色，避免 Tkinter 預設色背景在載入時閃爍白框
        self.root.config(bg=self.theme["bg"])

        self.week_offset    = 0          # 0 = 當週; 正數/負數 = 偏移幾週的顯示
        self._drag          = {"x": 0, "y": 0}
        self._resize_start  = None

        self.setup_window()
        self.create_widgets()
        
        # 這是 DesktopCalendar 最關鍵的環節：將視窗釘在 Windows 桌面底層 (Desktop Layer)
        self.apply_desktop_layer()
        self.setup_tray()
        
        self._resize_job  = None
        self._last_size   = (0, 0)
        self.root.bind("<Configure>", self._on_configure)
        
        # 強制在主迴圈啟動前，完成所有的版面幾何運算與渲染
        # 這是為了解決剛啟動時按鈕未對齊或出現白邊的問題
        self.root.update_idletasks()

    def _on_configure(self, event):
        """
        Debounce 機制：監聽視窗拖曳縮放事件。
        在停止縮放/移動後 150 毫秒才重新繪製日曆格線與文字折行寬度，避免效能低落。
        """
        if event.widget is not self.root:
            return
        new_size = (event.width, event.height)
        if new_size == self._last_size:
            return
        self._last_size = new_size
        if self._resize_job:
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(150, self.draw_calendar)

        # ── Show Desktop survivability ─────────────────────────────────────────
        # 當使用者觸發「顯示桌面 (Win+D)」時，Windows 會強制 Unmap 所有視窗 (包含本程式)。
        # 因此我們必須掛上 <Unmap> 事件監聽器來攔截這個行為。
        self.root.bind("<Unmap>", self._on_unmap)

    def _on_unmap(self, event):
        """攔截被系統強制隱藏的事件，並觸發恢復可視的程序。"""
        if event.widget is self.root:
            # 延遲 80 毫秒以讓系統處理完畢，接著強制將自己喚醒
            self.root.after(80, self._restore_visibility)

    def _restore_visibility(self):
        """將視窗解除隱藏狀態，並立刻重新把 Z 軸排序壓回底層。"""
        self.root.deiconify()
        # Re-apply desktop z-order
        self.apply_desktop_layer()

    # ── Theme ─────────────────────────────────────────────────────────────────

    def _rebuild_theme(self):
        """
        根據使用者設定的基準底色，產生所有 UI 元件所需的衍生色系字典。
        包含 Header、Grid 與高亮顏色。
        """
        base = self.settings.get("bg_color", "#3c7ea1")
        self.theme = {
            "bg":        base,
            "header_bg": self._shift(base, +20),     # 標題列較亮一點
            "grid_bg":   self._shift(base, +8),      # 網格微亮
            "text":      "#ffffff",
            "highlight": "#ffd700",                  # 今日或特別高亮的顏色
            "subtle":    "#a0c4d9",
        }

    @staticmethod
    def _shift(hex_color, d):
        """輔助函式：對 Hex 顏色碼微調亮度 (加或減 RGB 分量)。"""
        try:
            h = hex_color.lstrip("#")
            r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
            return f"#{min(255,r+d):02x}{min(255,g+d):02x}{min(255,b+d):02x}"
        except Exception:
            return hex_color

    # ── Data ──────────────────────────────────────────────────────────────────

    def load_data(self):
        """
        嘗試從磁碟讀取 user 配置與記事紀錄 (JSON 格式)。
        若檔案不存在則套用預設值 (預設透明度 0.6、視窗大小 450x550)。
        同時處理從舊版 tasks.json 遷移至新版 tasks.config 的邏輯。
        """
        try:
            # 自動遷移舊版資料
            if not os.path.exists(self.data_path) and os.path.exists(self.old_data_path):
                # 遷移仍保留此 INFO，因為這屬於重要的一次性系統行為轉變
                logger.warning(f"正在遷移舊版資料檔 {self.old_data_path} -> {self.data_path}")
                os.rename(self.old_data_path, self.data_path)

            if os.path.exists(self.data_path):
                with open(self.data_path, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
            else:
                self.data = {"settings": {"bg_color": "#3c7ea1",
                                          "opacity": 0.6, "edit_opacity": 0.9,
                                          "pos_x": 100, "pos_y": 100,
                                          "width": 450, "height": 550}, "tasks": {}}
            self.settings = self.data.get("settings", {})
            self.tasks    = self.data.get("tasks",    {})
        except Exception as e:
            logger.error(f"載入資料失敗: {e}", exc_info=True)
            messagebox.showerror("錯誤", f"無法載入資料檔案: {e}\n詳情請見 app.log")
            # 發生嚴重錯誤時給予空資料夾
            self.data = {"settings": {}, "tasks": {}}
            self.settings = {}
            self.tasks = {}

    def save_data(self):
        """
        寫入使用者的變更。
        包含目前的視窗位置 (pos_x, pos_y) 與縮放大小，以及行事曆內的記事。
        """
        try:
            self.settings["pos_x"]  = self.root.winfo_x()
            self.settings["pos_y"]  = self.root.winfo_y()
            self.settings["width"]  = self.root.winfo_width()
            self.settings["height"] = self.root.winfo_height()
            self.data["settings"] = self.settings
            self.data["tasks"]    = self.tasks
            with open(self.data_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"儲存資料失敗: {e}", exc_info=True)

    # ── Window ────────────────────────────────────────────────────────────────

    def setup_window(self):
        """
        設定視窗的初始無邊框狀態 (Overrideredirect) 並取得先前寫入的大小與座標。
        同時對系統注入 Layered Attributes 以支援全視窗半透明化。
        """
        self.root.overrideredirect(True)
        w = self.settings.get("width",  450)
        h = self.settings.get("height", 550)
        x = self.settings.get("pos_x",  self.root.winfo_screenwidth() - w - 20)
        y = self.settings.get("pos_y",  100)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        
        # 取得系統層級控制代碼 (Handle)，設定擴充風格
        hwnd = self.get_hwnd()
        ex   = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex | WS_EX_LAYERED)

    def get_hwnd(self):
        """強制更新並取得目前視窗的真實作業系統控制代碼 (HWND)。"""
        self.root.update()
        return user32.GetParent(self.root.winfo_id())

    def apply_desktop_layer(self):
        """
        Pin the window to the desktop layer: always visible, never covers other apps.
        透過 Win32 API SetWindowPos 將視窗推入 HWND_BOTTOM (系統底層)，
        確保它能永遠浮在桌布之上，但被所有其他一般視窗覆蓋。
        """
        hwnd = self.get_hwnd()
        # 確保並未被設定為 TOPMOST 或穿透
        self.root.attributes("-topmost", False)
        
        # 使用 WinAPI 兩段式推入底層
        HWND_NOTOPMOST = -2
        user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0,
                            SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE)
        # Place below all normal windows (but above the desktop shell/wallpaper)
        user32.SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0,
                            SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE)
                            
        self.root.attributes("-alpha", self.settings.get("edit_opacity", 0.9))
        self.resize_grip.place(relx=1.0, rely=1.0, anchor="se")
        self.draw_calendar()

    # ── Widgets ───────────────────────────────────────────────────────────────

    def create_widgets(self):
        self.container = tk.Frame(self.root, bg=self.theme["bg"], highlightthickness=0)
        self.container.pack(fill=tk.BOTH, expand=True)

        # Header
        self.header = tk.Frame(self.container, bg=self.theme["header_bg"], height=46)
        self.header.pack(fill=tk.X)
        self.header.pack_propagate(False)
        for w in (self.header, self.container):
            w.bind("<Button-1>",  self.on_drag_start)
            w.bind("<B1-Motion>", self.on_drag_motion)

        self.month_lbl = tk.Label(self.header, text="",
                                  font=("Microsoft JhengHei", 14, "bold"),
                                  bg=self.theme["header_bg"], fg="white")
        self.month_lbl.pack(side=tk.LEFT, padx=14)
        self.month_lbl.bind("<Button-1>",  self.on_drag_start)
        self.month_lbl.bind("<B1-Motion>", self.on_drag_motion)

        # Nav & controls (right side)
        self.nav_frame = tk.Frame(self.header, bg=self.theme["header_bg"])
        self.nav_frame.pack(side=tk.RIGHT, padx=6)



        self.btn_gear = tk.Button(self.nav_frame, text="⚙",
                                  command=self.open_settings,
                                  bg=self.theme["header_bg"], fg="#c8e6f5", bd=0,
                                  font=("Arial", 13),
                                  activebackground=self.theme["header_bg"])
        self.btn_gear.pack(side=tk.LEFT, padx=2)

        self.btn_prev = tk.Button(self.nav_frame, text="‹", command=self.prev_month,
                                  bg=self.theme["header_bg"], fg="white", bd=0,
                                  font=("Arial", 15),
                                  activebackground=self.theme["header_bg"])
        self.btn_prev.pack(side=tk.LEFT)

        self.btn_next = tk.Button(self.nav_frame, text="›", command=self.next_month,
                                  bg=self.theme["header_bg"], fg="white", bd=0,
                                  font=("Arial", 15),
                                  activebackground=self.theme["header_bg"])
        self.btn_next.pack(side=tk.LEFT, padx=4)

        # Calendar grid
        self.grid_frame = tk.Frame(self.container, bg=self.theme["bg"])
        self.grid_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)



        # Resize grip
        self.resize_grip = tk.Label(self.container, text="⤡", font=("Arial", 14),
                                    bg=self.theme["bg"], fg="#a0c4d9", cursor="size_nw_se")
        self.resize_grip.bind("<Button-1>",        self.on_resize_start)
        self.resize_grip.bind("<B1-Motion>",       self.on_resize_motion)
        self.resize_grip.bind("<ButtonRelease-1>", self.on_resize_end)

    # ── Full restyle ──────────────────────────────────────────────────────────

    def _restyle_all(self):
        bg  = self.theme["bg"]
        hbg = self.theme["header_bg"]
        self.container.config(bg=bg)
        self.header.config(bg=hbg)
        self.month_lbl.config(bg=hbg)
        self.nav_frame.config(bg=hbg)
        for btn in (self.btn_gear, self.btn_prev, self.btn_next):
            btn.config(bg=hbg, activebackground=hbg)
        self.grid_frame.config(bg=bg)
        self.resize_grip.config(bg=bg)

    # ── Calendar ──────────────────────────────────────────────────────────────

    def _cell_wrap_width(self):
        """
        計算日曆單格 (Cell) 內預覽文字的自動折行寬度 (wraplength)。
        依據目前整個日曆 Grid 的總寬度等分推算，確保縮放視窗時文字能正確斷行。
        """
        gw = self.grid_frame.winfo_width()
        if gw < 50:          # 畫面尚未完全渲染完畢時提供安全預設值
            gw = self.root.winfo_width()
        return max(30, (gw // 7) - 10)

    def _get_weeks(self):
        """
        產生當前應顯示的 6 週日期矩陣 (6 x 7)。
        固定以包含「今日」的那一週的星期一為錨點，並透過 self.week_offset 支援前後翻閱。
        回傳值格式: [[date, date, ...], [date, ...], ...]
        """
        today = date.today()
        # 找出當週的星期一 (weekday() 0 為週一)
        monday = today - timedelta(days=today.weekday())
        # 根據翻頁偏移量決定畫面的起始日期
        start = monday + timedelta(weeks=self.week_offset)
        
        weeks = []
        for w in range(6):
            week = [start + timedelta(days=w*7 + d) for d in range(7)]
            weeks.append(week)
        return weeks

    def draw_calendar(self):
        """
        核心渲染邏輯：清空舊有的日曆網格，並根據 _get_weeks() 繪製最新月份與內容。
        處理包含跨月反灰、今日高亮、以及截斷預覽太長的記事項目。
        """
        for w in self.grid_frame.winfo_children():
            w.destroy()
            
        today = date.today()
        weeks = self._get_weeks()

        # ── 解析並渲染 Header 標題列 (例如 "2024年 3月 / 4月") ──
        months_seen = []
        for week in weeks:
            for d in week:
                key = (d.year, d.month)
                if key not in months_seen:
                    months_seen.append(key)
                    
        # 若當前 6 週都落在同一個月份內
        if len(months_seen) == 1:
            y, m = months_seen[0]
            header_text = f"{y}年 {m}月"
        # 若畫面出現跨月 (甚至跨年) 的情況
        else:
            parts = []
            prev_year = None
            for y, m in months_seen:
                if y != prev_year:
                    parts.append(f"{y}年 {m}月")
                    prev_year = y
                else:
                    parts.append(f"{m}月")
            header_text = " / ".join(parts)
        self.month_lbl.config(text=header_text)

        # ── 繪製星期標題列 (一~日) ──
        for i, d in enumerate(["一","二","三","四","五","六","日"]):
            tk.Label(self.grid_frame, text=d, font=("Microsoft JhengHei", 9),
                     bg=self.theme["bg"], fg="#c8e6f5").grid(
                row=0, column=i, sticky="nsew", pady=4)

        # ── 迴圈繪製 42 宮格的日期與內容 ──
        for r, week in enumerate(weeks):
            for c, day_date in enumerate(week):
                date_str  = day_date.strftime("%Y-%m-%d")
                raw_note  = self.tasks.get(date_str, None)
                note_text = self._note_text(raw_note) if raw_note else ""
                has_note  = bool(note_text)
                is_today  = (day_date == today)
                
                # 判定跨月淡化顯示：以畫面上第一週的週一作為主要月份判定基準
                dominant_month = week[0].month
                is_other_month = (day_date.month != dominant_month)
                
                # 若該天有記事，替網格加上亮底色
                cell_bg   = self._shift(self.theme["grid_bg"], 18) if has_note else self.theme["grid_bg"]
                
                # 字體顏色邏輯：今日 (金黃) > 非本月 (暗灰) > 正常 (白)
                if is_today:
                    day_fg = self.theme["highlight"]
                elif is_other_month:
                    day_fg = "#7ab0cc"   
                else:
                    day_fg = "white"
                    
                # 只有在月初 (1號) 時，額外顯示月份標籤 (除了整個畫面的第一格之外)
                if day_date.day == 1 and not (r == 0 and c == 0):
                    day_label = f"{day_date.month}/{day_date.day}"
                else:
                    day_label = str(day_date.day)
                    
                cell = tk.Frame(self.grid_frame, bg=cell_bg,
                                highlightbackground="#a0d8ef",
                                highlightthickness=1 if (is_today or has_note) else 0)
                cell.grid(row=r+1, column=c, sticky="nsew", padx=1, pady=1)
                
                # 放上左上角的日期數字
                tk.Label(cell, text=day_label, font=("Arial", 10, "bold"),
                         bg=cell_bg, fg=day_fg).pack(anchor="nw", padx=2, pady=1)
                         
                # ── 繪製最多三行的濃縮預覽記事 ──
                if has_note:
                    note_lines = parse_note_lines(raw_note)
                    sz_map = {"小": 7, "中": 8, "大": 10}
                    wrap   = self._cell_wrap_width()
                    
                    for idx, ln in enumerate(note_lines[:3]):
                        sz      = sz_map.get(ln.get("size", "中"), 8)
                        weight  = "bold" if ln.get("bold") else "normal"
                        justify = ln.get("align", "left")
                        anchor  = {"left": "w", "center": "center", "right": "e"}.get(justify, "w")
                        fg_col  = ln.get("color", "#dff0ff")
                        
                        tk.Label(cell, text=f"• {ln['text']}",
                                 font=("Microsoft JhengHei", sz, weight),
                                 bg=cell_bg, fg=fg_col,
                                 justify=justify, anchor=anchor,
                                 wraplength=wrap).pack(fill=tk.X, padx=2, anchor=anchor)
                                 
                    # 若超過三行，顯示省略提示
                    if len(note_lines) > 3:
                        tk.Label(cell, text=f"  …(+{len(note_lines)-3})",
                                 font=("Microsoft JhengHei", 7),
                                 bg=cell_bg, fg="#dff0ff").pack(fill=tk.X, padx=2, anchor="w")
                                 
                # 綁定雙擊事件：不論是點擊外框或是內部元件，皆觸發編輯視窗
                cell.bind("<Double-Button-1>", lambda e, d=date_str: self.edit_note(d))
                for ch in cell.winfo_children():
                    ch.bind("<Double-Button-1>", lambda e, d=date_str: self.edit_note(d))

        for i in range(7):
            self.grid_frame.grid_columnconfigure(i, weight=1)
        for i in range(1, 8):
            self.grid_frame.grid_rowconfigure(i, weight=1)

    # ── Drag / Resize ─────────────────────────────────────────────────────────

    def on_drag_start(self, e):
        """記錄滑鼠游標拖曳起始位置。"""
        self._drag["x"] = e.x; self._drag["y"] = e.y

    def on_drag_motion(self, e):
        """根據滑鼠移動的相對距離更新視窗的絕對座標。"""
        x = self.root.winfo_x() + (e.x - self._drag["x"])
        y = self.root.winfo_y() + (e.y - self._drag["y"])
        self.root.geometry(f"+{x}+{y}")

    def on_resize_start(self, e):
        """記錄縮放起始點，包含滑鼠與視窗寬高。"""
        self._resize_start = (e.x_root, e.y_root,
                               self.root.winfo_width(), self.root.winfo_height())

    def on_resize_motion(self, e):
        """右下角的縮放判定邏輯，限制最小尺寸不可低於 300x250。"""
        if self._resize_start:
            ox, oy, ow, oh = self._resize_start
            self.root.geometry(f"{max(300,ow+(e.x_root-ox))}x{max(250,oh+(e.y_root-oy))}")

    def on_resize_end(self, e):
        """縮放結束後保存新的視窗大小至 JSON 配置中。"""
        self._resize_start = None
        self.save_data()

    # ── Note helpers ──────────────────────────────────────────────────────────

    @staticmethod
    def _note_text(note):
        """工具函式：將包含多種格式的複雜記事結構，硬行轉換為純文字字串以供檢查。"""
        lines = parse_note_lines(note)
        return "\n".join(ln["text"] for ln in lines)

    def edit_note(self, date_str):
        raw = self.tasks.get(date_str, {})
        dlg = NoteDialog(self.root, date_str, raw, self.theme)
        if dlg.result is not None:
            lines = dlg.result.get("lines", [])
            if lines:
                self.tasks[date_str] = dlg.result
            else:
                self.tasks.pop(date_str, None)
            self.save_data()
            self.draw_calendar()

    # ── Navigation ────────────────────────────────────────────────────────────

    def prev_month(self):
        """Scroll back one week."""
        self.week_offset -= 1
        self.draw_calendar()

    def next_month(self):
        """Scroll forward one week."""
        self.week_offset += 1
        self.draw_calendar()

    # ── Settings ──────────────────────────────────────────────────────────────

    def open_settings(self):
        def live(edit_opacity, bg_color):
            self.settings["edit_opacity"] = round(edit_opacity, 2)
            self.settings["bg_color"]     = bg_color
            self._rebuild_theme()
            self._restyle_all()
            self.draw_calendar()
            self.root.attributes("-alpha", self.settings["edit_opacity"])

        def apply():
            self.save_data()

        def cancel():
            self._rebuild_theme()
            self._restyle_all()
            self.draw_calendar()
            self.root.attributes("-alpha", self.settings["edit_opacity"])

        SettingsDialog(self.root, self.settings, live, apply, cancel,
                       self.import_data, self.export_data)

    # ── Import / Export ───────────────────────────────────────────────────────

    def export_data(self):
        """
        將日曆的 JSON 結構匯出存為額外備份擋案。
        """
        path = filedialog.asksaveasfilename(defaultextension=".json",
                                            filetypes=[("JSON","*.json")],
                                            title="匯出行事曆")
        if path:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2, ensure_ascii=False)
            messagebox.showinfo("完成", "匯出成功。")

    def import_data(self):
        """
        從外部 JSON 檔案讀取記事，並詢問合併或取代目前行事曆。
        """
        path = filedialog.askopenfilename(filetypes=[("JSON","*.json")], title="匯入行事曆")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                new = json.load(f)
                
            # 提供三選一模式
            choice = messagebox.askyesnocancel("匯入", "「是」合併　「否」取代　「取消」放棄")
            if choice is True:
                self.tasks.update(new.get("tasks", {}))
            elif choice is False:
                self.tasks = new.get("tasks", {})
            else:
                return
                
            self.save_data()
            self.draw_calendar()
            messagebox.showinfo("完成", "匯入成功。")
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

    # ── Tray (工作列圖示) ─────────────────────────────────────────────────────

    def setup_tray(self):
        """
        建置 pystray 系統匣常駐圖示，為 DesktopCalendar 提供關閉程式的入口。
        本體視窗為了能無干擾釘存在桌面，已被移出一般工作列名單，
        因此系統匣變成唯一的應用程式進入點之一。
        """
        img  = Image.new("RGBA", (64,64), (0,0,0,0))
        draw = ImageDraw.Draw(img)
        # 用簡單的幾何繪製一個模擬日曆圖標 (相容舊版 Pillow)
        _fill = "#3c7ea1"
        _r = 10
        _x0, _y0, _x1, _y1 = 4, 4, 60, 60
        draw.ellipse([_x0, _y0, _x0+_r*2, _y0+_r*2], fill=_fill)
        draw.ellipse([_x1-_r*2, _y0, _x1, _y0+_r*2], fill=_fill)
        draw.ellipse([_x0, _y1-_r*2, _x0+_r*2, _y1], fill=_fill)
        draw.ellipse([_x1-_r*2, _y1-_r*2, _x1, _y1], fill=_fill)
        draw.rectangle([_x0+_r, _y0, _x1-_r, _y1], fill=_fill)
        draw.rectangle([_x0, _y0+_r, _x1, _y1-_r], fill=_fill)
        draw.rectangle([14,20,50,44], outline="white", width=2)
        draw.line([14,27,50,27], fill="white", width=1)
        draw.line([26,14,26,22], fill="white", width=2)
        draw.line([38,14,38,22], fill="white", width=2)
        menu = pystray.Menu(
            pystray.MenuItem("設定",      lambda i,it: self.root.after(0, self.open_settings)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("結束程式",  lambda i,it: self.root.after(0, self.exit_app))
        )
        self.tray = pystray.Icon("DesktopCalendar", img, "DesktopCalendar", menu)
        threading.Thread(target=self.tray.run, daemon=True).start()

    def exit_app(self, *_):
        self.save_data()
        try:
            self.tray.stop()
        except Exception:
            pass
        self.root.destroy()


# ── Entry ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app  = DesktopCalendar(root)
    root.protocol("WM_DELETE_WINDOW", app.exit_app)
    root.mainloop()
