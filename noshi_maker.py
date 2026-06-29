#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
のし紙メーカー v3.1
横向き / 縦書き / ドラッグ配置 / フォント選択 / 水引プレビュー
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont
import os, sys, tempfile, json, re

try:
    import winreg
    HAS_WINREG = True
except ImportError:
    HAS_WINREG = False

try:
    import win32print
    HAS_WIN32PRINT = True
except ImportError:
    HAS_WIN32PRINT = False

try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pillow"])
    from PIL import Image, ImageDraw, ImageFont, ImageTk


# ======================================================
# 定数
# ======================================================
PAPER_SIZES = {
    "美濃版  (394×267mm)": (394, 267),
    "半紙版  (337×245mm)": (337, 245),
    "切手版  (260×185mm)": (260, 185),
    "豆のし  (227×127mm)": (227, 127),
}
PRINT_DPI = 300

# 各要素のデフォルト配置 (x比率, y比率, 文字サイズ/紙高さ比率)
# 文字サイズは美濃版(高さ267mm)基準: 22mm/267≈0.082, 19mm/267≈0.071, 11mm/267≈0.041
DEFAULTS = {
    "表書き": (0.50, 0.22, 0.082),   # ①中央・少し上・22mm
    "社名":   (0.50, 0.70, 0.071),   # ②中央・少し下・19mm
    "役職":   (0.40, 0.67, 0.041),   # ③社名の左・社名の1/3から・11mm
    "氏名":   (0.50, 0.75, 0.071),   # ④デフォルトは中央・社名/役職があれば左に自動移動
}

MIZUHIKI_Y_FRAC = 0.45   # 水引の縦位置（プレビュー用）
PRINT_BIN       = 4       # DMBIN_MANUAL = 手差し（トレイ5）
SAVE_DIR        = os.path.join(os.path.dirname(os.path.abspath(__file__)), "のし紙")


# ======================================================
# フォント管理
# ======================================================
def scan_system_fonts():
    """
    Windowsレジストリから全インストールフォントを取得。
    戻り値: {表示名: ファイルパス}
    """
    fonts = {}
    font_dir = os.path.join(
        os.environ.get("SystemRoot", r"C:\Windows"), "Fonts")

    if HAS_WINREG:
        try:
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            )
            i = 0
            while True:
                try:
                    name, path, _ = winreg.EnumValue(key, i)
                    # "(TrueType)" "(OpenType)" 等を除去してクリーンな表示名に
                    display = name
                    for suffix in [" (TrueType)", " (OpenType)",
                                   " (TrueType,OpenType)", " (OpenType,TrueType)"]:
                        display = display.replace(suffix, "")
                    display = display.strip()
                    if not os.path.isabs(path):
                        path = os.path.join(font_dir, path)
                    if os.path.exists(path) and display:
                        fonts[display] = path
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)
        except Exception as e:
            print(f"レジストリ読み込みエラー: {e}")

    # フォールバック: 固定リスト
    if not fonts:
        for name, fname in {
            "游明朝":    "yumin.ttf",
            "MS明朝":    "msmincho.ttc",
            "游ゴシック": "YuGothM.ttc",
            "メイリオ":   "meiryo.ttc",
            "MSゴシック": "msgothic.ttc",
        }.items():
            p = os.path.join(font_dir, fname)
            if os.path.exists(p):
                fonts[name] = p

    return fonts


# 起動時にスキャン
FONT_REGISTRY  = scan_system_fonts()
ALL_FONT_NAMES = sorted(FONT_REGISTRY.keys())

# デフォルトフォント（優先順に探す）
DEFAULT_FONT = next(
    (n for n in ["HG正楷書体-PRO", "游明朝", "MS明朝", "游ゴシック", "メイリオ", "MSゴシック"]
     if n in FONT_REGISTRY),
    ALL_FONT_NAMES[0] if ALL_FONT_NAMES else "Yu Gothic UI"
)


def mm_to_px(mm, dpi):
    return int(mm * dpi / 25.4)


def get_pil_font(font_key, size_px):
    """PIL用フォントを取得（ファイルパス経由）"""
    path = FONT_REGISTRY.get(font_key)
    if path:
        try:
            return ImageFont.truetype(path, size_px)
        except Exception:
            pass
    try:
        return ImageFont.load_default(size=size_px)
    except Exception:
        return ImageFont.load_default()


def get_tk_font(font_key, size_px):
    """tkinter用フォントを取得（フォントファミリー名経由）"""
    return tkfont.Font(family=font_key, size=-abs(size_px))


def draw_vertical_pil(draw, text, cx, cy, font, spacing=1.25):
    """PIL: テキストを (cx, cy) 中心に縦書きで描画（spacing=字間倍率）
    半角スペース → 0.5文字分の縦送り
    全角スペース → 1文字分の縦送り（通常）
    """
    if not text:
        return
    chars = list(text)

    # 基準文字サイズ（「一」で計測）
    ref_box = draw.textbbox((0, 0), "一", font=font)
    ref_h   = max(1, ref_box[3] - ref_box[1])
    ref_w   = max(1, ref_box[2] - ref_box[0])

    def _ch(c):
        """縦方向の文字高さ（スペースは特別扱い）"""
        if c == " ":       return ref_h * 0.5   # 半角スペース = 半文字
        if c == "　":  return ref_h          # 全角スペース = 1文字
        b = draw.textbbox((0, 0), c, font=font)
        h = b[3] - b[1]
        return h if h > 0 else ref_h

    def _cw(c):
        if c in (" ", "　"): return ref_w
        b = draw.textbbox((0, 0), c, font=font)
        return max(1, b[2] - b[0])

    chs     = [_ch(c) for c in chars]
    cws     = [_cw(c) for c in chars]
    total_h = sum(ch * spacing for ch in chs)
    y = cy - total_h / 2
    for i, char in enumerate(chars):
        if char not in (" ", "　"):          # スペース自体は描画しない
            draw.text((cx - cws[i] / 2, y), char, font=font, fill="black")
        y += chs[i] * spacing


# ======================================================
# テキスト要素
# ======================================================
class TextElement:
    def __init__(self, text_var, label):
        self.text_var = text_var
        self.label    = label
        self.bbox     = None
        self.reset()

    @property
    def text(self):
        return self.text_var.get()

    def reset(self):
        x, y, s = DEFAULTS[self.label]
        self.x_frac    = x
        self.y_frac    = y
        self.size_frac = s


# ======================================================
# インタラクティブキャンバス
# ======================================================
class InteractiveCanvas:
    HANDLE = 8

    def __init__(self, canvas):
        self.canvas        = canvas
        self._elements     = []
        self._font_key     = DEFAULT_FONT
        self._char_spacing = 1.25  # デフォルト字間
        self._selected     = None
        self._px = self._py = self._pw = self._ph = 0

        self._drag_sm   = None
        self._drag_sd   = None
        self._drag_mode = None

        self.on_select = lambda e: None

        canvas.bind("<ButtonPress-1>",   self._on_press)
        canvas.bind("<B1-Motion>",       self._on_motion)
        canvas.bind("<ButtonRelease-1>", self._on_release)

    def set_elements(self, elems):
        self._elements = elems

    def set_font(self, key):
        self._font_key = key

    def set_spacing(self, v):
        self._char_spacing = max(0.8, float(v))

    # ── 全体描画 ─────────────────────────────────────
    def render_all(self, w_mm, h_mm):
        c = self.canvas
        c.delete("all")
        cw = c.winfo_width()  or 640
        ch = c.winfo_height() or 400
        m  = 16

        if w_mm / h_mm > (cw - 2*m) / (ch - 2*m):
            pw = cw - 2*m
            ph = int(pw * h_mm / w_mm)
        else:
            ph = ch - 2*m
            pw = int(ph * w_mm / h_mm)

        px = (cw - pw) // 2
        py = (ch - ph) // 2
        self._px, self._py, self._pw, self._ph = px, py, pw, ph

        # 影
        c.create_rectangle(px+4, py+4, px+pw+4, py+ph+4,
                            fill="#999999", outline="")
        # 紙の背景
        c.create_rectangle(px, py, px+pw, py+ph,
                            fill="white", outline="#777777", width=1)

        # ★ 中央ガイドライン（プレビューのみ）
        cx = px + pw // 2
        c.create_line(cx, py + 4, cx, py + ph - 4,
                      fill="#BBDDFF", width=1, dash=(6, 4), tags="guide")
        c.create_text(cx, py + 10,
                      text="┃ 中央", font=("Yu Gothic UI", 7),
                      fill="#99BBDD", anchor="n", tags="guide")

        # ★ 水引（プレビューのみ・印刷対象外）
        self._draw_mizuhiki(px, py, pw, ph)

        # テキスト要素（水引の上に重なるようにこの後に描画）
        for elem in self._elements:
            self._draw_elem(elem, selected=(elem is self._selected))

    # ── 水引描画（プレビュー専用） ────────────────────
    def _draw_mizuhiki(self, px, py, pw, ph):
        tag = "mizuhiki"
        n       = 5
        colors  = ["#CC0000", "#CC0000", "#CC0000", "#FFD700", "#CC0000"]
        spacing = max(2, int(ph * 0.009))
        lw      = max(1, int(pw * 0.003))
        cy      = py + int(ph * MIZUHIKI_Y_FRAC)
        cx      = px + pw // 2
        knot_w  = int(pw * 0.10)

        ys = [cy + (i - (n - 1) / 2) * spacing for i in range(n)]

        # 左右の線（結び目部分を空ける）
        for i, y in enumerate(ys):
            self.canvas.create_line(px, y, cx - knot_w, y,
                                    fill=colors[i], width=lw, tags=tag)
            self.canvas.create_line(cx + knot_w, y, px + pw, y,
                                    fill=colors[i], width=lw, tags=tag)

        # 蝶結び
        ks  = int(ph * 0.042)
        lw2 = max(2, lw + 1)

        # 左の輪
        self.canvas.create_oval(
            cx - ks*2, cy - int(ks*0.6),
            cx - int(ks*0.1), cy + int(ks*0.6),
            outline="#CC0000", width=lw2, tags=tag
        )
        # 右の輪
        self.canvas.create_oval(
            cx + int(ks*0.1), cy - int(ks*0.6),
            cx + ks*2,        cy + int(ks*0.6),
            outline="#CC0000", width=lw2, tags=tag
        )
        # 中央の結び目
        kr = max(3, int(ks * 0.27))
        self.canvas.create_oval(
            cx - kr, cy - kr, cx + kr, cy + kr,
            fill="#CC0000", outline="#CC0000", tags=tag
        )
        # 垂れ（下部）
        tail = int(ks * 0.65)
        for dx in [-int(ks*0.15), int(ks*0.15)]:
            self.canvas.create_line(
                cx + dx, cy + kr,
                cx + dx + int(ks*0.3), cy + kr + tail,
                fill="#CC0000", width=lw, tags=tag
            )

        # ガイドラベル
        self.canvas.create_text(
            px + pw - 3, cy,
            text="← 水引（印刷されません）",
            font=("Yu Gothic UI", 8), fill="#BBBBBB",
            anchor="e", tags=tag
        )

    # ── テキスト要素描画 ──────────────────────────────
    def _draw_elem(self, elem, selected=False):
        etag = f"E{id(elem)}"
        stag = f"S{id(elem)}"
        self.canvas.delete(etag)
        self.canvas.delete(stag)

        if not elem.text:
            elem.bbox = None
            return

        cx    = self._px + int(elem.x_frac * self._pw)
        cy    = self._py + int(elem.y_frac * self._ph)
        fsize = max(6, int(elem.size_frac * self._ph))

        try:
            f = get_tk_font(self._font_key, fsize)
        except Exception:
            f = tkfont.Font(size=-abs(fsize))

        chars   = list(elem.text)
        lh      = f.metrics("linespace")
        step    = lh * self._char_spacing       # 通常の1文字送り幅

        # 基準幅（最大幅の計算用）
        ref_lh = lh  # スペース以外の基準

        # 各文字の送り幅と表示幅を決定（半角/全角スペース区別）
        steps = []
        ws    = []
        ref_w = f.measure("一") or fsize
        for ch in chars:
            if ch == " ":              # 半角スペース → 0.5文字送り
                steps.append(step * 0.5)
                ws.append(ref_w)
            elif ch == "　":       # 全角スペース → 1文字送り
                steps.append(step)
                ws.append(ref_w)
            else:
                steps.append(step)
                ws.append(f.measure(ch) or fsize)

        maxw    = max(ws) if ws else fsize
        total_h = sum(steps)

        y = cy - total_h / 2
        for i, ch in enumerate(chars):
            if ch not in (" ", "　"):   # スペースは描画しない
                self.canvas.create_text(
                    cx, y + lh / 2,
                    text=ch, font=f, fill="black",
                    anchor="center", tags=etag
                )
            y += steps[i]

        pad = 4
        x0 = cx - maxw / 2 - pad
        y0 = cy - total_h / 2 - pad
        x1 = cx + maxw / 2 + pad
        y1 = cy + total_h / 2 + pad
        # step込みの高さで bbox を設定済み
        elem.bbox = (x0, y0, x1, y1)

        if selected:
            self.canvas.create_rectangle(
                x0, y0, x1, y1,
                outline="#0078D7", width=2, dash=(5, 3), tags=stag
            )
            h = self.HANDLE
            self.canvas.create_rectangle(
                x1 - h, y1 - h, x1 + h, y1 + h,
                fill="#0078D7", outline="white", width=1, tags=stag
            )

    # ── ヒットテスト ──────────────────────────────────
    def _hit_test(self, x, y):
        # リサイズハンドル（選択中のみ）
        if self._selected and self._selected.bbox:
            bx0, by0, bx1, by1 = self._selected.bbox
            h = self.HANDLE
            if bx1 - h <= x <= bx1 + h and by1 - h <= y <= by1 + h:
                return self._selected, "resize"
        for elem in reversed(self._elements):
            if elem.bbox:
                bx0, by0, bx1, by1 = elem.bbox
                if bx0 <= x <= bx1 and by0 <= y <= by1:
                    return elem, "move"
        return None, None

    # ── マウスイベント ────────────────────────────────
    def _on_press(self, event):
        old = self._selected
        elem, mode = self._hit_test(event.x, event.y)
        self._selected  = elem
        self._drag_mode = mode
        self._drag_sm   = (event.x, event.y)
        self._drag_sd   = (elem.x_frac, elem.y_frac, elem.size_frac) if elem else None

        if old and old is not elem:
            self._draw_elem(old, selected=False)
        if elem:
            self._draw_elem(elem, selected=True)

        cursor = {"move": "fleur", "resize": "sizing"}.get(mode, "arrow")
        self.canvas.configure(cursor=cursor)
        self.on_select(elem)

    def _on_motion(self, event):
        if not self._selected or not self._drag_sm:
            return
        if self._pw == 0 or self._ph == 0:
            return

        dx = event.x - self._drag_sm[0]
        dy = event.y - self._drag_sm[1]
        x0, y0, s0 = self._drag_sd

        if self._drag_mode == "move":
            self._selected.x_frac = max(0.03, min(0.97, x0 + dx / self._pw))
            self._selected.y_frac = max(0.03, min(0.97, y0 + dy / self._ph))
        elif self._drag_mode == "resize":
            self._selected.size_frac = max(0.015, min(0.45, s0 + dy / self._ph))

        self._draw_elem(self._selected, selected=True)
        self.on_select(self._selected)

    def _on_release(self, event):
        self._drag_sm = self._drag_sd = None
        self.canvas.configure(cursor="arrow")


# ======================================================
# メインアプリ
# ======================================================
class NoshiApp:

    def __init__(self, root):
        self.root = root
        root.title("のし紙メーカー")
        root.geometry("1100x650")
        root.minsize(820, 520)
        root.configure(bg="#EBEBEB")

        self.upper_text   = tk.StringVar(value="御祝")
        self.company      = tk.StringVar(value="")
        self.position     = tk.StringVar(value="")
        self.name         = tk.StringVar(value="山田 太郎")
        self.paper_size   = tk.StringVar(value=list(PAPER_SIZES.keys())[0])
        self.font_name    = tk.StringVar(value=DEFAULT_FONT)
        self.size_mm_var  = tk.StringVar(value="0.0")
        self.char_spacing = tk.DoubleVar(value=1.25)

        self.elements = [
            TextElement(self.upper_text, "表書き"),
            TextElement(self.company,    "社名"),
            TextElement(self.position,   "役職"),
            TextElement(self.name,       "氏名"),
        ]

        self._selected_elem = None
        self._after_id      = None
        self._lock_size_cb  = False

        self._build_ui()

        self.icanvas.set_elements(self.elements)
        self.icanvas.on_select = self._on_elem_selected

        for var in (self.upper_text, self.company, self.position,
                    self.name, self.paper_size, self.font_name, self.char_spacing):
            var.trace_add("write", lambda *_: self._schedule_render())

        self.company.trace_add("write", self._update_name_position)
        self.position.trace_add("write", self._update_name_position)

        self.size_mm_var.trace_add("write", self._on_size_spin_changed)
        root.bind("<Configure>", lambda e: self._schedule_render())

    # ── UI構築 ───────────────────────────────────────
    def _build_ui(self):
        s = ttk.Style()
        s.theme_use("clam")
        bg = "#EBEBEB"
        s.configure("TFrame",            background=bg)
        s.configure("TLabelframe",       background=bg)
        s.configure("TLabelframe.Label", background=bg,
                    font=("Yu Gothic UI", 10, "bold"))
        s.configure("TLabel",       background=bg, font=("Yu Gothic UI", 10))
        s.configure("TButton",      font=("Yu Gothic UI", 10))
        s.configure("TRadiobutton", background=bg)
        s.configure("Accent.TButton",
                    font=("Yu Gothic UI", 11, "bold"),
                    foreground="white", background="#0078D7")
        s.map("Accent.TButton",
              background=[("active", "#005FA3"), ("pressed", "#003D6B")])

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(main, text="  設 定  ", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)

        right = ttk.LabelFrame(
            main,
            text="  プレビュー  ─  クリックで選択 ／ ドラッグで移動 ／ ■でサイズ変更  ",
            padding=6
        )
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        self._build_controls(left)
        self._build_canvas(right)

        self.status_var = tk.StringVar(value="要素をクリックして選択してください")
        ttk.Label(main, textvariable=self.status_var,
                  font=("Yu Gothic UI", 9), foreground="#555555"
                  ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 0))

    def _build_controls(self, parent):
        row = 0

        # ── フォント ──
        ttk.Label(parent, text="フォント",
                  font=("Yu Gothic UI", 10, "bold")).grid(
            row=row, column=0, sticky="w", pady=(0, 3))
        row += 1

        font_frame = ttk.Frame(parent)
        font_frame.grid(row=row, column=0, sticky="ew", pady=(0, 2))
        font_frame.columnconfigure(0, weight=1)

        self._font_combo = ttk.Combobox(
            font_frame, textvariable=self.font_name,
            values=ALL_FONT_NAMES,
            font=("Yu Gothic UI", 10)
        )
        self._font_combo.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self._font_combo.bind("<KeyRelease>",  self._on_font_key)
        self._font_combo.bind("<FocusIn>",     self._on_font_focus_in)

        ttk.Button(font_frame, text="参照…", width=6,
                   command=self._browse_font).grid(row=0, column=1)
        row += 1

        ttk.Label(parent,
                  text=f"（{len(ALL_FONT_NAMES)}フォント検出  ・  名前を入力して絞り込み可）",
                  font=("Yu Gothic UI", 8), foreground="#888888").grid(
            row=row, column=0, sticky="w", pady=(0, 6))
        row += 1

        # ── 字間 ──
        sp_frame = ttk.Frame(parent)
        sp_frame.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        sp_frame.columnconfigure(1, weight=1)
        ttk.Label(sp_frame, text="字間:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._spacing_slider = ttk.Scale(
            sp_frame, from_=0.9, to=2.0,
            orient="horizontal", variable=self.char_spacing
        )
        self._spacing_slider.grid(row=0, column=1, sticky="ew")
        self._spacing_label = ttk.Label(sp_frame, text="1.25", width=4)
        self._spacing_label.grid(row=0, column=2, padx=(6, 0))
        self.char_spacing.trace_add("write", self._on_spacing_changed)
        row += 1

        # ── 表書き ──
        ttk.Label(parent, text="表書き",
                  font=("Yu Gothic UI", 10, "bold")).grid(
            row=row, column=0, sticky="w", pady=(0, 3))
        row += 1

        pf = ttk.Frame(parent)
        pf.grid(row=row, column=0, sticky="ew", pady=(0, 4))
        for r_i, pr in enumerate([
            ["御祝",   "御礼",  "御中元", "御歳暮"],
            ["御年賀", "粗品",  "寸志",   "内祝"],
            ["御見舞", "快気祝", "御霊前", "御仏前"],
        ]):
            for c_i, p in enumerate(pr):
                ttk.Button(pf, text=p, width=7,
                           command=lambda v=p: self.upper_text.set(v)
                           ).grid(row=r_i, column=c_i, padx=2, pady=2, sticky="ew")
        row += 1

        ttk.Entry(parent, textvariable=self.upper_text,
                  font=("Yu Gothic UI", 12)).grid(
            row=row, column=0, sticky="ew", pady=(0, 10))
        row += 1

        # ── お名前 ──
        ttk.Label(parent, text="お名前",
                  font=("Yu Gothic UI", 10, "bold")).grid(
            row=row, column=0, sticky="w", pady=(0, 3))
        row += 1

        for label, var in [("社名（任意）", self.company),
                            ("役職（任意）", self.position),
                            ("氏名",        self.name)]:
            ttk.Label(parent, text=label, foreground="#666666",
                      font=("Yu Gothic UI", 9)).grid(
                row=row, column=0, sticky="w")
            row += 1
            ttk.Entry(parent, textvariable=var,
                      font=("Yu Gothic UI", 12)).grid(
                row=row, column=0, sticky="ew", pady=(0, 5))
            row += 1

        # ── 選択中の要素 ──
        ttk.Separator(parent, orient="horizontal").grid(
            row=row, column=0, sticky="ew", pady=7)
        row += 1

        self.sel_label = ttk.Label(
            parent, text="選択中: なし",
            font=("Yu Gothic UI", 9, "bold"), foreground="#0078D7")
        self.sel_label.grid(row=row, column=0, sticky="w")
        row += 1

        sf = ttk.Frame(parent)
        sf.grid(row=row, column=0, sticky="ew", pady=(3, 3))
        ttk.Label(sf, text="文字サイズ (mm):").pack(side="left")
        self.size_spin = ttk.Spinbox(
            sf, from_=2.0, to=80.0, increment=0.5,
            textvariable=self.size_mm_var,
            width=7, font=("Yu Gothic UI", 10), state="disabled"
        )
        self.size_spin.pack(side="left", padx=6)
        row += 1

        ttk.Button(parent, text="⟳  位置・サイズをリセット",
                   command=self._reset_positions).grid(
            row=row, column=0, sticky="ew", pady=(3, 10))
        row += 1

        # ── 用紙サイズ ──
        ttk.Label(parent, text="用紙サイズ",
                  font=("Yu Gothic UI", 10, "bold")).grid(
            row=row, column=0, sticky="w", pady=(0, 3))
        row += 1
        ttk.Combobox(
            parent, textvariable=self.paper_size,
            values=list(PAPER_SIZES.keys()),
            state="readonly", font=("Yu Gothic UI", 10)
        ).grid(row=row, column=0, sticky="ew", pady=(0, 14))
        row += 1

        # ── ボタン ──
        bf = ttk.Frame(parent)
        bf.grid(row=row, column=0, sticky="ew")
        bf.columnconfigure(0, weight=1)
        bf.columnconfigure(1, weight=1)
        ttk.Button(bf, text="💾 のし紙を保存",
                   command=self._save_noshi).grid(
            row=0, column=0, sticky="ew", padx=(0, 3), ipady=4)
        ttk.Button(bf, text="📂 読み込む",
                   command=self._load_noshi).grid(
            row=0, column=1, sticky="ew", padx=(3, 0), ipady=4)
        ttk.Button(bf, text="🖨️ 印  刷",
                   command=self._print,
                   style="Accent.TButton").grid(
            row=1, column=0, columnspan=2, sticky="ew",
            pady=(4, 0), ipady=6)

    def _build_canvas(self, parent):
        cv = tk.Canvas(parent, bg="#AAAAAA",
                       relief="sunken", bd=1, highlightthickness=0)
        cv.grid(row=0, column=0, sticky="nsew")
        self.icanvas = InteractiveCanvas(cv)

    # ── フォント絞り込み・参照 ───────────────────────
    def _on_font_key(self, event):
        typed = self._font_combo.get()
        filtered = [n for n in ALL_FONT_NAMES if typed.lower() in n.lower()]
        self._font_combo["values"] = filtered if filtered else ALL_FONT_NAMES

    def _on_font_focus_in(self, event):
        """フォーカス時にリストをリセットして全選択"""
        self._font_combo["values"] = ALL_FONT_NAMES
        self._font_combo.select_range(0, tk.END)

    def _browse_font(self):
        """フォントファイルを直接参照して追加"""
        path = filedialog.askopenfilename(
            title="フォントファイルを選択",
            filetypes=[
                ("フォントファイル", "*.ttf *.otf *.ttc"),
                ("すべてのファイル", "*.*"),
            ],
            initialdir=os.path.join(
                os.environ.get("SystemRoot", r"C:\Windows"), "Fonts")
        )
        if not path:
            return
        name = os.path.splitext(os.path.basename(path))[0]
        FONT_REGISTRY[name] = path
        if name not in ALL_FONT_NAMES:
            ALL_FONT_NAMES.append(name)
            ALL_FONT_NAMES.sort()
        self._font_combo["values"] = ALL_FONT_NAMES
        self.font_name.set(name)

    # ── 選択コールバック ─────────────────────────────
    def _on_elem_selected(self, elem):
        self._selected_elem = elem
        if elem:
            _, h_mm = PAPER_SIZES[self.paper_size.get()]
            size_mm = round(elem.size_frac * h_mm, 1)
            self._lock_size_cb = True
            self.size_mm_var.set(f"{size_mm:.1f}")
            self._lock_size_cb = False
            self.sel_label.config(text=f"選択中: {elem.label}")
            self.size_spin.config(state="normal")
            self.status_var.set(
                f"【{elem.label}】  "
                f"X:{elem.x_frac:.2f}  Y:{elem.y_frac:.2f}  "
                f"サイズ: {size_mm:.1f}mm"
            )
        else:
            self.sel_label.config(text="選択中: なし")
            self.size_spin.config(state="disabled")
            self.status_var.set("要素をクリックして選択してください")

    def _on_spacing_changed(self, *_):
        try:
            v = self.char_spacing.get()
            self._spacing_label.config(text=f"{v:.2f}")
        except Exception:
            pass

    def _on_size_spin_changed(self, *_):
        if self._lock_size_cb or not self._selected_elem:
            return
        try:
            mm = float(self.size_mm_var.get())
        except ValueError:
            return
        _, h_mm = PAPER_SIZES[self.paper_size.get()]
        self._selected_elem.size_frac = mm / h_mm
        self._redraw()

    # ── 描画 ─────────────────────────────────────────
    def _schedule_render(self):
        if self._after_id:
            self.root.after_cancel(self._after_id)
        self._after_id = self.root.after(150, self._redraw)

    def _redraw(self):
        self._after_id = None
        self.icanvas.set_font(self.font_name.get())
        try:
            self.icanvas.set_spacing(self.char_spacing.get())
        except Exception:
            pass
        w_mm, h_mm = PAPER_SIZES[self.paper_size.get()]
        self.icanvas.render_all(w_mm, h_mm)

    def _update_name_position(self, *_):
        name_elem = next(e for e in self.elements if e.label == "氏名")
        if self.company.get() or self.position.get():
            name_elem.x_frac = 0.29
        else:
            name_elem.x_frac = 0.50

    def _reset_positions(self):
        for elem in self.elements:
            elem.reset()
        self._update_name_position()
        self._redraw()

    # ── 印刷用画像生成（水引なし） ────────────────────
    def _make_image(self):
        key = self.paper_size.get()
        w_mm, h_mm = PAPER_SIZES[key]
        W = mm_to_px(w_mm, PRINT_DPI)
        H = mm_to_px(h_mm, PRINT_DPI)
        img  = Image.new("RGB", (W, H), "white")
        draw = ImageDraw.Draw(img)
        draw.rectangle([1, 1, W-2, H-2], outline="#BBBBBB", width=1)

        fk      = self.font_name.get()
        spacing = self.char_spacing.get()
        for elem in self.elements:
            if not elem.text:
                continue
            cx  = int(elem.x_frac * W)
            cy  = int(elem.y_frac * H)
            fsz = max(8, int(elem.size_frac * H))
            f   = get_pil_font(fk, fsz)
            draw_vertical_pil(draw, elem.text, cx, cy, f, spacing=spacing)

        return img   # ← 水引は描画しない

    # ── 保存・印刷 ─────────────────────────────────────
    # ── のし紙を保存（JSON）────────────────────────────
    def _save_noshi(self):
        os.makedirs(SAVE_DIR, exist_ok=True)

        data = {
            "version":      1,
            "paper_size":   self.paper_size.get(),
            "font_name":    self.font_name.get(),
            "char_spacing": self.char_spacing.get(),
            "elements": {
                elem.label: {
                    "text":      elem.text,
                    "x_frac":    elem.x_frac,
                    "y_frac":    elem.y_frac,
                    "size_frac": elem.size_frac,
                }
                for elem in self.elements
            }
        }

        fname = self._noshi_filename()
        path  = os.path.join(SAVE_DIR, fname)

        # 同名ファイルが存在する場合は連番付加
        if os.path.exists(path):
            base, ext = os.path.splitext(fname)
            n = 2
            while os.path.exists(os.path.join(SAVE_DIR, f"{base}_{n}{ext}")):
                n += 1
            fname = f"{base}_{n}{ext}"
            path  = os.path.join(SAVE_DIR, fname)

        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        messagebox.showinfo("保存完了",
                            f"のし紙を保存しました：\n{fname}\n\n保存先：{SAVE_DIR}")

    def _noshi_filename(self):
        """保存ファイル名：用紙名-表書き-氏名.json"""
        paper = self.paper_size.get().split()[0]           # 例「美濃版」
        uwa   = self._safe_name(self.upper_text.get()) or "表書きなし"
        name  = self._safe_name(self.name.get())      or "氏名なし"
        return f"{paper}-{uwa}-{name}.json"

    def _safe_name(self, s):
        """ファイル名に使えない文字を除去"""
        return re.sub(r'[\\/:*?"<>|\r\n\t]', '', s).strip().replace(' ', '_').replace('　', '')

    # ── のし紙を読み込む ─────────────────────────────
    def _load_noshi(self):
        os.makedirs(SAVE_DIR, exist_ok=True)
        files = sorted(
            [f for f in os.listdir(SAVE_DIR) if f.endswith(".json")],
            key=lambda f: os.path.getmtime(os.path.join(SAVE_DIR, f)),
            reverse=True
        )

        if not files:
            messagebox.showinfo("読み込む",
                                f"保存済みのし紙がありません。\n\n保存先：{SAVE_DIR}")
            return

        self._show_load_dialog(files)

    def _show_load_dialog(self, files):
        dlg = tk.Toplevel(self.root)
        dlg.title("のし紙を読み込む")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.focus_set()

        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="読み込むのし紙を選んでください",
                  font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        lb_frm = ttk.Frame(frm)
        lb_frm.pack(fill="both", expand=True, pady=(0, 10))

        sb = ttk.Scrollbar(lb_frm, orient="vertical")
        lb = tk.Listbox(lb_frm, font=("Yu Gothic UI", 10),
                        height=min(len(files), 12),
                        width=42,
                        yscrollcommand=sb.set,
                        selectmode="single")
        sb.config(command=lb.yview)
        sb.pack(side="right", fill="y")
        lb.pack(side="left", fill="both", expand=True)

        for f in files:
            lb.insert(tk.END, os.path.splitext(f)[0])   # 拡張子なし表示
        lb.selection_set(0)

        def do_load():
            sel = lb.curselection()
            if not sel:
                return
            fname = files[sel[0]]
            path  = os.path.join(SAVE_DIR, fname)
            try:
                with open(path, encoding="utf-8") as fp:
                    data = json.load(fp)
                self._apply_noshi_data(data)
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("読み込みエラー", str(e), parent=dlg)

        def do_delete():
            sel = lb.curselection()
            if not sel:
                return
            fname = files[sel[0]]
            if messagebox.askyesno("削除確認",
                                   f"「{os.path.splitext(fname)[0]}」を削除しますか？",
                                   parent=dlg):
                os.remove(os.path.join(SAVE_DIR, fname))
                files.pop(sel[0])
                lb.delete(sel[0])
                if not files:
                    dlg.destroy()

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(fill="x")
        ttk.Button(btn_frm, text="読み込む", style="Accent.TButton",
                   command=do_load).pack(side="left", padx=(0, 6), ipady=4, ipadx=8)
        ttk.Button(btn_frm, text="削除",
                   command=do_delete).pack(side="left", ipady=4, ipadx=6)
        ttk.Button(btn_frm, text="キャンセル",
                   command=dlg.destroy).pack(side="right", ipady=4, ipadx=6)

        lb.bind("<Double-Button-1>", lambda e: do_load())

    def _apply_noshi_data(self, data):
        """読み込んだ JSON をアプリに反映"""
        if data.get("paper_size") in PAPER_SIZES:
            self.paper_size.set(data["paper_size"])
        if data.get("font_name"):
            self.font_name.set(data["font_name"])
        if data.get("char_spacing") is not None:
            self.char_spacing.set(float(data["char_spacing"]))

        elem_map = {e.label: e for e in self.elements}
        text_var_map = {
            "表書き": self.upper_text,
            "社名":   self.company,
            "役職":   self.position,
            "氏名":   self.name,
        }
        for label, d in data.get("elements", {}).items():
            if label in elem_map:
                text_var_map[label].set(d.get("text", ""))
                elem_map[label].x_frac    = float(d.get("x_frac",    elem_map[label].x_frac))
                elem_map[label].y_frac    = float(d.get("y_frac",    elem_map[label].y_frac))
                elem_map[label].size_frac = float(d.get("size_frac", elem_map[label].size_frac))

        self._redraw()

    def _print(self):
        try:
            img = self._make_image()
            tmp = tempfile.NamedTemporaryFile(
                suffix=".png", delete=False, prefix="noshi_")
            tmp.close()
            img.save(tmp.name, dpi=(PRINT_DPI, PRINT_DPI))

            # ── トレイ5（手差し）を自動設定（pywin32が入っている場合）──
            tray_ok = False
            if HAS_WIN32PRINT:
                try:
                    pname = win32print.GetDefaultPrinter()
                    hp    = win32print.OpenPrinter(pname)
                    try:
                        props = win32print.GetPrinter(hp, 2)
                        props['pDevMode'].DefaultSource = PRINT_BIN
                        win32print.SetPrinter(hp, 2, props, 0)
                        tray_ok = True
                    finally:
                        win32print.ClosePrinter(hp)
                except Exception as _e:
                    print(f"[トレイ設定スキップ] {_e}")

            os.startfile(tmp.name, "print")
            self._show_print_confirm_dialog(tray_ok)
        except Exception as e:
            messagebox.showerror("印刷エラー", str(e))

    # ── 印刷確認ダイアログ（給紙イメージ付き）────────────
    def _show_print_confirm_dialog(self, tray_ok=False):
        key  = self.paper_size.get()
        dlg  = tk.Toplevel(self.root)
        dlg.title("印刷")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.focus_set()

        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)

        # 確認事項テキスト
        if tray_ok:
            tray_line = "・給紙: トレイ5（手差し）✓ 自動設定済み"
        else:
            tray_line = ("・給紙: トレイ5（手差し）\n"
                         "  ※ 印刷ダイアログで「手差し」を手動選択してください")

        info = (
            "印刷ダイアログを開きました。\n\n"
            "【確認事項】\n"
            "・用紙の向き: 横向き\n"
            f"・用紙サイズ: {key}\n"
            "・印刷サイズ: 実際のサイズ（100%）\n"
            f"{tray_line}\n"
            "※ 水引はプレビューのみ（印刷されません）"
        )
        ttk.Label(frm, text=info,
                  font=("Yu Gothic UI", 10), justify="left"
                  ).pack(anchor="w", pady=(0, 10))

        ttk.Separator(frm, orient="horizontal").pack(fill="x", pady=(0, 8))

        ttk.Label(frm,
                  text="【給紙方法】  手差しトレイに「裏返し・上辺を手前」にセット",
                  font=("Yu Gothic UI", 10, "bold")
                  ).pack(anchor="w", pady=(0, 6))

        cv = tk.Canvas(frm, width=400, height=305,
                       bg="#F0F0F0",
                       highlightthickness=1, highlightbackground="#CCCCCC")
        cv.pack(pady=(0, 12))
        self._draw_feed_guide(cv, 400, 305)

        ttk.Button(frm, text="   OK   ", style="Accent.TButton",
                   command=dlg.destroy).pack()

    # ── 給紙イメージ図 ────────────────────────────────
    def _draw_feed_guide(self, cv, W, H):
        """
        正面図：スロットに向かって
          左 = 下辺 / 右 = 上辺（表書き側）/ 裏向き
        """
        BLUE  = "#0078D7"
        RED   = "#CC0000"
        GRAY1 = "#444444"

        # ─ プリンター本体（上部・横長スロット）─
        cv.create_rectangle(10, 10, W-10, 58,
                            fill=GRAY1, outline="#222222", width=2)
        cv.create_text(W//2, 22,
                       text="プリンター本体（正面）",
                       fill="#AAAAAA", font=("Yu Gothic UI", 8))
        # スロット口
        cv.create_rectangle(28, 31, W-28, 52,
                            fill="#111111", outline="#000000")
        cv.create_text(W//2, 41,
                       text="手差しスロット（トレイ 5）",
                       fill="#DDDDDD", font=("Yu Gothic UI", 9))

        # ─ 矢印（↑ スロットへ紙が入る方向）─
        arrow_y0 = 85
        arrow_y1 = 61
        for dx in (-80, -40, 0, 40, 80):
            x = W//2 + dx
            cv.create_line(x, arrow_y0, x, arrow_y1,
                           arrow=tk.LAST, fill=BLUE, width=2,
                           arrowshape=(7, 9, 3))

        # ─ 用紙（横長・裏向き）─
        pw = 240                         # 横（394mm 側）
        ph = int(pw * 267 / 394)         # 縦（267mm 側）≈ 162
        px0 = (W - pw) // 2             # 左端 x
        px1 = px0 + pw                  # 右端 x
        ptop = 90
        pbot = ptop + ph                 # ≈ 252

        # 影
        cv.create_rectangle(px0+4, ptop+4, px1+4, pbot+4,
                            fill="#BBBBBB", outline="")
        # 用紙
        cv.create_rectangle(px0, ptop, px1, pbot,
                            fill="#FFFDE7", outline="#999999", width=1)

        # 裏面ラベル
        mid_y = (ptop + pbot) // 2
        cv.create_text(W//2, mid_y - 14,
                       text="裏　面",
                       font=("Yu Gothic UI", 14, "bold"), fill="#555555")
        cv.create_text(W//2, mid_y + 10,
                       text="（印刷面を下向きに）",
                       font=("Yu Gothic UI", 9), fill="#888888")

        # ─ 左辺 = 下辺 ─
        cv.create_line(px0, ptop, px0, pbot,
                       fill=BLUE, width=3, dash=(6, 3))
        cv.create_text(px0 - 6, mid_y,
                       text="下\n辺",
                       font=("Yu Gothic UI", 10, "bold"),
                       fill=BLUE, anchor="e", justify="center")

        # ─ 右辺 = 上辺（表書き）─
        cv.create_line(px1, ptop, px1, pbot,
                       fill=RED, width=3)
        cv.create_text(px1 + 6, mid_y,
                       text="上辺\n表書き",
                       font=("Yu Gothic UI", 10, "bold"),
                       fill=RED, anchor="w", justify="center")

        # ─ 下部の凡例 ─
        cv.create_text(W//2, pbot + 16,
                       text="← スロットに向かって  左：下辺　右：上辺（表書き）→",
                       font=("Yu Gothic UI", 9), fill="#444444")
        cv.create_text(W//2, pbot + 34,
                       text="裏返し（印刷面を下に向けてセット）",
                       font=("Yu Gothic UI", 9, "bold"), fill="#CC6600")


# ======================================================
def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    root.option_add("*Font", ("Yu Gothic UI", 10))

    # アプリアイコン設定
    _icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "noshi_icon.ico")
    if os.path.exists(_icon_path):
        try:
            root.iconbitmap(_icon_path)
        except Exception:
            pass

    root.state("zoomed")   # 起動時に最大化
    NoshiApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
