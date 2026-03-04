"""
长截屏工具
- Ctrl+Alt+PrintScreen  →  框选区域 → 自动滚动截图
- ESC 或到达底部        →  停止，拼接，预览保存
- 系统托盘常驻后台
"""
from __future__ import annotations

import ctypes
import os
import threading
import time
from pathlib import Path

import tkinter as tk
from PIL import Image, ImageDraw, ImageTk
from dotenv import load_dotenv

# 加载项目根目录的 .env
load_dotenv(Path(__file__).parent / ".env")


_SOUND_FILE = str(Path(__file__).parent / "提示音.mp3")
_MCI = ctypes.windll.winmm.mciSendStringW


def _play_sound():
    """用 Windows MCI 播放 MP3，无需额外依赖"""
    if not Path(_SOUND_FILE).exists():
        return
    _MCI(f'open "{_SOUND_FILE}" type mpegvideo alias _snd', None, 0, None)
    _MCI('play _snd wait', None, 0, None)
    _MCI('close _snd', None, 0, None)


def _get_output_dir() -> str:
    """读取 .env 中的 output_dir，不存在则返回用户桌面"""
    d = os.getenv("output_dir", "")
    if d:
        Path(d).mkdir(parents=True, exist_ok=True)
        return d
    return str(Path.home() / "Desktop")

import keyboard
import pyautogui
import numpy as np
import cv2
import pystray

# ── DPI 感知（必须最早设置，否则坐标会错位）────────────────────────
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)   # Per-Monitor DPI Aware
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

pyautogui.FAILSAFE = False

# ── 底层滚轮常量 ──────────────────────────────────────────────────
MOUSEEVENTF_MOVE      = 0x0001
MOUSEEVENTF_LEFTDOWN  = 0x0002
MOUSEEVENTF_LEFTUP    = 0x0004
MOUSEEVENTF_WHEEL     = 0x0800
WHEEL_DELTA           = 120


def _move(x: int, y: int):
    ctypes.windll.user32.SetCursorPos(x, y)


def _click(x: int, y: int):
    """左键单击（让目标窗口获得焦点）"""
    _move(x, y)
    time.sleep(0.05)
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.03)
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP,   0, 0, 0, 0)


def _scroll_down(x: int, y: int, clicks: int = 3):
    """用 ctypes 直接发送鼠标滚轮事件（比 pyautogui 可靠）"""
    _move(x, y)
    delta = ctypes.c_uint32(-WHEEL_DELTA * clicks).value   # 负值 = 向下
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_WHEEL, 0, 0, delta, 0)


class LongScreenshot:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

        self.is_capturing = False
        self.stop_flag = threading.Event()
        self.screenshots: list[Image.Image] = []
        self._indicator: tk.Toplevel | None = None

        # 全局快捷键
        keyboard.add_hotkey("ctrl+alt+print screen", self._on_long_hotkey,     suppress=True)
        keyboard.add_hotkey("shift+print screen",    self._on_rect_hotkey,     suppress=True)
        keyboard.add_hotkey("print screen",          self._on_fullscreen_hotkey, suppress=True)
        self._start_tray()

    # ─────────────────────────── 快捷键触发 ───────────────────────────
    def _on_long_hotkey(self):
        if not self.is_capturing:
            self.root.after(0, lambda: self._show_selection(self._start_long))

    def _on_rect_hotkey(self):
        if not self.is_capturing:
            self.root.after(0, lambda: self._show_selection(self._take_rect_shot))

    def _on_fullscreen_hotkey(self):
        if not self.is_capturing:
            threading.Thread(target=self._take_fullscreen, daemon=True).start()

    # ─────────────────────────── 区域选择覆盖层 ──────────────────────
    def _show_selection(self, on_region):
        """通用框选覆盖层，选完后调用 on_region(x, y, w, h)"""
        overlay = tk.Toplevel(self.root)
        overlay.attributes("-fullscreen", True)
        overlay.attributes("-alpha", 0.25)
        overlay.attributes("-topmost", True)
        overlay.configure(bg="black")

        tk.Label(
            overlay,
            text="  拖动鼠标框选截屏区域    ESC 取消  ",
            bg="#111", fg="white", font=("微软雅黑", 13), pady=6,
        ).place(relx=0.5, rely=0.01, anchor="n")

        canvas = tk.Canvas(overlay, cursor="cross", highlightthickness=0, bg="black")
        canvas.pack(fill="both", expand=True)

        sx = sy = 0
        rect = [None]

        def press(e):
            nonlocal sx, sy
            sx, sy = e.x, e.y

        def drag(e):
            if rect[0]:
                canvas.delete(rect[0])
            rect[0] = canvas.create_rectangle(
                sx, sy, e.x, e.y,
                outline="#00e676", width=2, fill="white", stipple="gray12",
            )

        def release(e):
            x1, y1 = min(sx, e.x), min(sy, e.y)
            x2, y2 = max(sx, e.x), max(sy, e.y)
            overlay.destroy()
            if x2 - x1 > 20 and y2 - y1 > 20:
                region = (x1, y1, x2 - x1, y2 - y1)
                threading.Thread(target=on_region, args=(region,), daemon=True).start()

        canvas.bind("<ButtonPress-1>",   press)
        canvas.bind("<B1-Motion>",       drag)
        canvas.bind("<ButtonRelease-1>", release)
        overlay.bind("<Escape>", lambda _: overlay.destroy())
        overlay.focus_force()

    # ─────────────────────────── 矩形截图 ────────────────────────────
    def _take_rect_shot(self, region: tuple[int, int, int, int]):
        time.sleep(0.3)   # 等覆盖层消失
        img = pyautogui.screenshot(region=region)
        self._save_image(img)

    # ─────────────────────────── 全屏截图 ────────────────────────────
    def _take_fullscreen(self):
        time.sleep(0.1)
        img = pyautogui.screenshot()
        self._save_image(img)

    # ─────────────────────────── 长截屏入口 ──────────────────────────
    def _start_long(self, region: tuple[int, int, int, int]):
        self._run_capture(region)

    # ─────────────────────────── 截图主循环 ──────────────────────────
    def _run_capture(self, region: tuple[int, int, int, int]):
        self.is_capturing = True
        self.stop_flag.clear()
        self.screenshots.clear()

        x, y, w, h = region
        cx, cy = x + w // 2, y + h // 2

        self.root.after(0, self._show_indicator)

        # 等覆盖层完全消失，再点击目标窗口获取焦点
        time.sleep(0.5)
        _click(cx, cy)
        time.sleep(0.3)

        # 首张截图
        self.screenshots.append(pyautogui.screenshot(region=(x, y, w, h)))

        dup_count = 0
        while not self.stop_flag.is_set():

            if keyboard.is_pressed("esc"):
                break

            # ★ 用 ctypes 直接发送滚轮事件
            _scroll_down(cx, cy, clicks=3)
            time.sleep(0.4)   # 等页面渲染

            shot = pyautogui.screenshot(region=(x, y, w, h))

            # 到底检测：连续 2 次与上张相同则停止
            if self._is_same(self.screenshots[-1], shot):
                dup_count += 1
                if dup_count >= 2:
                    break
            else:
                dup_count = 0
                self.screenshots.append(shot)

        self.root.after(0, self._hide_indicator)
        self.is_capturing = False

        if self.screenshots:
            self.root.after(0, self._on_capture_done)

    # ─────────────────────────── 悬浮指示器 ──────────────────────────
    def _show_indicator(self):
        win = tk.Toplevel(self.root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.attributes("-alpha", 0.90)
        win.geometry("+8+8")
        tk.Label(
            win, text="  ⏺ 截图中…  按 ESC 结束  ",
            bg="#b71c1c", fg="white", font=("微软雅黑", 11, "bold"), pady=5,
        ).pack()
        self._indicator = win

    def _hide_indicator(self):
        if self._indicator:
            self._indicator.destroy()
            self._indicator = None

    # ─────────────────────────── 相似度检测 ──────────────────────────
    @staticmethod
    def _is_same(a: Image.Image, b: Image.Image, thresh: float = 0.997) -> bool:
        ag = cv2.cvtColor(np.array(a), cv2.COLOR_RGB2GRAY).astype(np.float32)
        bg = cv2.cvtColor(np.array(b), cv2.COLOR_RGB2GRAY).astype(np.float32)
        if ag.shape != bg.shape:
            return False
        na, nb = np.linalg.norm(ag), np.linalg.norm(bg)
        if na == 0 or nb == 0:
            return False
        return float(np.dot(ag.flatten(), bg.flatten()) / (na * nb)) >= thresh

    # ─────────────────────────── 图像拼接 ────────────────────────────
    def _stitch(self) -> Image.Image | None:
        shots = self.screenshots
        if not shots:
            return None
        if len(shots) == 1:
            return shots[0]

        arrays = [np.array(s) for s in shots]
        result = arrays[0]

        for i in range(1, len(arrays)):
            pg = cv2.cvtColor(arrays[i - 1], cv2.COLOR_RGB2GRAY)
            cg = cv2.cvtColor(arrays[i],     cv2.COLOR_RGB2GRAY)
            h = pg.shape[0]
            tpl_h = max(h // 3, 60)
            template = pg[h - tpl_h:, :]
            search_h = min(int(h * 0.85), cg.shape[0])
            search   = cg[:search_h, :]

            if template.shape[1] != search.shape[1] or template.shape[0] >= search.shape[0]:
                result = np.vstack([result, arrays[i]])
                continue

            res = cv2.matchTemplate(search, template, cv2.TM_CCOEFF_NORMED)
            _, val, _, loc = cv2.minMaxLoc(res)

            if val > 0.4:
                new_start = loc[1] + tpl_h
                if 0 < new_start < arrays[i].shape[0]:
                    result = np.vstack([result, arrays[i][new_start:]])
            else:
                result = np.vstack([result, arrays[i]])

        return Image.fromarray(result)

    # ─────────────────────────── 保存 & 提示音（公共）────────────────
    @staticmethod
    def _save_image(img: Image.Image):
        ts   = time.strftime("%Y%m%d_%H%M%S")
        name = f"screenshot_{ts}.png"
        path = str(Path(_get_output_dir()) / name)
        img.save(path)
        threading.Thread(target=_play_sound, daemon=True).start()

    # ─────────────────────────── 长截屏完成 ──────────────────────────
    def _on_capture_done(self):
        img = self._stitch()
        if img is not None:
            self._save_image(img)

    # ─────────────────────────── 系统托盘 ────────────────────────────
    def _make_icon(self) -> Image.Image:
        img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.rounded_rectangle([6, 18, 58, 54], radius=8, fill="#1565c0")
        d.ellipse([18, 24, 46, 48], fill="white")
        d.ellipse([24, 30, 40, 44], fill="#1565c0")
        d.rounded_rectangle([30, 12, 46, 20], radius=3, fill="#1565c0")
        return img

    def _start_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem(
                "全屏截图  (PrtSc)",
                lambda: threading.Thread(target=self._take_fullscreen, daemon=True).start(),
            ),
            pystray.MenuItem(
                "矩形截图  (Shift+PrtSc)",
                lambda: self.root.after(0, lambda: self._show_selection(self._take_rect_shot)),
            ),
            pystray.MenuItem(
                "长截屏    (Ctrl+Alt+PrtSc)",
                lambda: self.root.after(0, lambda: self._show_selection(self._start_long)),
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("退出 (&X)", self._quit),
        )
        self.tray = pystray.Icon("长截屏工具", self._make_icon(), "长截屏工具", menu)
        threading.Thread(target=self.tray.run, daemon=True).start()

    def _quit(self):
        keyboard.unhook_all()
        self.tray.stop()
        self.root.after(0, self.root.quit)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = LongScreenshot()
    app.run()
