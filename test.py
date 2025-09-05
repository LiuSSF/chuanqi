# find_and_click_image_loop.py
# 在指定窗口客户区里匹配两张模板中的“任意一张”，命中就把鼠标移到中心 -> 停 2 秒 -> 左键点击
import time
import ctypes
from ctypes import wintypes
import numpy as np
import cv2
import mss
import win32gui
import win32con

# ---------- 兼容: 部分环境没有 wintypes.ULONG_PTR ----------
if not hasattr(wintypes, "ULONG_PTR"):
    if ctypes.sizeof(ctypes.c_void_p) == ctypes.sizeof(ctypes.c_ulonglong):
        wintypes.ULONG_PTR = ctypes.c_ulonglong
    else:
        wintypes.ULONG_PTR = ctypes.c_ulong
# -----------------------------------------------------------

user32   = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
try:
    user32.SetProcessDPIAware()  # 避免高DPI坐标偏移
except Exception:
    pass

# ================= SendInput 鼠标 =================
INPUT_MOUSE    = 0
MOUSEEVENTF_MOVE        = 0x0001
MOUSEEVENTF_LEFTDOWN    = 0x0002
MOUSEEVENTF_LEFTUP      = 0x0004
MOUSEEVENTF_ABSOLUTE    = 0x8000
MOUSEEVENTF_VIRTUALDESK = 0x4000

class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx",          wintypes.LONG),
        ("dy",          wintypes.LONG),
        ("mouseData",   wintypes.DWORD),
        ("dwFlags",     wintypes.DWORD),
        ("time",        wintypes.DWORD),
        ("dwExtraInfo", wintypes.ULONG_PTR),
    ]

class _INPUTUNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT)]

class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("union", _INPUTUNION)]

SendInput = user32.SendInput
SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
SendInput.restype  = wintypes.UINT

def send_mouse_move_abs(ax, ay):
    inp = INPUT(); inp.type = INPUT_MOUSE
    mi = MOUSEINPUT()
    mi.dx, mi.dy = int(ax), int(ay)
    mi.mouseData = 0
    mi.dwFlags   = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK
    mi.time      = 0
    mi.dwExtraInfo = 0
    inp.union.mi = mi
    if SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp)) != 1:
        raise RuntimeError("SendInput MOVE failed")

def send_left_click():
    def _btn(flag):
        inp = INPUT(); inp.type = INPUT_MOUSE
        mi = MOUSEINPUT()
        mi.dx = mi.dy = 0; mi.mouseData = 0
        mi.dwFlags = flag; mi.time = 0; mi.dwExtraInfo = 0
        inp.union.mi = mi
        if SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp)) != 1:
            raise RuntimeError("SendInput CLICK failed")
    _btn(MOUSEEVENTF_LEFTDOWN); time.sleep(0.02); _btn(MOUSEEVENTF_LEFTUP)

# ================= 虚拟桌面/坐标换算 =================
def get_virtual_bounds():
    L = user32.GetSystemMetrics(76)   # SM_XVIRTUALSCREEN
    T = user32.GetSystemMetrics(77)   # SM_YVIRTUALSCREEN
    W = user32.GetSystemMetrics(78)   # SM_CXVIRTUALSCREEN
    H = user32.GetSystemMetrics(79)   # SM_CYVIRTUALSCREEN
    return L, T, L + W - 1, T + H - 1

def to_abs_on_virtual(x, y):
    L, T, R, B = get_virtual_bounds()
    w = max(1, R - L); h = max(1, B - T)
    ax = int((x - L) * 65535 / w); ay = int((y - T) * 65535 / h)
    return ax, ay

# ================= 窗口/截图工具 =================
def find_hwnd_contains(title_sub: str):
    target = None
    def _enum(hwnd, _):
        nonlocal target
        if not win32gui.IsWindowVisible(hwnd): return
        t = win32gui.GetWindowText(hwnd)
        if t and title_sub in t: target = hwnd
    win32gui.EnumWindows(_enum, None)
    return target

def bring_foreground_soft(hwnd, timeout=1.5):
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE); time.sleep(0.15)
    else:
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW);    time.sleep(0.05)
    try: user32.AllowSetForegroundWindow(ctypes.c_uint(-1).value)
    except Exception: pass
    try:
        fg = user32.GetForegroundWindow()
        cur_tid = kernel32.GetCurrentThreadId()
        fg_tid  = user32.GetWindowThreadProcessId(fg, None)
        tgt_tid = user32.GetWindowThreadProcessId(hwnd, None)
        user32.AttachThreadInput(cur_tid, fg_tid, True)
        user32.AttachThreadInput(cur_tid, tgt_tid, True)
        try:
            user32.SetForegroundWindow(hwnd)
            user32.BringWindowToTop(hwnd)
            try: user32.SwitchToThisWindow(hwnd, True)
            except Exception: pass
        finally:
            user32.AttachThreadInput(cur_tid, fg_tid, False)
            user32.AttachThreadInput(cur_tid, tgt_tid, False)
    except Exception:
        pass
    t0 = time.time()
    while time.time() - t0 < timeout:
        if user32.GetForegroundWindow() == hwnd: return True
        time.sleep(0.05)
    return False

def get_client_region(hwnd):
    Lc, Tc, Rc, Bc = win32gui.GetClientRect(hwnd)
    (sx0, sy0) = win32gui.ClientToScreen(hwnd, (Lc, Tc))
    (sx1, sy1) = win32gui.ClientToScreen(hwnd, (Rc, Bc))
    return {"left": sx0, "top": sy0, "width": sx1 - sx0, "height": sy1 - sy0}

def grab_gray(region):
    with mss.mss() as sct:
        img = np.array(sct.grab(region))
    return cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)

# ================= 模板匹配 =================
def load_template_gray(path):
    tmpl = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
    if tmpl is None:
        raise FileNotFoundError(f"读取模板失败：{path}")
    return tmpl

def match_template_multiscale(screen_gray, tmpl_gray, scales, method=cv2.TM_CCOEFF_NORMED):
    best = (-1, None, None, None)  # (score, top_left, w, h)
    H, W = screen_gray.shape[:2]
    for s in scales:
        th = int(round(tmpl_gray.shape[0] * s))
        tw = int(round(tmpl_gray.shape[1] * s))
        if th < 8 or tw < 8 or th > H or tw > W:
            continue
        resized = cv2.resize(tmpl_gray, (tw, th), interpolation=cv2.INTER_AREA if s < 1 else cv2.INTER_LINEAR)
        res = cv2.matchTemplate(screen_gray, resized, method)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        score, loc = (max_val, max_loc) if method in (cv2.TM_CCOEFF_NORMED, cv2.TM_CCORR_NORMED) else (1-min_val, min_loc)
        if score > best[0]:
            best = (score, loc, tw, th)
    return best  # (score, top_left, w, h)

# ================= 查找任意模板并点击（移动→停2秒→点击） =================
def find_any_and_move_click(template_paths, window_title, threshold=0.88,
                            scales=np.linspace(0.8,1.2,17), pause_seconds=2.0,
                            bring_to_front=True):
    hwnd = find_hwnd_contains(window_title)
    if not hwnd:
        raise RuntimeError(f"找不到窗口（标题包含）：{window_title}")
    if bring_to_front:
        bring_foreground_soft(hwnd)

    region = get_client_region(hwnd)
    screen_gray = grab_gray(region)

    tmpls = [(p, load_template_gray(p)) for p in template_paths]
    best_overall = (-1, None, None, None, None)  # score, tl, tw, th, which

    for idx, (p, tmpl) in enumerate(tmpls):
        score, tl, tw, th = match_template_multiscale(screen_gray, tmpl, scales)
        print(f"[调试] 模板 {p} 最佳分数 {score:.3f}")
        if score > best_overall[0]:
            best_overall = (score, tl, tw, th, idx)

    score, tl, tw, th, which = best_overall
    if score < threshold or tl is None:
        raise RuntimeError(f"未找到匹配度 >= {threshold} 的模板（最高 {score:.3f}）")

    # 命中点（窗口客户区 -> 屏幕 -> 绝对坐标）
    center_x = tl[0] + tw // 2
    center_y = tl[1] + th // 2
    screen_x = region["left"] + center_x
    screen_y = region["top"]  + center_y
    ax, ay   = to_abs_on_virtual(screen_x, screen_y)

    print(f"[命中] 模板 #{which+1} 分数 {score:.3f}，点击点 screen({screen_x},{screen_y}) -> abs({ax},{ay})")

    # 鼠标移动到命中位置 -> 停 2 秒 -> 左键点击
    send_mouse_move_abs(ax, ay)
    time.sleep(pause_seconds)
    send_left_click()

# ================= 配置与循环 =================
if __name__ == "__main__":
    TEMPLATE_1 = r"H:\pythontest\chuanqi\img1.png"
    TEMPLATE_2 = r"H:\pythontest\chuanqi\img2.png"
    WINDOW_TITLE_SUBSTR = "重生之旧梦04号05区 - 阿西吧啊"

    while True:
        try:
            find_any_and_move_click(
                [TEMPLATE_1, TEMPLATE_2],
                window_title=WINDOW_TITLE_SUBSTR,
                threshold=0.88,
                pause_seconds=0.2,              # 停 2 秒让你看到鼠标
                scales=np.linspace(0.8, 1.2, 17),
                bring_to_front=True             # 如需完全“后台不切窗”，改为 False（命中位置若被遮挡就点不到）
            )
            print("操作完成，10秒后再试...")
            time.sleep(1)
        except Exception as e:
            print(f"发生错误：{e}，5秒后重试...")
            time.sleep(3)
