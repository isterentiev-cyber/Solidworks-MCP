# -*- coding: utf-8 -*-
"""Verify the BUSY branch of ext.sw_busy without involving SolidWorks.

A window whose owning thread never pumps its message queue looks, from the
outside, exactly like SolidWorks with a modal dialog open: SendMessageTimeout
cannot deliver, and must come back with ERROR_TIMEOUT instead of blocking.
"""
import ctypes
import ctypes.wintypes as w
import threading
import time
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))
from solidworks_mcp import ext

u32 = ctypes.WinDLL("user32", use_last_error=True)
k32 = ctypes.WinDLL("kernel32", use_last_error=True)
WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_long, w.HWND, w.UINT, w.WPARAM, w.LPARAM)


class WNDCLASS(ctypes.Structure):
    _fields_ = [("style", w.UINT), ("lpfnWndProc", WNDPROC),
                ("cbClsExtra", ctypes.c_int), ("cbWndExtra", ctypes.c_int),
                ("hInstance", w.HINSTANCE), ("hIcon", w.HICON),
                ("hCursor", w.HANDLE), ("hbrBackground", w.HBRUSH),
                ("lpszMenuName", w.LPCWSTR), ("lpszClassName", w.LPCWSTR)]


u32.DefWindowProcW.restype = ctypes.c_long
u32.DefWindowProcW.argtypes = [w.HWND, w.UINT, w.WPARAM, w.LPARAM]
proc = WNDPROC(lambda h, m, wp, lp: u32.DefWindowProcW(h, m, wp, lp))

k32.GetModuleHandleW.restype = w.HMODULE
k32.GetModuleHandleW.argtypes = [w.LPCWSTR]
u32.CreateWindowExW.restype = w.HWND
u32.CreateWindowExW.argtypes = [w.DWORD, w.LPCWSTR, w.LPCWSTR, w.DWORD,
                                ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                ctypes.c_int, w.HWND, w.HMENU, w.HINSTANCE,
                                w.LPVOID]

box, ready = {}, threading.Event()


def owner():
    wc = WNDCLASS()
    wc.lpfnWndProc = proc
    wc.lpszClassName = "SwMcpProbeTest"
    wc.hInstance = k32.GetModuleHandleW(None)
    if not u32.RegisterClassW(ctypes.byref(wc)):
        box["err"] = ctypes.get_last_error()
        ready.set()
        return
    box["h"] = u32.CreateWindowExW(0, "SwMcpProbeTest", "probe", 0,
                                   0, 0, 10, 10, None, None, wc.hInstance, None)
    box["err"] = ctypes.get_last_error()
    ready.set()
    time.sleep(8)  # deliberately no message loop -- this is the whole point


threading.Thread(target=owner, daemon=True).start()
ready.wait(3)
hwnd = box.get("h")
print("test window hwnd:", hwnd, " last_error:", box.get("err"))
print("IsWindow:", bool(u32.IsWindow(w.HWND(hwnd))) if hwnd else False)

if not hwnd:
    print("\nFAILED to create the test window -- cannot verify")
    sys.exit(1)


class FakeApp:
    pass


app = FakeApp()
ext._HWND_CACHE[id(app)] = hwnd  # aim the probe at the wedged window

t0 = time.perf_counter()
verdict = ext.sw_busy(app, 800)
dt = (time.perf_counter() - t0) * 1000

print("\nsw_busy(timeout=800ms) took %.0f ms" % dt)
print("verdict:", verdict)
print()
print("BUSY detected      :", verdict is not None)
print("did not block      :", dt < 1500)
print("respects the budget:", 700 < dt < 1500)
sys.exit(0 if verdict is not None and dt < 1500 else 1)
