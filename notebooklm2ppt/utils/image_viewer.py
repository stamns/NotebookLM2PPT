# -*- coding: utf-8 -*-
# D:/research/ICCV/vis_pixel_fake_rebuttal3/具象抽象_模型建构_Page1.png

# 使用 OpenCV 打开图片，按屏幕比例缩放并全屏显示；无边框/工具栏。

import os
import sys
import ctypes
import cv2
import numpy as np
from PIL import Image


def _get_screen_resolution():
    # Windows 获取屏幕分辨率，并启用 DPI 感知避免缩放影响
    try:
        user32 = ctypes.windll.user32
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass
        return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
    except Exception:
        return 1920, 1080


def show_image_fullscreen(image_path: str, display_height: int = None):
    """
    显示图片在屏幕左上角
    
    Args:
        image_path: 图片路径
        display_height: 指定显示高度（像素），如果为None则自动适配屏幕
    """
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"未找到图片: {image_path}")
    img = Image.open(image_path).convert("RGB")
    # img = cv2.imread(image_path, cv2.IMREAD_UNCHANGED)
    img = np.array(img)[...,::-1]  # RGB to BGR
    if img is None:
        raise ValueError(f"无法读取图片: {image_path}")

    screen_w, screen_h = _get_screen_resolution()
    h, w = img.shape[:2]

    # 根据指定高度或屏幕大小计算缩放比例
    if display_height is not None:
        scale = display_height / h
    else:
        scale = min(screen_w / w, screen_h / h)
    
    new_w = max(1, int(w * scale))
    new_h = max(1, int(h * scale))

    interp = cv2.INTER_AREA if scale < 1.0 else cv2.INTER_CUBIC
    resized = cv2.resize(img, (new_w, new_h), interpolation=interp)

    # 放置到左上角（不居中）
    pad_left = 0
    pad_right = max(0, screen_w - new_w)
    pad_top = 0
    pad_bottom = max(0, screen_h - new_h)
    canvas = cv2.copyMakeBorder(
        resized, pad_top, pad_bottom, pad_left, pad_right,
        borderType=cv2.BORDER_CONSTANT, value=(0, 0, 0)
    )

    win_name = "__opencv_fullscreen__"
    cv2.namedWindow(win_name, cv2.WINDOW_NORMAL)
    cv2.setWindowProperty(win_name, cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
    cv2.imshow(win_name, canvas)

    # 置顶窗口（Windows）
    try:
        hwnd = ctypes.windll.user32.FindWindowW(None, win_name)
        if hwnd:
            # SetWindowPos(HWND_TOPMOST= -1)
            ctypes.windll.user32.SetWindowPos(
                hwnd,
                -1,
                0,
                0,
                0,
                0,
                0x0002 | 0x0001  # SWP_NOSIZE | SWP_NOMOVE
            )
            # 强制激活窗口，避免第二次被压到后台
            ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            ctypes.windll.user32.BringWindowToTop(hwnd)
    except Exception:
        pass

    # # 按任意键或 ESC 退出
    # key = cv2.waitKey(0)
    # if key == 27:  # ESC
    #     pass
    # cv2.destroyAllWindows()


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "具象抽象_模型建构_Page1.png"
    # 如果提供了第二个参数，则作为显示高度
    height = int(sys.argv[2]) if len(sys.argv) > 2 else None
    show_image_fullscreen(path, display_height=height)