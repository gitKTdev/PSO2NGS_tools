# -*- coding: utf-8 -*-
import sys
import os
import tkinter.ttk
from tkinter import font

import win32gui
import win32ui
import win32con

import numpy as np
import json
import cv2
import statistics
import copy

# from memory_profiler import profile

global img_width, img_height, width, height
global xtrim_start, xtrim_span, ytrim_start, ytrim_span, span, mask_right, mask_left, mask_top, mask_bottom, fps
global target_idx, vr_log
global windc


# @profile(precision=4)
def window_capture(window_name: str):
    global width, height
    global windc
    if "windc" not in globals():
        # 現在アクティブなウィンドウ名を探す
        process_list = []

        def callback(handle, _):
            process_list.append(win32gui.GetWindowText(handle))
        win32gui.EnumWindows(callback, None)

        # ターゲットウィンドウ名を探す
        for process_name in process_list:
            if window_name in process_name:
                hnd = win32gui.FindWindow(None, process_name)
                break
        else:
            # 見つからなかったら画面全体を取得
            hnd = win32gui.GetDesktopWindow()

        # ウィンドウサイズ取得
        x0, y0, x1, y1 = win32gui.GetWindowRect(hnd)
        width = x1 - x0
        height = y1 - y0

        # ウィンドウのデバイスコンテキスト取得
        windc = win32gui.GetWindowDC(hnd)

    srcdc = win32ui.CreateDCFromHandle(windc)
    memdc = srcdc.CreateCompatibleDC()
    # デバイスコンテキストからピクセル情報コピー, bmp化
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(srcdc, width, height)
    memdc.SelectObject(bmp)
    memdc.BitBlt((0, 0), (width, height), srcdc, (0, 0), win32con.SRCCOPY)

    # bmpの書き出し
    img_monitor = np.frombuffer(bmp.GetBitmapBits(True), np.uint8).reshape(height, width, 4)

    # 後片付け
    memdc.DeleteDC()
    win32gui.DeleteObject(bmp.GetHandle())

    return img_monitor


def detect_color(img, color: str):
    # HSV色空間に変換
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)

    # HSVの値域
    if color == "green":
        hsv_min = np.array([60, 50, 130])
        hsv_max = np.array([80, 255, 255])
    elif color == "white":
        hsv_min = np.array([0, 0, 95])
        hsv_max = np.array([180, 50, 200])
    else:
        return None, None

    # マスク（255：対象色、0：対象色以外）
    mask = cv2.inRange(hsv, hsv_min, hsv_max)

    # マスキング処理
    masked_img = cv2.bitwise_and(img, img, mask=mask)
    del hsv, hsv_min, hsv_max

    return mask, masked_img


# @profile(precision=4)
def create_images():
    global xtrim_start, xtrim_span, ytrim_start, ytrim_span, mask_right, mask_left, mask_top, mask_bottom

    image_monitor = window_capture("PHANTASY STAR ONLINE 2 NEW GENESIS")

    # VRゲージをトリム
    trim_image_monitor = image_monitor[ytrim_start:ytrim_start + ytrim_span, xtrim_start:xtrim_start + xtrim_span]
    _, trim_image_vr = detect_color(trim_image_monitor, color="green")
    trim_image_vr_max, trim_image_vr_maxc = detect_color(trim_image_monitor, color="white")

    del image_monitor

    return trim_image_monitor, trim_image_vr, trim_image_vr_max, trim_image_vr_maxc


def mask_image(image, mask_color):
    global mask_right, mask_left, mask_top, mask_bottom
    # 左
    cv2.rectangle(image, (0, 0), (mask_left, len(image) - 1), mask_color, -1)
    # 右
    cv2.rectangle(image, (mask_right, 0), (len(image[0]) - 1, len(image) - 1), mask_color, -1)
    # 上
    cv2.rectangle(image, (0, 0), (len(image[0]) - 1, mask_top), mask_color, -1)
    # 下
    cv2.rectangle(image, (0, mask_bottom), (len(image[0]) - 1, len(image) - 1), mask_color, -1)


def calc_vr_percent(image_vr, image_vr_max):
    global mask_top, mask_bottom, target_idx, vr_log
    vr_max_padding = 2

    target_idx = round((mask_top + mask_bottom) / 2)
    row_span = [target_idx - 1, target_idx, target_idx + 1]

    try:
        vr_now = statistics.mode(
            [len([j for j in image_vr[i] if sum(j) > 0]) for i in row_span]
        )
        vr_max = statistics.mode(
            [max(np.diff([j for j in range(len(image_vr_max[i])) if image_vr_max[i][j] > 0])) for i in row_span]
        )

        if vr_max <= vr_max_padding:
            return "-", "-"

        vr_per = abs(max([min([round(vr_now / (vr_max - vr_max_padding) * 100, 2), 100]), 0]))
        print(vr_per, vr_log)

        if vr_log[0] is None:
            vr_log[0] = copy.deepcopy(vr_now)
            vr_per_pre = "-"
        else:
            if vr_log[0] != vr_now:
                vr_log[1] = copy.deepcopy(vr_now)
                vr_log.reverse()
            if vr_log[1] is not None:
                vr_per_pre = abs(max([min([round((vr_now - vr_log[1] + vr_log[0]) / (vr_max - vr_max_padding) * 100, 2), 100]), 0]))
            else:
                vr_per_pre = "-"

    except Exception:
        vr_per = "-"
        vr_per_pre = "-"
    return vr_per, vr_per_pre


# @profile(precision=4)
def realtime_calc_vr():
    global main, config_json
    global img_width, img_height, xtrim_start, xtrim_span, ytrim_start, ytrim_span, span, mask_right, mask_left, mask_top, mask_bottom, fps

    if config_json is None:
        label_text.set(f"設定ファイルがありません。"
                       f"\n以下手順を実行してください。"
                       f"\n①WindowCaptureController.exeを実行"
                       f"\n②設定をsaveしexitボタンを押下"
                       f"\n③Refreshボタンを押下")

        main.geometry("460x180")
        main.after(5000, realtime_calc_vr)
    else:
        if "fps" not in globals():
            img_width = config_json.get("img_width", 305)
            img_height = config_json.get("img_height", 30)
            xtrim_start = config_json.get("xtrim_start", 1610)
            xtrim_span = config_json.get("xtrim_span", 305)
            ytrim_start = config_json.get("ytrim_start", 285)
            ytrim_span = config_json.get("ytrim_span", 30)
            span = config_json.get("span", 5)
            mask_right = config_json.get("mask_right", 289)
            mask_left = config_json.get("mask_left", 45)
            mask_top = config_json.get("mask_top", 9)
            mask_bottom = config_json.get("mask_bottom", 17)
            fps = config_json.get("fps", 10)

        img_monitor, img_vr, img_vr_max, img_vr_maxc = create_images()
        # 余分な所をマスク
        mask_color = (0, 0, 0)
        mask_image(img_vr, mask_color)
        mask_image(img_vr_max, mask_color)
        # VR計算
        vr_percent, vr_percent_pre = calc_vr_percent(img_vr, img_vr_max)

        del img_monitor, img_vr, img_vr_max, img_vr_maxc

        label_text.set(f"VRゲージ残量: {vr_percent}%\n  (減少予測 >> {vr_percent_pre}%)")

        main.after(round(fps2ms(fps)), realtime_calc_vr)


def fps2ms(frame_per_sec):
    return 1 / frame_per_sec * 1000


def refresh():
    os.execv(sys.executable, ['python'] + sys.argv)


# 初期値定義
load_file = "./vrgmt_config.cfg"
vr_log = [None, None]
config_json = {}
try:
    with open(load_file, "r") as f:
        config_json = json.loads(f.read())
except Exception:
    config_json = None

# ウィンドウ作成
main = tkinter.Tk()
main.title("VRゲージ管理ツール")
main.geometry(f"270x90")
main.attributes("-topmost", True)

# フレーム定義
frame_0_0 = tkinter.ttk.Frame(main)
frame_1_0 = tkinter.ttk.Frame(main)
frame_0_0.grid(row=0, column=0, rowspan=1, sticky="w")
frame_1_0.grid(row=1, column=0, rowspan=1)

# フォント
label_font = font.Font(family='Arial', size=18, weight='bold')
label_font2 = font.Font(family='Arial', size=10, weight='bold')
label_font2_5 = font.Font(family='Arial', size=10)
label_font3 = font.Font(family='Arial', size=10)

# VR解析結果
label_text = tkinter.StringVar()
label_vr_percent = tkinter.Label(frame_1_0, fg="green", font=label_font, textvariable=label_text, justify="left")
label_vr_percent.grid(row=3, column=0, columnspan=5)

refresh_button = tkinter.ttk.Button(frame_0_0, text="Refresh", width=8, command=refresh, style="MyWidget.TButton")
refresh_button.grid(row=0, column=0, padx=2, pady=1)


# 描画色々
realtime_calc_vr()


# ループ
main.mainloop()
