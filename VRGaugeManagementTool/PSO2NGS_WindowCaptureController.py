# -*- coding: utf-8 -*-
import tkinter
import tkinter.ttk
from tkinter import font
from PIL import ImageTk, Image

import win32gui
import win32ui
import win32con

import sys
import numpy as np
import json
import cv2
import statistics

# from memory_profiler import profile

global img_width, img_height, width, height
global xtrim_start, xtrim_span, ytrim_start, ytrim_span, span, mask_right, mask_left, mask_top, mask_bottom
global target_idx, fps
global entry_span, label_title, entry_fps
global tk_img_monitor, tk_img_vr, tk_img_vr_max, cvmonit_id, cvvr_id, cvvrmax_id
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


def cv2_to_tk(cv2_image):
    """
    CV2 -> Tkinter
    """

    # BGR -> RGB
    rgb_cv2_image = cv2.cvtColor(cv2_image, cv2.COLOR_BGR2RGB)

    # NumPy配列からPIL画像オブジェクトを生成
    pil_image = Image.fromarray(rgb_cv2_image)

    # PIL画像オブジェクトをTkinter画像オブジェクトに変換
    tk_image = ImageTk.PhotoImage(pil_image)
    del rgb_cv2_image, pil_image

    return tk_image


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


def plot_judge_line(masked_image, color):
    global mask_right, mask_left, mask_top, mask_bottom, target_idx

    cv2.rectangle(masked_image, (mask_left, target_idx - 1), (mask_right, target_idx + 1), color, -1)


def calc_vr_percent(image_vr, image_vr_max):
    global mask_top, mask_bottom, target_idx
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
            return "-"

        vr_per = abs(max([min([round(vr_now / (vr_max - vr_max_padding) * 100, 2), 100]), 0]))

    except Exception:
        vr_per = "-"
    return vr_per


# @profile(precision=4)
def realtime_grid_image():
    global main, tk_img_monitor, tk_img_vr, tk_img_vr_max, label_text
    global label_vr_percent, canvas_monitor, canvas_vr, canvas_vr_max, cvmonit_id, cvvr_id, cvvrmax_id

    img_monitor, img_vr, img_vr_max, img_vr_maxc = create_images()
    # 余分な所をマスク
    mask_color = (0, 0, 0)
    mask_image(img_vr, mask_color)
    mask_image(img_vr_max, mask_color)
    # VR計算
    vr_percent = calc_vr_percent(img_vr, img_vr_max)
    # マスクを分かりやすく
    mask_color = (100, 100, 100)
    mask_image(img_vr, mask_color)
    mask_image(img_vr_maxc, mask_color)
    if bln_judge_line.get():
        judge_color = (0, 0, 255)
        plot_judge_line(img_vr, judge_color)
        plot_judge_line(img_vr_maxc, judge_color)

    # 変換(cv2->tk)
    tk_img_monitor = cv2_to_tk(img_monitor)
    tk_img_vr = cv2_to_tk(img_vr)
    tk_img_vr_max = cv2_to_tk(img_vr_maxc)
    del img_monitor, img_vr, img_vr_max, img_vr_maxc

    if "cvmonit_id" not in globals():
        cvmonit_id = None
    if "cvvr_id" not in globals():
        cvvr_id = None
    if "cvvrmax_id" not in globals():
        cvvrmax_id = None

    if cvmonit_id is None:
        cvmonit_id = canvas_monitor.create_image(0, 0, image=tk_img_monitor, anchor=tkinter.NW)
    else:
        canvas_monitor.itemconfigure(cvmonit_id, image=tk_img_monitor)
    if cvvr_id is None:
        cvvr_id = canvas_vr.create_image(0, 0, image=tk_img_vr, anchor=tkinter.NW)
    else:
        canvas_vr.itemconfigure(cvvr_id, image=tk_img_vr)
    if cvvrmax_id is None:
        cvvrmax_id = canvas_vr_max.create_image(0, 0, image=tk_img_vr_max, anchor=tkinter.NW)
    else:
        canvas_vr_max.itemconfigure(cvvrmax_id, image=tk_img_vr_max)
    canvas_monitor.grid()
    canvas_vr.grid()
    canvas_vr_max.grid()

    label_text.set(f"VRゲージ残量: {vr_percent}%")

    main.after(round(fps2ms(fps)), realtime_grid_image)


def fps2ms(frame_per_sec):
    return 1 / frame_per_sec * 1000


def change_fps():
    global entry_fps, fps

    fps = int(entry_fps.get())


def control_space_window():
    global span
    global entry_span, label_title, entry_fps

    row_idx = 0

    label_title = tkinter.Label(frame_0_0, font=label_font2, text="キャプチャ範囲調整")
    label_title.grid(row=row_idx, column=0, columnspan=6, sticky="w")
    row_idx += 1

    # save, exitボタン
    save_button = tkinter.ttk.Button(frame_0_0, text="save", width=4, command=save_config, style="MyWidget.TButton", default="active")
    save_button.grid(row=row_idx, column=0, columnspan=2, sticky="e")
    exit_button = tkinter.ttk.Button(frame_0_0, text="exit", width=4, command=exit_progrram, style="MyWidget.TButton")
    exit_button.grid(row=row_idx, column=2, columnspan=2, sticky="w")
    row_idx += 1

    # fps
    label_fps = tkinter.Label(frame_0_0, font=label_font2_5, text="FPS : ")
    label_fps.grid(row=row_idx, column=0, columnspan=3, sticky="w")
    entry_fps = tkinter.Entry(frame_0_0, width=5, font=("Meiryo", 9))
    entry_fps.grid(row=row_idx, column=3, columnspan=2)
    entry_fps.insert(tkinter.END, fps)
    label_fps2 = tkinter.Label(frame_0_0, font=label_font3, text="[fps]")
    label_fps2.grid(row=row_idx, column=5)
    change_fps_button = tkinter.ttk.Button(frame_0_0, text="反映", width=4, command=change_fps, style="MyWidget.TButton")
    change_fps_button.grid(row=row_idx, column=6)
    row_idx += 1

    # 図サイズ
    # 調整幅
    label_span = tkinter.Label(frame_0_0, font=label_font2_5, text="調整幅 : ")
    label_span.grid(row=row_idx, column=0, columnspan=3, sticky="w")
    entry_span = tkinter.Entry(frame_0_0, width=5, font=("Meiryo", 9))
    entry_span.grid(row=row_idx, column=3, columnspan=2)
    entry_span.insert(tkinter.END, span)
    label_span2 = tkinter.Label(frame_0_0, font=label_font3, text="[px]")
    label_span2.grid(row=row_idx, column=5)
    row_idx += 1
    # 右
    label_sizeup_width_right = tkinter.Label(frame_0_0, font=label_font2_5, text="右 : ")
    label_sizeup_width_right.grid(row=row_idx, column=0, columnspan=2)
    sizeup_width_right_button = tkinter.ttk.Button(
        frame_0_0, text="+", width=2, command=sizeup_width_right, style="MyWidget.TButton"
    )
    sizedown_width_right_button = tkinter.ttk.Button(
        frame_0_0, text="-", width=2, command=sizedown_width_right, style="MyWidget.TButton"
    )
    sizeup_width_right_button.grid(row=row_idx, column=2)
    sizedown_width_right_button.grid(row=row_idx, column=3)
    row_idx += 1
    # 左
    label_sizeup_width_left = tkinter.Label(frame_0_0, font=label_font2_5, text="左 : ")
    label_sizeup_width_left.grid(row=row_idx, column=0, columnspan=2)
    sizeup_width_left_button = tkinter.ttk.Button(
        frame_0_0, text="+", width=2, command=sizeup_width_left, style="MyWidget.TButton"
    )
    sizedown_width_left_button = tkinter.ttk.Button(
        frame_0_0, text="-", width=2, command=sizedown_width_left, style="MyWidget.TButton"
    )
    sizeup_width_left_button.grid(row=row_idx, column=2)
    sizedown_width_left_button.grid(row=row_idx, column=3)
    row_idx += 1
    # 上
    label_sizeup_height_top = tkinter.Label(frame_0_0, font=label_font2_5, text="上 : ")
    label_sizeup_height_top.grid(row=row_idx, column=0, columnspan=2)
    sizeup_height_top_button = tkinter.ttk.Button(
        frame_0_0, text="+", width=2, command=sizeup_height_top, style="MyWidget.TButton"
    )
    sizedown_height_top_button = tkinter.ttk.Button(
        frame_0_0, text="-", width=2, command=sizedown_height_top, style="MyWidget.TButton"
    )
    sizeup_height_top_button.grid(row=row_idx, column=2)
    sizedown_height_top_button.grid(row=row_idx, column=3)
    row_idx += 1
    # 下
    label_sizeup_height_bottom = tkinter.Label(frame_0_0, font=label_font2_5, text="下 : ")
    label_sizeup_height_bottom.grid(row=row_idx, column=0, columnspan=2)
    sizeup_height_bottom_button = tkinter.ttk.Button(
        frame_0_0, text="+", width=2, command=sizeup_height_bottom, style="MyWidget.TButton"
    )
    sizedown_height_bottom_button = tkinter.ttk.Button(
        frame_0_0, text="-", width=2, command=sizedown_height_bottom, style="MyWidget.TButton"
    )
    sizeup_height_bottom_button.grid(row=row_idx, column=2)
    sizedown_height_bottom_button.grid(row=row_idx, column=3)
    row_idx += 1

    # マスク範囲調整モード
    checkbox_mask_check = tkinter.Checkbutton(frame_0_0, variable=bln_mask_check, font=label_font2_5, text="マスク範囲調整")
    checkbox_mask_check.grid(row=row_idx, column=0, columnspan=14, sticky="w")
    row_idx += 1

    # 判定ラインプロットモード
    checkbox_judge_line = tkinter.Checkbutton(frame_0_0, variable=bln_judge_line, font=label_font2_5, text="判定ラインプロット")
    checkbox_judge_line.grid(row=row_idx, column=0, columnspan=14, sticky="w")
    row_idx += 1


def save_config():
    global img_width, img_height
    global xtrim_start, xtrim_span, ytrim_start, ytrim_span, span, mask_right, mask_left, mask_top, mask_bottom, fps

    # fpsを反映
    change_fps()

    config_json["img_width"] = img_width
    config_json["img_height"] = img_height
    config_json["xtrim_start"] = xtrim_start
    config_json["xtrim_span"] = xtrim_span
    config_json["ytrim_start"] = ytrim_start
    config_json["ytrim_span"] = ytrim_span
    config_json["span"] = span
    config_json["mask_right"] = mask_right
    config_json["mask_left"] = mask_left
    config_json["mask_top"] = mask_top
    config_json["mask_bottom"] = mask_bottom
    config_json["fps"] = fps

    with open(save_file, "w") as ff:
        ff.write(json.dumps(config_json, indent=2))


def exit_progrram():
    sys.exit()


def sizeup_width_right():
    global img_width, xtrim_start, xtrim_span, width, entry_span
    global mask_right

    if bln_mask_check.get():
        mask_right += min([int(entry_span.get()), xtrim_span - mask_right - 1])
    else:
        if xtrim_start + xtrim_span >= width:
            return

        img_width += min([int(entry_span.get()), width - xtrim_start - xtrim_span])
        xtrim_span += min([int(entry_span.get()), width - xtrim_start - xtrim_span])

        canvas_monitor.configure(width=img_width)
        canvas_vr.configure(width=img_width)
        canvas_vr_max.configure(width=img_width)


def sizedown_width_right():
    global img_width, xtrim_start, xtrim_span, width, entry_span
    global mask_right, mask_left

    if bln_mask_check.get():
        mask_right -= min([int(entry_span.get()), mask_right - mask_left - 1])
    else:
        if xtrim_span <= 3:
            return

        img_width -= min([int(entry_span.get()), xtrim_span - 3])
        xtrim_span -= min([int(entry_span.get()), xtrim_span - 3])

        canvas_monitor.configure(width=img_width)
        canvas_vr.configure(width=img_width)
        canvas_vr_max.configure(width=img_width)


def sizeup_width_left():
    global img_width, xtrim_start, xtrim_span, width, entry_span
    global mask_right, mask_left

    if bln_mask_check.get():
        mask_left += min([int(entry_span.get()), mask_right - mask_left - 1])
    else:
        if xtrim_start == 0:
            return

        img_width += min([int(entry_span.get()), xtrim_start])
        xtrim_start -= min([int(entry_span.get()), xtrim_start])
        xtrim_span += min([int(entry_span.get()), xtrim_start])

        canvas_monitor.configure(width=img_width)
        canvas_vr.configure(width=img_width)
        canvas_vr_max.configure(width=img_width)


def sizedown_width_left():
    global img_width, xtrim_start, xtrim_span, width, entry_span
    global mask_right, mask_left

    if bln_mask_check.get():
        mask_left -= min([int(entry_span.get()), mask_left])
    else:
        if xtrim_span <= 3:
            return

        img_width -= min([int(entry_span.get()), xtrim_span - 3])
        xtrim_start += min([int(entry_span.get()), xtrim_span - 3])
        xtrim_span -= min([int(entry_span.get()), xtrim_span - 3])

        canvas_monitor.configure(width=img_width)
        canvas_vr.configure(width=img_width)
        canvas_vr_max.configure(width=img_width)


def sizeup_height_top():
    global img_height, ytrim_start, ytrim_span, height, entry_span
    global mask_top, mask_bottom

    if bln_mask_check.get():
        mask_top += min([int(entry_span.get()), mask_bottom - mask_top - 1])
    else:
        if ytrim_start == 0:
            return

        img_height += min([int(entry_span.get()), ytrim_start])
        ytrim_start -= min([int(entry_span.get()), ytrim_start])
        ytrim_span += min([int(entry_span.get()), ytrim_start])

        canvas_monitor.configure(height=img_height)
        canvas_vr.configure(height=img_height)
        canvas_vr_max.configure(height=img_height)


def sizedown_height_top():
    global img_height, ytrim_start, ytrim_span, height, entry_span
    global mask_top, mask_bottom

    if bln_mask_check.get():
        mask_top -= min([int(entry_span.get()), mask_top])
    else:
        if ytrim_span <= 3:
            return

        img_height -= min([int(entry_span.get()), ytrim_span - 3])
        ytrim_start += min([int(entry_span.get()), ytrim_span - 3])
        ytrim_span -= min([int(entry_span.get()), ytrim_span - 3])

        canvas_monitor.configure(height=img_height)
        canvas_vr.configure(height=img_height)
        canvas_vr_max.configure(height=img_height)


def sizeup_height_bottom():
    global img_height, ytrim_start, ytrim_span, height, entry_span
    global mask_bottom

    if bln_mask_check.get():
        mask_bottom += min([int(entry_span.get()), ytrim_span - mask_bottom - 1])
    else:
        if ytrim_start + ytrim_span > height:
            return

        img_height += min([int(entry_span.get()), height - ytrim_start - ytrim_span])
        ytrim_span += min([int(entry_span.get()), height - ytrim_start - ytrim_span])

        canvas_monitor.configure(height=img_height)
        canvas_vr.configure(height=img_height)
        canvas_vr_max.configure(height=img_height)


def sizedown_height_bottom():
    global img_height, ytrim_start, ytrim_span, height, entry_span
    global mask_top, mask_bottom

    if bln_mask_check.get():
        mask_bottom -= min([int(entry_span.get()), mask_bottom - mask_top - 1])
    else:
        if ytrim_span <= 3:
            return

        img_height -= min([int(entry_span.get()), ytrim_span - 3])
        ytrim_span -= min([int(entry_span.get()), ytrim_span - 3])

        canvas_monitor.configure(height=img_height)
        canvas_vr.configure(height=img_height)
        canvas_vr_max.configure(height=img_height)


# 初期値定義
save_file = "./vrgmt_config.cfg"
config_json = {}
try:
    with open(save_file, "r") as f:
        config_json = json.loads(f.read())
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
except Exception:
    img_width = 305
    img_height = 30
    xtrim_start = 1610
    xtrim_span = 305
    ytrim_start = 285
    ytrim_span = 30
    span = 5
    mask_right = 289
    mask_left = 45
    mask_top = 9
    mask_bottom = 17
    fps = 10
    pass

# ウィンドウ作成
main = tkinter.Tk()
main.title("設定ファイル作成ツール")
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

# 図
canvas_monitor = tkinter.Canvas(frame_1_0, width=img_width, height=img_height)
canvas_vr = tkinter.Canvas(frame_1_0, width=img_width, height=img_height)
canvas_vr_max = tkinter.Canvas(frame_1_0, width=img_width, height=img_height)
canvas_monitor.grid(row=0, column=0, columnspan=5, sticky="n")
canvas_vr.grid(row=1, column=0, columnspan=5, sticky="n")
canvas_vr_max.grid(row=2, column=0, columnspan=5, sticky="n")

# VR解析結果
label_text = tkinter.StringVar()
label_vr_percent = tkinter.Label(frame_1_0, fg="green", font=label_font, textvariable=label_text)
label_vr_percent.grid(row=3, column=0, columnspan=5)

# チェックボックス用
bln_mask_check = tkinter.BooleanVar()
bln_mask_check.set(False)
bln_judge_line = tkinter.BooleanVar()
bln_judge_line.set(False)


# 描画色々
realtime_grid_image()
control_space_window()


# ループ
main.mainloop()
