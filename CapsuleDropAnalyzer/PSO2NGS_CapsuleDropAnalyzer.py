# -*- coding: utf-8 -*-
import datetime
import math
import tkinter
import tkinter.scrolledtext
import tkinter.filedialog
import tkinter.ttk
import re


global main, log_data, log_file_path, realtime_parse_flag, stop_realtime_parse_flag, job, output_text, special_items
global display_control_space_button, hide_control_space_button, checkbutton,\
    realtime_parse_logfile_button, stop_realtime_parse_logfile_button
global start_year, start_month, start_day, start_hour, start_minute, play_minutes, capsule_name_regexp

log_data = None
log_file_path = None
realtime_parse_flag = False
stop_realtime_parse_flag = False
job = None
capsule_name_regexp = ""

special_items = ["ストラグメント", "リファイナー", "Bトリガー"]


class UnknownFile(Exception):
    pass


def read_logfile(reset_flag=False):
    global log_data, log_file_path, output_text
    try:
        if log_file_path is None or reset_flag:
            log_file_path = tkinter.filedialog.askopenfilenames(title="ファイル選択", filetypes=[("", "*.txt")])
        if log_file_path == "":
            log_file_path = None
            return
        if "ActionLog" not in log_file_path[0]:
            raise UnknownFile(f"{log_file_path[0]} is not an ActionLog file.")
        with open(log_file_path[0], encoding="utf16") as f:
            del log_data
            log_data = f.read()
    except Exception as e:
        output_text.delete('0.0', tkinter.END)
        output_text["fg"] = "red"
        output_text.insert(tkinter.END, e)
        log_data = None


def realtime_parse_logfile():
    global main, realtime_parse_flag, job, realtime_parse_logfile_button, stop_realtime_parse_logfile_button
    if realtime_parse_flag is False:
        realtime_parse_logfile_button["state"] = tkinter.DISABLED
        stop_realtime_parse_logfile_button["state"] = tkinter.ACTIVE
    realtime_parse_flag = True
    # print("realtime_parse_flag : ", realtime_parse_flag)
    parse_logfile()
    job = main.after(1000, realtime_parse_logfile)


def stop_realtime_parse_logfile():
    global main, realtime_parse_flag, job, realtime_parse_logfile_button, stop_realtime_parse_logfile_button
    # print("stopping...")
    main.after_cancel(job)
    realtime_parse_flag = False
    realtime_parse_logfile_button["state"] = tkinter.ACTIVE
    stop_realtime_parse_logfile_button["state"] = tkinter.DISABLED


def parse_logfile():
    global log_data, realtime_parse_flag, frame_0_1, play_minutes, output_text, capsule_name_regexp, special_items
    try:
        if log_data is None:
            return
        read_logfile()
        datetime_format = "%Y-%m-%dT%H:%M:%S"

        datetime_start = datetime.datetime.strptime(
            str(start_year.get()).zfill(4) + "-"
            + str(start_month.get()).zfill(2) + "-"
            + str(start_day.get()).zfill(2) + "T"
            + str(start_hour.get()).zfill(2) + ":"
            + str(start_minute.get()).zfill(2) + ":"
            + "00",
            datetime_format
        )

        if realtime_parse_flag:
            datetime_end = datetime.datetime.now()
        else:
            datetime_end = datetime_start + datetime.timedelta(minutes=int(play_minutes.get()))
        # print(datetime_start, datetime_end, realtime_parse_flag)

        pickup_log_format = ["datetime", "id", "actiontype", "playerid", "playername", "itemname", "additionalinfo"]

        pickup_actiontype_name = "[Pickup]"

        log_data = [i.split("\t") for i in log_data.split("\n") if i != ""]

        # ドロップログ抽出
        pickup_log_data = [dict(zip(pickup_log_format, i)) for i in log_data if i[2] == pickup_actiontype_name]

        # 指定時刻データ抽出
        pickup_log_data = [
            i for i in pickup_log_data
            if datetime_end > datetime.datetime.strptime(i["datetime"], datetime_format) > datetime_start
        ]

        # 各カテゴリ抽出
        pickup_item_list = list(set([i["itemname"] for i in pickup_log_data if i["itemname"] != ""]))
        pickup_capsule_list = sorted([i for i in pickup_item_list if i[0] == "C"])
        pickup_capsule_num_list = [len([j for j in pickup_log_data if j["itemname"] == i]) for i in pickup_capsule_list]

        pickup_capsule_result = dict(zip(pickup_capsule_list, pickup_capsule_num_list))

        pickup_meseta_list = [i for i in pickup_log_data if i["itemname"] == "" and "N-Meseta" in i["additionalinfo"]]

        pickup_teims_list = [i for i in pickup_log_data if "肉" in i["itemname"]]

        pickup_special_items_list = [i["itemname"] for i in pickup_log_data if sum([j in i["itemname"] for j in special_items]) != 0]

        # 出力
        datetime_text = datetime_start.strftime(datetime_format) + " ～ " + datetime_end.strftime(datetime_format)
        if capsule_name_regexp.get() != "":
            pickup_text = "\n\nピックアップカプセル : \n" + "\n".join(
                [
                    i + " : " + str(pickup_capsule_result[i])
                    for i in pickup_capsule_result.keys()
                    if re.search(capsule_name_regexp.get(), i) is not None
                ]
            )
        else:
            pickup_text = ""

        if pickup_special_items_list:
            pickup_special_items_unique = sorted(list(set(pickup_special_items_list)))
            pickup_special_items_list = dict(zip(pickup_special_items_unique,
                                             [pickup_special_items_list.count(i) for i in pickup_special_items_unique]))
            pickup_si_text = "\n\nピックアップアイテム : \n" + "\n".join(
                [
                    i + " : " + str(pickup_special_items_list[i])
                    for i in pickup_special_items_list.keys()
                ]
            )
        else:
            pickup_si_text = ""

        enemy_text = "全討伐数概算 : " + str(len(pickup_meseta_list) + len(pickup_teims_list)) + "\n" \
                     + "(テイムズ : " + str(len(pickup_teims_list)) + ", その他 : " + str(len(pickup_meseta_list)) + ")"
        capsule_text = "\n\nドロップカプセル一覧 :\n" \
                       + "\n".join([i + " : " + str(pickup_capsule_result[i]) for i in pickup_capsule_result.keys()])

        output_text.delete('0.0', tkinter.END)
        output_text["fg"] = "white"
        output_text.insert(tkinter.END, datetime_text + "\n" + enemy_text + pickup_text + pickup_si_text + capsule_text)
    except Exception as e:
        output_text.delete('0.0', tkinter.END)
        output_text["fg"] = "red"
        output_text.insert(tkinter.END, e)


def set_datetime_now():
    global start_year, start_month, start_day, start_hour, start_minute
    try:
        now = datetime.datetime.now()

        start_year.delete(0, tkinter.END)
        start_month.delete(0, tkinter.END)
        start_day.delete(0, tkinter.END)
        start_hour.delete(0, tkinter.END)
        start_minute.delete(0, tkinter.END)

        start_year.insert(tkinter.END, now.year)
        start_month.insert(tkinter.END, now.month)
        start_day.insert(tkinter.END, now.day)
        start_hour.insert(tkinter.END, now.hour)
        start_minute.insert(tkinter.END, now.minute)
    except Exception as e:
        output_text.delete('0.0', tkinter.END)
        output_text["fg"] = "red"
        output_text.insert(tkinter.END, e)


def hide_control_space():
    global display_control_space_button

    for name, widget in frame_0_0.children.items():
        widget.grid_remove()
    display_control_space_button.grid()
    main.geometry(f"{max_width - x_trim}x{main.winfo_height()}+{main.winfo_x() + x_trim}+{main.winfo_y()}")


def display_control_space():
    global display_control_space_button, hide_control_space_button

    for name, widget in frame_0_0.children.items():
        widget.grid()
    display_control_space_button.grid_remove()
    hide_control_space_button.grid()
    main.geometry(f"{max_width}x{main.winfo_height()}+{main.winfo_x() - x_trim}+{main.winfo_y()}")


def create_control_space():
    global display_control_space_button, hide_control_space_button, checkbutton,\
        realtime_parse_logfile_button, stop_realtime_parse_logfile_button, output_text
    global start_year, start_month, start_day, start_hour, start_minute, play_minutes, capsule_name_regexp
    # ボタン作成・配置
    hide_control_space_button = tkinter.ttk.Button(frame_0_0, text=">>", width=4, command=hide_control_space)
    hide_control_space_button.grid(row=0, column=0, columnspan=12, sticky="e")

    display_control_space_button = tkinter.ttk.Button(frame_0_0, text="<<", width=4, command=display_control_space)
    display_control_space_button.grid(row=0, column=0, columnspan=12, sticky="e")
    display_control_space_button.grid_remove()

    read_logfile_button = tkinter.ttk.Button(frame_0_0, text="ファイル選択", width=13, default="active",
                                             command=lambda: read_logfile(reset_flag=True)
                                             )
    read_logfile_button.grid(row=1, column=0, columnspan=12, pady=20, sticky="n")

    start_datetime_lbl = tkinter.Label(frame_0_0, text='開始 :', font=("Meiryo", 9))
    start_datetime_lbl.grid(row=2, column=0, padx=10, sticky="nw")

    datetime_row = 2
    datetime_column = 2

    start_year = tkinter.Entry(frame_0_0, width=5, font=("Meiryo", 9))
    start_year.grid(row=datetime_row, column=datetime_column, sticky="nw")
    lbl = tkinter.Label(frame_0_0, text='年', font=("Meiryo", 9))
    lbl.grid(row=datetime_row, column=datetime_column + 1, sticky="nw")

    start_month = tkinter.Entry(frame_0_0, width=3, font=("Meiryo", 9))
    start_month.grid(row=datetime_row, column=datetime_column + 2, sticky="nw")
    lbl = tkinter.Label(frame_0_0, text='月', font=("Meiryo", 9))
    lbl.grid(row=datetime_row, column=datetime_column + 3, sticky="nw")

    start_day = tkinter.Entry(frame_0_0, width=3, font=("Meiryo", 9))
    start_day.grid(row=datetime_row, column=datetime_column + 4, sticky="nw")
    lbl = tkinter.Label(frame_0_0, text='日', font=("Meiryo", 9))
    lbl.grid(row=datetime_row, column=datetime_column + 5, sticky="nw")

    start_hour = tkinter.Entry(frame_0_0, width=3, font=("Meiryo", 9))
    start_hour.grid(row=datetime_row, column=datetime_column + 6, sticky="nw")
    lbl = tkinter.Label(frame_0_0, text='時', font=("Meiryo", 9))
    lbl.grid(row=datetime_row, column=datetime_column + 7, sticky="nw")

    start_minute = tkinter.Entry(frame_0_0, width=3, font=("Meiryo", 9))
    start_minute.grid(row=datetime_row, column=datetime_column + 8, sticky="nw")
    lbl = tkinter.Label(frame_0_0, text='分   ', font=("Meiryo", 9))
    lbl.grid(row=datetime_row, column=datetime_column + 9, sticky="nw")

    set_datetime_now()

    lbl = tkinter.Label(frame_0_0, text='マルグル時間 : ', font=("Meiryo", 9))
    lbl.grid(row=3, column=0, columnspan=2, padx=10, sticky="nw")

    play_minutes = tkinter.Entry(frame_0_0, width=5, font=("Meiryo", 9))
    play_minutes.grid(row=3, column=2, sticky="nw")
    play_minutes.insert(tkinter.END, 120)
    lbl = tkinter.Label(frame_0_0, text='分', font=("Meiryo", 9))
    lbl.grid(row=3, column=3, sticky="nw")

    parse_logfile_button = tkinter.ttk.Button(frame_0_0, text="結果解析", width=13, default="active", command=parse_logfile)
    parse_logfile_button.grid(row=4, column=0, pady=10, columnspan=12, sticky="n")

    realtime_parse_logfile_button = tkinter.ttk.Button(
        frame_0_0,
        text="リアルタイム解析", width=15, default="active",
        command=realtime_parse_logfile
    )
    realtime_parse_logfile_button.grid(row=5, column=0, columnspan=6, sticky="ne")

    stop_realtime_parse_logfile_button = tkinter.ttk.Button(
        frame_0_0,
        text="停止", width=7, state="disabled",
        command=stop_realtime_parse_logfile)
    stop_realtime_parse_logfile_button.grid(row=5, column=6, columnspan=6, sticky="nw")

    input_datetime_now_button = tkinter.ttk.Button(frame_0_0, text="now", width=5, command=set_datetime_now)
    input_datetime_now_button.grid(row=datetime_row, column=1, sticky="nw")

    lbl = tkinter.Label(frame_0_0, text='ピックアップ : ', font=("Meiryo", 9))
    lbl.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="nw")

    capsule_name_regexp = tkinter.Entry(frame_0_0, width=20, font=("Meiryo", 9))
    capsule_name_regexp.grid(row=6, column=2, pady=10, columnspan=5, sticky="nw")

    checkbutton_topmost = tkinter.Checkbutton(frame_0_0, variable=bln_topmost, text="常時最前面表示")
    checkbutton_topmost.grid(row=7, column=0, padx=10, pady=1, columnspan=6, sticky="nw")

    scale_alpha = tkinter.Scale(
        frame_0_0,
        variable=alpha_val, from_=0, to=70, orient=tkinter.HORIZONTAL,
        command=lambda e: set_alpha(alpha_rate=alpha_val.get())
    )
    scale_alpha.grid(row=8, column=0, padx=10, columnspan=3, sticky="w")
    lbl = tkinter.Label(frame_0_0, text='ウィンドウ透過率[%]', font=("Meiryo", 8))
    lbl.grid(row=8, column=2, padx=5, columnspan=4, sticky="ws")

    output_text = tkinter.scrolledtext.ScrolledText(
        frame_0_1,
        width=60,
        height=40,
        fg="white",
        bg="black",
        font=("Meiryo", 10)
    )
    output_text.grid(row=0, column=0, sticky="nw")


def set_alpha(alpha_rate):
    main.attributes("-alpha", 1 - alpha_rate/100)


def set_topmost():
    if bln_topmost.get():
        main.attributes("-topmost", True)
    else:
        main.attributes("-topmost", False)


def change_output_text_height(event):
    global output_text
    output_text["height"] = math.floor(max([frame_0_0.winfo_height(), main.winfo_height()]) / 20)


def buttonrelease_function(event):
    set_topmost()


# 初期値設定&ウィンドウ作成
max_width = 880
max_height = 810
main = tkinter.Tk()
main.title("カプセルドロップ解析ツール")
main.attributes("-topmost", False, "-alpha", 1)
main.geometry(f"{max_width}x{max_height}")
number_of_capsule = 50
x_trim = 343
bln_topmost = tkinter.BooleanVar()
bln_topmost.set(False)
alpha_val = tkinter.DoubleVar()

s = tkinter.ttk.Style()
s.configure('MyWidget.TButton', font=5, background='#ffffcc')

# フレーム定義
frame_0_0 = tkinter.ttk.Frame(main)
frame_0_1 = tkinter.ttk.Frame(main)
frame_0_0.grid(row=0, column=0, rowspan=2)
frame_0_1.grid(row=0, column=1, rowspan=number_of_capsule)

# 境界線
separator_0_1 = tkinter.ttk.Separator(main, orient="vertical")
separator_0_1.grid(row=0, column=0, rowspan=number_of_capsule, sticky="nes")

# ボタン配置
create_control_space()

# イベント設定
main.bind("<ButtonRelease>", buttonrelease_function)
main.bind("<Configure>", change_output_text_height)
# print(tkinter.EventType.__members__)

# イベントループ
main.mainloop()
