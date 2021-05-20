import xlrd
from Vitualkeystrokeexample import *
from win32gui import *
from Parse_Music import *


def isFF14ForeGround():
    try:
        fg_window_name = GetWindowText(GetForegroundWindow()).lower()
        FF14_window_name = "最终幻想XIV".lower()
        if fg_window_name == FF14_window_name:
            return True
    except Exception as e:
        print(e)
        return False


def press_keyboard_vskey(name):
    while not isFF14ForeGround():
        time.sleep(2)
    parse_list(name)
    print(GetWindowText(GetForegroundWindow()).lower())


def counting_tempo(tempo):
    tempo = int(tempo)
    list = []
    base_time = tempo / 60
    quarter_time = 0.25 / base_time  # 0
    list.append(quarter_time)
    half_time = 0.5 / base_time  # 1
    list.append(half_time)
    half_quarter_time = 0.75 / base_time  # 2
    list.append(half_quarter_time)
    one_time = 1 / base_time  # 3
    list.append(one_time)
    one_half_time = 1.5 / base_time  # 4
    list.append(one_half_time)
    one_quarter_time = 1.25 / base_time  # 5
    list.append(one_quarter_time)
    one_half_quarter_time = 1.75 / base_time  # 6
    list.append(one_half_quarter_time)
    two_time = 2 / base_time  # 7
    list.append(two_time)
    two_half_time = 2.5 / base_time  # 8
    list.append(two_half_time)
    three_time = 3 / base_time  # 9
    list.append(three_time)
    four_time = 4 / base_time  # 10
    list.append(four_time)
    print(list)
    return list


def transform_tempo(ttime):
    dict = {0.25: 0, 0.5: 1, 0.75: 2, 1.0: 3, 1.5: 4, 1.25: 5, 1.75: 6, 2.0: 7, 2.5: 8, 3.0: 9, 4.0: 10}
    index = dict[ttime]
    #print(index)
    return index

def read_data(name):
    xlsx = xlrd.open_workbook(name)
    index = xlsx.sheets()[0]
    list = []
    for i in range(index.nrows):
        list.append(index.row_values(i))
    return list


def parse_list(name, tempo):
    tempo_list = counting_tempo(tempo)

    list = read_data(name)

    for i in list:
        cmd1 = int(i[0])
        cmd1 = str(cmd1)
        tempo_list[transform_tempo(i[2])]
        print(cmd1)

        #print(cmd1, cmd2, tempo_list[transform_tempo(i[2])])
        while not isFF14ForeGround():
            time.sleep(2)
        if i[1] == 'Normal':
            pressAndHold(cmd1)
            time.sleep(tempo_list[transform_tempo(i[2])])
            release(cmd1)
        elif i[1] == 'High octave':
            pressAndHold('shift', cmd1)
            time.sleep(tempo_list[transform_tempo(i[2])])
            release('shift', cmd1)
        elif i[1] == 'Lower octave':
            pressAndHold('ctrl', cmd1)
            time.sleep(tempo_list[transform_tempo(i[2])])
            release('ctrl', cmd1)
        elif i[1] == 'High semitone':
            pressAndHold('[', cmd1)
            time.sleep(tempo_list[transform_tempo(i[2])])
            release('[', cmd1)
        elif i[1] == 'Lower semitone':
            pressAndHold(']', cmd1)
            time.sleep(tempo_list[transform_tempo(i[2])])
            release(']', cmd1)
            # release(cmd1)

