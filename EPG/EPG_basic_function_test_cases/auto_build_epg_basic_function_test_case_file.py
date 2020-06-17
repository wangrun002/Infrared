#!/usr/bin/python3
# -*- coding: utf-8 -*-

epg_basic_case = [
    ["00", "GX", "TV", "EPG+Exit", "EnterExitEPGInterface"],
    ["01", "GX", "Radio", "EPG+Exit", "EnterExitEPGInterface"],
    ["02", "GX", "TV", "UP", "EPGSwitchChannel"],
    ["03", "GX", "TV", "DOWN", "EPGSwitchChannel"],
    ["04", "GX", "TV", "UP+Random", "EPGSwitchChannel"],
    ["05", "GX", "TV", "DOWN+Random", "EPGSwitchChannel"],
    ["06", "GX", "TV", "UP+DOWN+Random", "EPGSwitchChannel"],
    ["07", "GX", "Radio", "UP", "EPGSwitchChannel"],
    ["08", "GX", "Radio", "DOWN", "EPGSwitchChannel"],
    ["09", "GX", "Radio", "UP+Random", "EPGSwitchChannel"],
    ["10", "GX", "Radio", "DOWN+Random", "EPGSwitchChannel"],
    ["11", "GX", "Radio", "UP+DOWN+Random", "EPGSwitchChannel"],
    ["12", "GX", "TV", "LEFT", "SwitchEPGEvent"],
    ["13", "GX", "TV", "RIGHT", "SwitchEPGEvent"],
    ["14", "GX", "TV", "LEFT+Random", "SwitchEPGEvent"],
    ["15", "GX", "TV", "RIGHT+Random", "SwitchEPGEvent"],
    ["16", "GX", "TV", "RIGHT+LEFT+Random", "SwitchEPGEvent"],
    ["17", "GX", "TV", "Day+", "SwitchEPGEvent"],
    ["18", "GX", "TV", "Day-", "SwitchEPGEvent"],
    ["19", "GX", "TV", "Day++RIGHT+Random", "SwitchEPGEvent"],
    ["20", "GX", "TV", "Day++LEFT+Random", "SwitchEPGEvent"],
    ["21", "GX", "TV", "Day++LEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["22", "GX", "TV", "Day-+RIGHT+Random", "SwitchEPGEvent"],
    ["23", "GX", "TV", "Day-+LEFT+Random", "SwitchEPGEvent"],
    ["24", "GX", "TV", "Day-+LEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["25", "GX", "TV", "Day+orDay-orLEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["26", "GX", "Radio", "LEFT", "SwitchEPGEvent"],
    ["27", "GX", "Radio", "RIGHT", "SwitchEPGEvent"],
    ["28", "GX", "Radio", "LEFT+Random", "SwitchEPGEvent"],
    ["29", "GX", "Radio", "RIGHT+Random", "SwitchEPGEvent"],
    ["30", "GX", "Radio", "RIGHT+LEFT+Random", "SwitchEPGEvent"],
    ["31", "GX", "Radio", "Day+", "SwitchEPGEvent"],
    ["32", "GX", "Radio", "Day-", "SwitchEPGEvent"],
    ["33", "GX", "Radio", "Day++RIGHT+Random", "SwitchEPGEvent"],
    ["34", "GX", "Radio", "Day++LEFT+Random", "SwitchEPGEvent"],
    ["35", "GX", "Radio", "Day++LEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["36", "GX", "Radio", "Day-+RIGHT+Random", "SwitchEPGEvent"],
    ["37", "GX", "Radio", "Day-+LEFT+Random", "SwitchEPGEvent"],
    ["38", "GX", "Radio", "Day-+LEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["39", "GX", "Radio", "Day+orDay-orLEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["40", "GX", "TV", "LEFT+OK", "SwitchAndDetailEPGEvent"],
    ["41", "GX", "TV", "RIGHT+OK", "SwitchAndDetailEPGEvent"],
    ["42", "GX", "TV", "LEFT+Random+OK", "SwitchAndDetailEPGEvent"],
    ["43", "GX", "TV", "RIGHT+Random+OK", "SwitchAndDetailEPGEvent"],
    ["44", "GX", "Radio", "LEFT+OK", "SwitchAndDetailEPGEvent"],
    ["45", "GX", "Radio", "RIGHT+OK", "SwitchAndDetailEPGEvent"],
    ["46", "GX", "Radio", "LEFT+Random+OK", "SwitchAndDetailEPGEvent"],
    ["47", "GX", "Radio", "RIGHT+Random+OK", "SwitchAndDetailEPGEvent"]
]

read_data = []

with open("epg_basic_function_example_file.py", "r") as f:
    for line in f.readlines():
        print(line, end='')
        read_data.append(line)


def build_x_numb_python_file():
    for i in range(len(epg_basic_case)):
        # if all_test_case[i][4] == "Power Off":      # 这里要注意py文件的名称不能有空格，会导致python3 xxx.py执行错误
        #     python_file_path = "test_{}_{}_{}_{}_{}_{}.py".format(
        #         all_test_case[i][0], all_test_case[i][2],
        #         all_test_case[i][3], all_test_case[i][4].replace(" ", ''),
        #         all_test_case[i][5], all_test_case[i][6])
        # else:
        python_file_path = "case_{}_{}_{}_{}.py".format(epg_basic_case[i][0], epg_basic_case[i][2],
                                                        epg_basic_case[i][3], epg_basic_case[i][4])
        with open(python_file_path, "a+") as fo:
            for j in range(len(read_data)):
                if "choice_case_numb =" in read_data[j]:
                    fo.write("choice_case_numb = {}\n".format(int(epg_basic_case[i][0])))
                else:
                    fo.write(read_data[j])


build_x_numb_python_file()
