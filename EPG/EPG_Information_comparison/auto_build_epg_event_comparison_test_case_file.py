#!/usr/bin/python3
# -*- coding: utf-8 -*-

epg_event_comparison_case = [
    ['00', 'GX', 'TV', 'RIGHT', 'EPGEventComparison'],
    ['01', 'GX', 'Radio', 'RIGHT', 'EPGEventComparison']
]

read_data = []

with open("epg_event_comparison_example_file.py", "r") as f:
    for line in f.readlines():
        print(line, end='')
        read_data.append(line)


def build_x_numb_python_file():
    for i in range(len(epg_event_comparison_case)):
        python_file_path = "comparison_{}_{}_{}_{}.py".format(
            epg_event_comparison_case[i][0], epg_event_comparison_case[i][2],
            epg_event_comparison_case[i][3], epg_event_comparison_case[i][4])
        with open(python_file_path, "a+") as fo:
            for j in range(len(read_data)):
                if "choice_case_numb =" in read_data[j]:
                    fo.write("choice_case_numb = {}\n".format(int(epg_event_comparison_case[i][0])))
                elif "case_info" in read_data[j]:
                    fo.write(f"case_info = {epg_event_comparison_case[i]}\n")
                else:
                    fo.write(read_data[j])


build_x_numb_python_file()
