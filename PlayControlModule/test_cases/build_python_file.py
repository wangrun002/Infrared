#!/usr/bin/python
# -*- coding: utf-8 -*-

INTERVAL_TIME = [1.0, 5.0]

ALL_TEST_CASE = [
        ["numb_key", "one_by_one", "timeout", "All", "TV", INTERVAL_TIME[1]],
        ["numb_key", "one_by_one", "timeout", "All", "Radio", INTERVAL_TIME[1]],
        ["numb_key", "one_by_one", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["numb_key", "random", "timeout", "All", "TV", INTERVAL_TIME[1]],
        ["screen", "up", "timeout", "All", "TV", INTERVAL_TIME[1]],
        ["screen", "down", "timeout", "All", "TV", INTERVAL_TIME[1]],
        ["screen", "random", "timeout", "All", "TV", INTERVAL_TIME[1]],
        ["screen", "up", "continuous", "All", "TV", INTERVAL_TIME[0]],
        ["ch_list", "up", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_list", "down", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_list", "random_up", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_list", "random_down", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_list", "page_up", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_list", "page_down", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_list", "left", "group", "All", "TV", INTERVAL_TIME[1]],
        ["ch_list", "right", "group", "All", "TV", INTERVAL_TIME[1]],
        ["ch_list", "left_group_random_down", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_list", "right_group_random_down", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["epg", "up", "timeout", "All", "TV", INTERVAL_TIME[1]],
        ["epg", "down", "timeout", "All", "TV", INTERVAL_TIME[1]],
        ["epg", "page_up", "timeout", "All", "TV", INTERVAL_TIME[1]],
        ["epg", "page_down", "timeout", "All", "TV", INTERVAL_TIME[1]],
        ["epg", "random", "timeout", "All", "TV", INTERVAL_TIME[1]],
        ["epg", "up", "continuous", "All", "TV", INTERVAL_TIME[0]],
        ["ch_edit", "up", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_edit", "down", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_edit", "random_up", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_edit", "random_down", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_edit", "page_up", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_edit", "page_down", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_edit", "left", "group", "All", "TV", INTERVAL_TIME[1]],
        ["ch_edit", "right", "group", "All", "TV", INTERVAL_TIME[1]],
        ["ch_edit", "left_group_random_down", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["ch_edit", "right_group_random_down", "add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["tv_radio", "free_tv", "free_radio", "All", "TV", INTERVAL_TIME[1]],
        ["tv_radio", "free_tv", "scr_radio", "All", "TV", INTERVAL_TIME[1]],
        ["tv_radio", "free_tv", "lock_radio", "All", "TV", INTERVAL_TIME[1]],
        ["tv_radio", "scr_tv", "free_radio", "All", "TV", INTERVAL_TIME[1]],
        ["tv_radio", "scr_tv", "scr_radio", "All", "TV", INTERVAL_TIME[1]],
        ["tv_radio", "scr_tv", "lock_radio", "All", "TV", INTERVAL_TIME[1]],
        ["tv_radio", "lock_tv", "free_radio", "All", "TV", INTERVAL_TIME[1]],
        ["tv_radio", "lock_tv", "scr_radio", "All", "TV", INTERVAL_TIME[1]],
        ["tv_radio", "lock_tv", "lock_radio", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "free_tv", "free_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "free_tv", "scr_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "free_tv", "lock_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "scr_tv", "scr_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "scr_tv", "lock_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "lock_tv", "lock_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "free_radio", "free_radio", "All", "Radio", INTERVAL_TIME[1]],
        ["recall", "free_radio", "scr_radio", "All", "Radio", INTERVAL_TIME[1]],
        ["recall", "free_radio", "lock_radio", "All", "Radio", INTERVAL_TIME[1]],
        ["recall", "scr_radio", "scr_radio", "All", "Radio", INTERVAL_TIME[1]],
        ["recall", "scr_radio", "lock_radio", "All", "Radio", INTERVAL_TIME[1]],
        ["recall", "lock_radio", "lock_radio", "All", "Radio", INTERVAL_TIME[1]],
        ["recall", "free_tv", "free_radio", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "free_tv", "scr_radio", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "free_tv", "lock_radio", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "scr_tv", "free_radio", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "scr_tv", "scr_radio", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "scr_tv", "lock_radio", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "lock_tv", "free_radio", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "lock_tv", "scr_radio", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "lock_tv", "lock_radio", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "same_tp_tv", "same_tp_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "diff_tp_tv", "diff_tp_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "same_codec_tv", "same_codec_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "diff_codec_tv", "diff_codec_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "hd_tv", "hd_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "sd_tv", "sd_tv", "All", "TV", INTERVAL_TIME[1]],
        ["recall", "hd_sd_tv", "hd_sd_tv", "All", "TV", INTERVAL_TIME[1]],
    ]

read_data = []

with open("new_file.py","r") as f:
    for line in f.readlines():
        print(line,end='')
        read_data.append(line)

# with open("new_file.py","a+") as fo:
#     for i in range(len(read_data)):
#         fo.write(read_data[i])



def build_x_numb_python_file():
    for i in range(len(ALL_TEST_CASE)):
        if i < 10:
            sequence_numb = "0{}".format(i)
        else:
            sequence_numb = "{}".format(i)
        python_file_path = "test_{}_{}_{}_{}_{}_switch_channel.py".format(sequence_numb,ALL_TEST_CASE[i][0],ALL_TEST_CASE[i][1],
                                                                          ALL_TEST_CASE[i][2],ALL_TEST_CASE[i][4])
        with open(python_file_path,"a+") as fo:
            for j in range(len(read_data)):
                if "choice_switch_mode =" in read_data[j]:
                    fo.write("choice_switch_mode = {}\n".format(i))
                else:
                    fo.write(read_data[j])

build_x_numb_python_file()