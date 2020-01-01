#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import shutil
import sys

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

choice_switch_mode = 1
case_loop_time = 1

parent_path = os.path.dirname(os.getcwd())
main_file_path = os.path.join(parent_path,"main_program","main.py")
test_file_path = os.path.join(os.getcwd(),"main.py")

for i in range(case_loop_time):
    shutil.copy(main_file_path, os.getcwd())
    os.system("python %s %d" % (test_file_path, choice_switch_mode))
    os.unlink(test_file_path)