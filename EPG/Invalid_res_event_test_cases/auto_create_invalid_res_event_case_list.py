#!/usr/bin/python3
# -*- coding: utf-8 -*-

screen_test_numb = 20
other_interface_test_numb = 50
total_list = []
n = 0
epg_res_type = ["Play", "PVR"]
timer_res_type = ["Play", "PVR", "Power Off", "Power On"]
res_mode = ["Once", "Daily", "Weekly"]
timer_res_mode = ["Once", "Daily", "Mon.", "Tues.", "Wed.", "Thurs.", "Fri.", "Sat.", "Sun."]
res_interface = ["EPG", "Timer"]
epg_res_scenes = ["Expired", "NowPlaying"]
ch_type = ["TV", "Radio"]
out_of_range = [
    "Boundary_before_lower_limit",
    "Boundary_before_upper_limit",
    "Boundary_after_lower_limit",
    "Boundary_after_upper_limit",
    "Random_before_save_time_range",
    "Random_after_save_time_range"
]
timer_expired_scenes = [
    "Boundary_lower_limit",
    "Boundary_upper_limit",
    "Random_expired_time_range"
]
time_zone = ["ZeroTimezone", "OtherTimezone"]

# EPG界面
for j in epg_res_scenes:
    for k in ch_type:
        for i in epg_res_type:
            single_list = ["", "GX", "", "", "Once", "", "EPG", "epg_test_numb"]
            single_list[0] = "{0:02d}".format(n)
            single_list[2] = k
            single_list[3] = i
            single_list[5] = j
            total_list.append(single_list)
            n += 1

# Timer Setting界面
#
# ["01", "GX", "TV", "Play", "Once", "OutOfSaveTimeRange", "Timer", "Boundary_before_upper_limit", "ZeroTimezone", "epg_test_numb"]
# for i in time_zone:
#     for j in timer_res_type:
#         for k in out_of_range:
#             # if i == "ZeroTimezone":
#             single_list = ["", "GX", "TV", "", "Once", "OutOfSaveTimeRange", "Timer", "", "", "epg_test_numb"]
#             single_list[0] = "{0:02d}".format(n)
#             single_list[3] = j
#             single_list[7] = k
#             single_list[8] = i
#             total_list.append(single_list)
#             n += 1

# out of range + play + ZeroTimezone + NoSummertime
for k in out_of_range:
    single_list = ["", "GX", "TV", "Play", "Once", "OutOfSaveTimeRange", "Timer", "",
                   "ZeroTimezone", "NoSummertime", "epg_test_numb"]
    single_list[0] = "{0:02d}".format(n)
    single_list[7] = k
    total_list.append(single_list)
    n += 1

# out of range + PVR + ZeroTimezone + Summertime
for k in out_of_range:
    single_list = ["", "GX", "TV", "PVR", "Once", "OutOfSaveTimeRange", "Timer", "",
                   "ZeroTimezone", "Summertime", "epg_test_numb"]
    single_list[0] = "{0:02d}".format(n)
    single_list[7] = k
    total_list.append(single_list)
    n += 1

# out of range + Power Off + OtherTimezone + NoSummertime
for k in out_of_range:
    single_list = ["", "GX", "TV", "Power Off", "Once", "OutOfSaveTimeRange", "Timer", "",
                   "OtherTimezone", "NoSummertime", "epg_test_numb"]
    single_list[0] = "{0:02d}".format(n)
    single_list[7] = k
    total_list.append(single_list)
    n += 1
# out of range + Power On + OtherTimezone + Summertime
for k in out_of_range:
    single_list = ["", "GX", "TV", "Power On", "Once", "OutOfSaveTimeRange", "Timer", "",
                   "OtherTimezone", "Summertime", "epg_test_numb"]
    single_list[0] = "{0:02d}".format(n)
    single_list[7] = k
    total_list.append(single_list)
    n += 1


# Timer setting Expired用例
# for i in timer_res_type:
#     for j in timer_expired_scenes:
#         single_list = ["", "GX", "TV", "", "Once", "Expired", "Timer", "", "ZeroTimezone", "epg_test_numb"]
#         single_list[0] = "{0:02d}".format(n)
#         single_list[3] = i
#         single_list[7] = j
#         total_list.append(single_list)
#         n += 1

# Expired + play + ZeroTimezone + NoSummertime
for k in timer_expired_scenes:
    single_list = ["", "GX", "TV", "Play", "Once", "Expired", "Timer", "",
                   "ZeroTimezone", "NoSummertime", "epg_test_numb"]
    single_list[0] = "{0:02d}".format(n)
    single_list[7] = k
    total_list.append(single_list)
    n += 1

# Expired + PVR + ZeroTimezone + Summertime
for k in timer_expired_scenes:
    single_list = ["", "GX", "TV", "PVR", "Once", "Expired", "Timer", "",
                   "ZeroTimezone", "Summertime", "epg_test_numb"]
    single_list[0] = "{0:02d}".format(n)
    single_list[7] = k
    total_list.append(single_list)
    n += 1

# Expired + Power Off + OtherTimezone + NoSummertime
for k in timer_expired_scenes:
    single_list = ["", "GX", "TV", "Power Off", "Once", "Expired", "Timer", "",
                   "OtherTimezone", "NoSummertime", "epg_test_numb"]
    single_list[0] = "{0:02d}".format(n)
    single_list[7] = k
    total_list.append(single_list)
    n += 1
# Expired + Power On + OtherTimezone + Summertime
for k in timer_expired_scenes:
    single_list = ["", "GX", "TV", "Power On", "Once", "Expired", "Timer", "",
                   "OtherTimezone", "Summertime", "epg_test_numb"]
    single_list[0] = "{0:02d}".format(n)
    single_list[7] = k
    total_list.append(single_list)
    n += 1

# Timer setting NowPlaying用例
for i in timer_res_type:
    single_list = ["", "GX", "TV", "", "Once", "NowPlaying", "Timer", "None",
                   "ZeroTimezone", "NoSummertime", "epg_test_numb"]
    single_list[0] = "{0:02d}".format(n)
    single_list[3] = i
    total_list.append(single_list)
    n += 1

# Timer setting InvalidDuration用例
for i in timer_res_type:
    for j in timer_res_mode:
        if i == "PVR":
            single_list = ["", "GX", "TV", "", "", "InvalidDuration", "Timer", "None",
                           "ZeroTimezone", "NoSummertime", "epg_test_numb"]
            single_list[0] = "{0:02d}".format(n)
            single_list[3] = i
            single_list[4] = j
            total_list.append(single_list)
            n += 1

# print(len(total_list))
for m in range(len(total_list)):
    print(f"{total_list[m]},")
