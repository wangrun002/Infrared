#!/usr/bin/python3
# -*- coding: utf-8 -*-

screen_test_numb = 20
other_interface_test_numb = 50
total_list = []
n = 0
res_type = ["Play", "PVR", "Power Off"]
res_mode = ["Once", "Daily", "Mon.", "Tues.", "Wed.", "Thurs.", "Fri.", "Sat.", "Sun."]
jump_mode = ["Manual_jump", "Auto_jump", "Cancel_jump"]
# ModifyTime
for i in res_type:
    for j in res_mode:
        single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                       "ModifyTime", "", "", "screen_test_numb"]
        single_list[0] = "{0:02d}".format(n)
        single_list[3] = j
        single_list[4] = i
        single_list[8] = j
        single_list[9] = i
        total_list.append(single_list)
        n += 1
# ModifyType
for x in res_mode:
    for y in res_type:
        for z in res_type:
            if z != y:
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "ModifyType", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = x
                single_list[4] = y
                single_list[8] = x
                single_list[9] = z
                total_list.append(single_list)
                n += 1
# ModifyDuration
for x in res_mode:
    for y in res_type:
        if y == "PVR":
            single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                           "ModifyDuration", "", "", "screen_test_numb"]
            single_list[0] = "{0:02d}".format(n)
            single_list[3] = x
            single_list[4] = y
            single_list[8] = x
            single_list[9] = y
            total_list.append(single_list)
            n += 1
# ModifyMode
for x in res_type:
    for y in res_mode:
        for z in res_mode:
            if z != y:
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "ModifyMode", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = y
                single_list[4] = x
                single_list[8] = z
                single_list[9] = x
                total_list.append(single_list)
                n += 1

# print(len(total_list))
for m in range(len(total_list)):
    print(f"{total_list[m]},")
