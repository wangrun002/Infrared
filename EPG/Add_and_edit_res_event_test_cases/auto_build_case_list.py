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
# ModifyTime+ModifyType
for x in res_mode:
    for y in res_type:
        for z in res_type:
            if z != y:
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "ModifyTime+ModifyType", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = x
                single_list[4] = y
                single_list[8] = x
                single_list[9] = z
                total_list.append(single_list)
                n += 1
# ModifyTime+ModifyDuration
for x in res_mode:
    for y in res_type:
        if y == "PVR":
            single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                           "ModifyTime+ModifyDuration", "", "", "screen_test_numb"]
            single_list[0] = "{0:02d}".format(n)
            single_list[3] = x
            single_list[4] = y
            single_list[8] = x
            single_list[9] = y
            total_list.append(single_list)
            n += 1
# ModifyTime+ModifyMode
for x in res_type:
    for y in res_mode:
        for z in res_mode:
            if z != y:
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "ModifyTime+ModifyMode", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = y
                single_list[4] = x
                single_list[8] = z
                single_list[9] = x
                total_list.append(single_list)
                n += 1
# ModifyType+ModifyDuration
for j in res_type:
    for k in res_type:
        for h in res_mode:
            if j != "PVR" and k == "PVR":
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "ModifyType+ModifyDuration", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = j
                single_list[4] = h
                single_list[8] = k
                single_list[9] = h
                total_list.append(single_list)
                n += 1
# ModifyType+ModifyMode
for j in res_type:
    for k in res_type:
        for h in res_mode:
            for i in res_mode:
                if k != j and i != h:
                    single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                                   "ModifyType+ModifyMode", "", "", "screen_test_numb"]
                    single_list[0] = "{0:02d}".format(n)
                    single_list[3] = j
                    single_list[4] = h
                    single_list[8] = k
                    single_list[9] = i
                    total_list.append(single_list)
                    n += 1
# ModifyDuration+ModifyMode
for x in res_type:
    for y in res_mode:
        for z in res_mode:
            if x == "PVR" and z != y:
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "ModifyDuration+ModifyMode", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = y
                single_list[4] = x
                single_list[8] = z
                single_list[9] = x
                total_list.append(single_list)
                n += 1
# ModifyTime+ModifyType+ModifyDuration
for j in res_type:
    for k in res_type:
        for h in res_mode:
            if j != "PVR" and k == "PVR":
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "ModifyTime+ModifyType+ModifyDuration", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = h
                single_list[4] = j
                single_list[8] = h
                single_list[9] = k
                total_list.append(single_list)
                n += 1
# ModifyTime+ModifyType+ModifyMode
for j in res_type:
    for k in res_type:
        for h in res_mode:
            for i in res_mode:
                if k != j and i != h:
                    single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                                   "ModifyTime+ModifyType+ModifyMode", "", "", "screen_test_numb"]
                    single_list[0] = "{0:02d}".format(n)
                    single_list[3] = h
                    single_list[4] = j
                    single_list[8] = i
                    single_list[9] = k
                    total_list.append(single_list)
                    n += 1
# ModifyType+ModifyDuration+ModifyMode
for j in res_type:
    for k in res_type:
        for h in res_mode:
            for i in res_mode:
                if j != "PVR" and k == "PVR":
                    if k != j and i != h:
                        single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                                       "ModifyType+ModifyDuration+ModifyMode", "", "", "screen_test_numb"]
                        single_list[0] = "{0:02d}".format(n)
                        single_list[3] = h
                        single_list[4] = j
                        single_list[8] = i
                        single_list[9] = k
                        total_list.append(single_list)
                        n += 1
# ModifyTime+ModifyType+ModifyDuration+ModifyMode
for j in res_type:
    for k in res_type:
        for h in res_mode:
            for i in res_mode:
                if j != "PVR" and k == "PVR":
                    if k != j and i != h:
                        single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                                       "ModifyTime+ModifyType+ModifyDuration+ModifyMode", "", "", "screen_test_numb"]
                        single_list[0] = "{0:02d}".format(n)
                        single_list[3] = h
                        single_list[4] = j
                        single_list[8] = i
                        single_list[9] = k
                        total_list.append(single_list)
                        n += 1

# print(len(total_list))
for m in range(len(total_list)):
    print(f"{total_list[m]},")
