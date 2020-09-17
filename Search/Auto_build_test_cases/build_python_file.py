#!/usr/bin/python3
# -*- coding: utf-8 -*-

sat_search_mode_list = [
                            "6b_blind",                         # 0
                            "6b_super_blind",                   # 1

                            "y3_blind",                         # 2
                            "y3_super_blind",                   # 3

                            "138_blind",                        # 4
                            "138_super_blind",                  # 5

                            "88_blind",                         # 6
                            "88_super_blind",                   # 7

                            "plp_blind",                        # 8
                            "plp_super_blind",                  # 9

                            "138_incremental_blind",            # 10 累加搜索

                            "138_ch_upper_limit_blind",          # 11 搜索节目达到上限,会删除所有节目,重新搜索
                            "138_ch_ul_later_cont_blind",        # 12 搜索节目达到上限后,不删除指定卫星下的tp,继续搜索
                            "138_ch_ul_later_del_tp_blind",      # 13 搜索节目达到上限后,删除指定卫星下的tp,继续搜索

                            "z6_tp_upper_limit_blind",          # 14 搜索tp达到上限,会恢复出厂设置,重新搜索
                            "z6_tp_ul_later_cont_blind",        # 15 搜索tp达到上限后,不删除指定卫星下的tp,继续搜索
                            "z6_tp_ul_later_del_tp_blind",      # 16 搜索tp达到上限后,删除指定卫星下的tp,继续搜索

                            "reset_factory",                    # 17 恢复出厂设置
                            "delete_all_channel",               # 18 删除所有节目
                        ]

read_data = []

with open("new_file.py", "r", encoding='UTF-8') as f:
    for line in f.readlines():
        print(line, end='')
        read_data.append(line)


def not_need_reset_factory(file_path):
    datas = [
        "shutil.copy(main_prog_path, os.getcwd())",
        "if platform.system() == 'Windows':",
        "    os.system('python ./NewAddSatBlind_IncludeArgvParam.py %d' % choice_sat_search_mode_numb)",
        "elif platform.system() == 'Linux':",
        "    os.system('python3 ./NewAddSatBlind_IncludeArgvParam.py %d' % choice_sat_search_mode_numb)",
        "os.unlink(os.path.join(os.getcwd(), 'NewAddSatBlind_IncludeArgvParam.py'))",
        ]
    with open(file_path, "a+", encoding='UTF-8') as fo:
        for data in datas:
            fo.write('{}\n'.format(data))


def need_reset_factory(file_path):
    datas = [
        "shutil.copy(main_prog_path, os.getcwd())",
        "if platform.system() == 'Windows':",
        "    os.system('python ./NewAddSatBlind_IncludeArgvParam.py %d' % 17)",
        "    os.system('python ./NewAddSatBlind_IncludeArgvParam.py %d' % choice_sat_search_mode_numb)",
        "elif platform.system() == 'Linux':",
        "    os.system('python3 ./NewAddSatBlind_IncludeArgvParam.py %d' % 17)",
        "    os.system('python3 ./NewAddSatBlind_IncludeArgvParam.py %d' % choice_sat_search_mode_numb)",
        "os.unlink(os.path.join(os.getcwd(), 'NewAddSatBlind_IncludeArgvParam.py'))",
        ]
    with open(file_path, "a+", encoding='UTF-8') as fo:
        for data in datas:
            fo.write('{}\n'.format(data))


def build_x_numb_python_file():
    for i in range(len(sat_search_mode_list)):
        if i < 17:
            if i < 10:
                sequence_numb = "0{}".format(i)
                python_file_path = "test_{}_{}.py".format(sequence_numb, sat_search_mode_list[i])
                with open(python_file_path, "a+", encoding='UTF-8') as fo:
                    for j in range(len(read_data)):
                        if "choice_sat_search_mode_numb =" in read_data[j]:
                            fo.write("choice_sat_search_mode_numb = {}\n".format(i))
                        else:
                            fo.write(read_data[j])
                not_need_reset_factory(python_file_path)
            else:
                sequence_numb = "{}".format(i)
                python_file_path = "test_{}_{}.py".format(sequence_numb, sat_search_mode_list[i])
                with open(python_file_path, "a+", encoding='UTF-8') as fo:
                    for j in range(len(read_data)):
                        if "choice_sat_search_mode_numb =" in read_data[j]:
                            fo.write("choice_sat_search_mode_numb = {}\n".format(i))
                        else:
                            fo.write(read_data[j])
                need_reset_factory(python_file_path)


build_x_numb_python_file()
