#!/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import shutil
import platform

sat_search_mode_list = [
                            "z6_blind",                         # 0
                            "z6_super_blind",                   # 1

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

choice_sat_search_mode_numb = 0

parent_of_current_path = os.path.abspath(os.path.join(os.getcwd(), ".."))        # 当前程序路径的上级路径
main_prog_path = os.path.join(parent_of_current_path, "MainProgram", "NewAddSatBlind_IncludeArgvParam.py")

shutil.copy(main_prog_path, os.getcwd())
if platform.system() == 'Windows':
    os.system('python ./NewAddSatBlind_IncludeArgvParam.py %d' % choice_sat_search_mode_numb)
elif platform.system() == 'Linux':
    os.system('python3 ./NewAddSatBlind_IncludeArgvParam.py %d' % choice_sat_search_mode_numb)
os.unlink(os.path.join(os.getcwd(), 'NewAddSatBlind_IncludeArgvParam.py'))
