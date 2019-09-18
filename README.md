Infrared send

NewSatSearch.py   "该程序为调试使用，不包含sys.argv参数的，可以直接执行"

NewSatSearchIncludeArgvPara.py    "该程序为给执行测试人员使用，包含sys.argv参数，需要在AutoSearchTest.py文件中执行"

AutoSearchTest.py   "该程序为自动化测试的程序，需与NewSatSearchIncludeArgvPara.py结合使用"

AutoSearch.py     "该程序为类管理的自动化测试的程序，需与SatSearchIncludeArgvPara_Class.py结合使用"

NewSatSearch.py   "该程序为更新KWS后重新用类管理变量的程序，调试使用，不包含sys.argv参数的"

SatSearchIncludeArgvPara_Class.py
"该程序为更新KWS后重新用类管理变量的程序，测试使用，包含sys.argv参数，需要在AutoSearch.py文件中执行"

PrsetSatBlind.py   "该程序用于预置卫星搜索，有卫星和卫星下的参数的判断和设置"