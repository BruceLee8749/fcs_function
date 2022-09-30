# coding=utf-8
from main import Test
from tools import *
import traceback

if __name__ == '__main__':
    """转换失败后重跑"""
    try:
        rerun_xlsx = input('请输入需要rerun的FCS结果表名：')
        print('--------------------------正在运行%s--------------------------------------' % rerun_xlsx)
        for sheet in get_case_sheet_name('FCS测试结果', rerun_xlsx):
            Test('FCS', os.path.join('FCS测试结果', rerun_xlsx), sheet).run()
    except:
        print(traceback.format_exc())
        critical(traceback.format_exc())

