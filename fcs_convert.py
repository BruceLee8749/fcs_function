# coding=utf-8
from main import Test
from tools import *
import traceback

if __name__ == '__main__':
    """接口转换"""
    try:
        for xlsx, sheets in get_target_cases('FCS').items():
            result_excel = copy_test_excel('FCS', xlsx)
            set_conf('FCS', '测试结果文件夹', result_excel)
            print('--------------------------正在运行%s--------------------------------------' % xlsx)
            for sheet in sheets:
                Test('FCS', result_excel, sheet).run()
    except:
        print(traceback.format_exc())
        critical(traceback.format_exc())
