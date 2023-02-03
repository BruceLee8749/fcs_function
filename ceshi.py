# coding=gbk
import os

dict = {"noCache": 1, "wmContent": "基线自动化", "wmSize": "10", "wmColor": "#0000FF", "wmFont": "宋体",
        "wmMargin": "30,50",
        "wmTransparency": "0.5", "wmRotate": "45", "wmPosition": "100,800"}
for key, value in dict.items():
    print(key)
    print("========================")
    print(value)

abs = os.path.abspath('ceshi.py')
print("打印文件的绝对路径，给个相对路径也能打印：", abs)
print("输入当前文件路径：", os.path.abspath(__file__))
