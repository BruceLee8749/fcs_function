# coding=gbk
import os

dict = {"noCache": 1, "wmContent": "�����Զ���", "wmSize": "10", "wmColor": "#0000FF", "wmFont": "����",
        "wmMargin": "30,50",
        "wmTransparency": "0.5", "wmRotate": "45", "wmPosition": "100,800"}
for key, value in dict.items():
    print(key)
    print("========================")
    print(value)

abs = os.path.abspath('ceshi.py')
print("��ӡ�ļ��ľ���·�����������·��Ҳ�ܴ�ӡ��", abs)
print("���뵱ǰ�ļ�·����", os.path.abspath(__file__))
