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
# ����һ���ж�tools.py�Ƿ������Ŀ��Ŀ¼��
tools = os.path.abspath(r"D:\fcs_function\tools.py")
for root, dirs, files in os.walk(os.path.dirname(tools)):
    if os.path.basename(tools) in files:
        print("����һ���ж�tools.py�Ƿ������Ŀ��Ŀ¼��", tools)
# ���������ж�tools.py�Ƿ������Ŀ��Ŀ¼��
if os.path.lexists(r"D:\fcs_function\tools1.py"):
    print("���������ж�tools.py�Ƿ������Ŀ��Ŀ¼��", tools)

path = "./"
ini_list = []
for filename in os.listdir(path):
    if filename.endswith('ini'):
        ini_list.append(filename)
print(ini_list)
jsondict = {
    "data": {
        "fileHash": "f07e50bbe581489dad324c096dc9c19c75057220008797a2110aa3254a537afc8",
        "code": 0,
        "destFileName": "1",
        "srcFileName": "1.docx",
        "srcFileSize": "101217",
        "destFileSize": "-1",
        "convertType": 0,
        "srcStoragePath": "f07e50bbe581489dad324c096dc9c19c7/1.docx",
        "destStoragePath": "2023/02/06/8ec997e1fd9945f9be72fd0040b1d7da/1",
        "convertTime": "185",
        "viewUrl": "http://172.18.21.30:8080/fcscloud/view/preview/mHR2CLRPjFerGhujYGcR1wv1QLoi09rAbL59IGBRoRw58KSmdwmS_NH-ZI3wU_ZUrkTJGCPUMFo_HI0y9Gh_Y1tB2eUP6CgrrpcBYZb3Ul71pbfcbLwKyRCKXnGYTuGod-QOSNuMzFX3mKPf2qziNIUO7e7rmNvhCcgKEYbpuwfy8LvDtinX-sYT3wNAyd2f1HfyJx-azQhsti9G6pkIJ0gWUrtJYfsiZbQE3tVqLzR-JfOCaw2NtIZE_PXjAUlMOpNKfHjo4VsX3WBvlNerKeE9A_T33H0ovbaBC10ipOpIEQgz9YaFOCMlfmZyglGju4mdncKxeX-OBzdy6CcXW5DCp6VD9QzgCW3UkF10D3qw-z2hIdpmrvJtk83qFIrIrvZjnJM2ClbG7qxHEWkH77FPzKRVMK8oilqCk2nsc3LlNLDXWaBHFkB4_iolHuu81kGsL_QjDzsLs7yup_AEZMdfEVCSW3zPveqVFkXgk3YQzTrJ4PIsCUVSmCBxfQBLKnXOcOV5NYQDDv9Blqt5ZOlz0mqzlEY5ImduN3EI2_Xt3Q3qDRA0azKbHJ9QbFh27PUEeCs_Sap76f1arFVwRkRohzteqdw3/"
    },
    "message": "�����ɹ�",
    "errorcode": 0
}
# python����ֱ�Ӷ�ȡǶ���ֵ䣡
print(jsondict["data"]["viewUrl"])
