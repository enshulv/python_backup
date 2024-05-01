from 翻译工具 import 翻译
from 公式全翻译 import 运行翻译
import time
import os
import shutil

def 运行(名字):
    位置 = 'E:\Desktop\文件\保存//'
    保存 = 'E:\Desktop\文件\结果//'
    try:
        #正文翻译引擎,跳过参考,表格翻译引擎,一中一英
        if 'en-' in 名字[:5]:
            翻译(名字, 'en', 位置, 保存,1,0,1,0)
        elif 'qk-' in 名字[:5]:
            翻译(名字, 'zh', 位置, 保存,1,0,1,0)
        elif 'gs-' in 名字[:5]:
            运行翻译('zh',名字,位置,保存)
        elif 'gn-' in 名字[:5]:
            运行翻译('en',名字,位置,保存)
        else:
            翻译(名字, 'zh', 位置, 保存,0,0,1,0)
    except Exception as 错误:
        print('\n' + str(错误))

def run():
    位置 = 'E:\Desktop\文件\保存//'
    备份='E:\Desktop\文件\保存//保存//'
    待翻译=os.listdir(位置)
    for a in 待翻译:
        位置1=位置+a
        备份1=备份+a
        if 'docx' in a and '~$' not in a:
            名字=a.replace('.docx','')
            print(名字)
            运行(名字)
            shutil.move(位置1,备份1)

if __name__ == '__main__':
    run()
