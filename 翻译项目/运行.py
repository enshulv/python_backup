from threading import Thread
from 接收文件 import 接收未读
from 估价 import 自动估价,发送结果

路径=r''
地址=r""
保存=r""
待发送=r''
保存结果=r''

def run():
    t1=Thread(target=接收未读,args=(路径,))
    t2=Thread(target=自动估价,args=(地址,1,))
    t3=Thread(target=发送结果,args=(待发送,保存结果,))
    t1.start()
    t2.start()
    t3.start()

run()
#发送结果(待发送,保存结果)
#列表.pop(0)
