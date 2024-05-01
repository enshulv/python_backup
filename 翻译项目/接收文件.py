import time
from imbox import Imbox

def 接收未读(地址):
    while True:
        try:
            访问=Imbox('','','',ssl=True)
        except Exception as err:
            print(err)
            time.sleep(60)
            continue
        with 访问 as 登录:
                未读邮件=登录.messages(unread=True)
                for uid,邮件 in 未读邮件:
                    附件=邮件.attachments
                    发件人=邮件.sent_from[0]['email']
                    if len(附件) != 0:
                        邮箱替换=发件人.replace('.', '邮')
                        for a in 附件:
                            文件名=a['filename']
                            if '\n' in 文件名:
                                文件名=文件名.replace('\n','')
                            if '\r' in 文件名:
                                文件名 = 文件名.replace('\r', '')
                            附件名 = 邮箱替换 + '-' +文件名
                            内容=a['content']
                            with open(地址+附件名,'wb') as f:
                                f.write(内容.getvalue())
                    登录.mark_seen(uid)
        time.sleep(60)



