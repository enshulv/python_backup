import time
from imbox import Imbox

def receive_unread(path):
    while True:
        try:
            access = Imbox('', '', '', ssl=True)
        except Exception as err:
            print(err)
            time.sleep(60)
            continue
        with access as login:
            unread_mails = login.messages(unread=True)
            for uid, mail in unread_mails:
                attachments = mail.attachments
                sender = mail.sent_from[0]['email']
                if len(attachments) != 0:
                    email_replaced = sender.replace('.', '邮')
                    for a in attachments:
                        filename = a['filename']
                        if '\n' in filename:
                            filename = filename.replace('\n', '')
                        if '\r' in filename:
                            filename = filename.replace('\r', '')
                        attachment_name = email_replaced + '-' + filename
                        content = a['content']
                        with open(path + attachment_name, 'wb') as f:
                            f.write(content.getvalue())
                login.mark_seen(uid)
        time.sleep(60)