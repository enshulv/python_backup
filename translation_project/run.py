from threading import Thread
from receive_files import receive_unread
from word_count import auto_estimate, send_results

path = r''
address = r""
save = r""
to_send = r''
save_results = r''

def run():
    t1 = Thread(target=receive_unread, args=(path,))
    t2 = Thread(target=auto_estimate, args=(address, 1,))
    t3 = Thread(target=send_results, args=(to_send, save_results,))
    t1.start()
    t2.start()
    t3.start()

run()
# send_results(to_send, save_results)
# list_.pop(0)