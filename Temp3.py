import thread

def input_thread(a_list):
    raw_input()
    a_list.append(True)

def do_stuff():
    a_list = []
    thread.start_new_thread(input_thread, (a_list,))
    while not a_list:
        stuff()


