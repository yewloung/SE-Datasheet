import sys
import keyboard

a=[1,2,3,4]
print("Press Enter to continue or press Esc to exit: ")
while True:
    try:
        if keyboard.is_pressed('ENTER'):
            print("you pressed Enter, so printing the list..")
            print(a)
            break
        if keyboard.is_pressed('Esc'):
            print("\nyou pressed Esc, so exiting...")
            sys.exit(0)
    except:
        break
