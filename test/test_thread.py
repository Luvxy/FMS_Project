import time
import threading as th
import keyboard
import sys
import _thread

is_true = False

def bread():
    count = 0
    while True:
        print("bread is running")
        time.sleep(1)
        count += 1
        if count == 5:
            break

def background_task():
    global is_true
    while True:
        # 주기적으로 실행되어야 하는 작업 시뮬레이션
        if keyboard.is_pressed('esc'):
            print("esc pressed")
            _thread.interrupt_main()
            break
        
if __name__ == "__main__":
    # 백그라운드 작업을 담당할 스레드 생성
    background_thread = th.Thread(target=background_task)
    # 메인 작업을 실행하는 스레드 생성
    main_thread = th.Thread(target=bread)
    # 백그라운드 스레드 시작
    background_thread.start()
    # 메인 스레드 시작
    main_thread.start()
    # 백그라운드 스레드가 종료되면 프로그램 종료
    background_thread.join()        
    # 메인 스레드 종료
    print("Main thread exits")
