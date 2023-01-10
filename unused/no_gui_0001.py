import pyautogui, keyboard,time
from multiprocessing import Process
from multiprocessing import Queue

READ_UPPER_PIXELS = False
DELAY = 0.072

direction = True #True = left
pixel_main_L = None
pixel_main_R = None
pixel_upper1_L = None
pixel_upper1_R = None
pixel_upper2_L = None
pixel_upper2_R = None
proc = None
qu = None
exit_cmd = False

def execute_program(queue):
    
    pyautogui.FAILSAFE = False
    pyautogui.PAUSE = 0.000
    
    global pixel_main_L, pixel_main_R, pixel_upper1_L, pixel_upper1_R, pixel_upper2_L, pixel_upper2_R
    pixel_main_L = pyautogui.pixel(1000,856) #lower left branch
    pixel_main_R = pyautogui.pixel(1580,856) #lower right branch
    pixel_upper1_L = pyautogui.pixel(984,644) #middle left branch  #1000,656
    pixel_upper1_R = pyautogui.pixel(1570,644) #middle right branch
    pixel_upper2_L = pyautogui.pixel(1000,456) #upper left branch
    pixel_upper2_R = pyautogui.pixel(1580,456) #upper left branch


    #start_time = time.time()
    #locate_branch(DELAY)
    #print("--- %s seconds ---" % (time.time() - start_time))
    while True:
        if queue.empty():
            locate_branch()
        elif queue.get():
            dash()
        
def dash():
    if direction:
        pyautogui.press('left')
        pyautogui.press('left')
        pyautogui.press('left')
        pyautogui.press('left')
    else:
        pyautogui.press('right')
        pyautogui.press('right')
        pyautogui.press('right')
        pyautogui.press('right')
        
def locate_branch():
    global pixel_main_L, pixel_main_R, pixel_upper1_L, pixel_upper1_R, pixel_upper2_L, pixel_upper2_R, direction
    if direction:
        current_pixel = pyautogui.pixel(984,644) #upper left branch check
        if pixel_upper1_L!=current_pixel:
            pyautogui.press('left')
            time.sleep(0.003)
            pyautogui.press('right')
            direction = False
        elif READ_UPPER_PIXELS:
            current_pixel_upper1 = pyautogui.pixel(1000,656) #middle left branch check
            if pixel_upper1_L==current_pixel_upper1:
                current_pixel_upper2 = pyautogui.pixel(1000,456) #upper left branch check
                if pixel_upper2_L==current_pixel_upper2:
                    pyautogui.press('left')
                    time.sleep(0.005)
                    pyautogui.press('left')
                    time.sleep(0.005)
                    pyautogui.press('left')
                    print("triple L")
                else:
                    #print("single L 1")
                    pyautogui.press('left')
            else:
                #print("single L 2")
                pyautogui.press('left')
        else:
            pyautogui.press('left')
    else:
        current_pixel = pyautogui.pixel(1570,644) #lower right branch check
        if pixel_upper1_R!=current_pixel:
            pyautogui.press('right')
            time.sleep(0.003)
            pyautogui.press('left')
            direction = True
        elif READ_UPPER_PIXELS:
            current_pixel_upper1 = pyautogui.pixel(1580,656) #middle right branch check
            if pixel_upper1_R==current_pixel_upper1:
                current_pixel_upper2 = pyautogui.pixel(1580,456) #upper right branch check
                if pixel_upper2_R==current_pixel_upper2:
                    pyautogui.press('right')
                    time.sleep(0.005)
                    pyautogui.press('right')
                    time.sleep(0.005)
                    pyautogui.press('right')
                    print("triple R")
                else:
                    #print("single R 1")
                    pyautogui.press('right')
            else:
                #print("single R 2")
                pyautogui.press('right')
        else:
            pyautogui.press('right')
    time.sleep(DELAY)


def on_press_reaction(event):
    global proc,exit_cmd
    match event.name:
        case '1':
            proc = Process(target=execute_program,args=(qu,))
            proc.start()
        case '2':
            proc.terminate()
        case '3':
            proc.terminate()
            print("Exited normally at",time.strftime("%H:%M:%S"))
            exit_cmd = True
            
        case '4':
            qu.put(True)
            
            
        


if __name__ == '__main__':

    game_win = pyautogui.getWindowsWithTitle("Timberman")[0]
    if game_win:
        game_win.maximize()
        #game_win.move(1600,0)
    else:
        print("Timberman is not launched.")
        exit()
    
    if pyautogui.locateCenterOnScreen('img/start.png'):
        pyautogui.press('enter')
        time.sleep(1)
    
        qu = Queue()
        proc = Process(target=execute_program,args=(qu,))
        proc.start()

        keyboard.on_press(on_press_reaction)
        
        while not exit_cmd:
            pass
    else:
        print("No Start button detected.")
