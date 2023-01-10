import win32com.client as comclt
import keyboard, time, mss, win32gui
from multiprocessing import Process
from multiprocessing import Queue

READ_UPPER_PIXELS = False
DELAY = 0.006

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
wsh = None

def get_pixel(left,top):
    with mss.mss() as sct:
        pic = sct.grab({'mon':1, 'top':top, 'left':left, 'width':1, 'height':1})
        return pic.pixel(0,0)

def execute_program(queue):
    global pixel_main_L, pixel_main_R, pixel_upper1_L, pixel_upper1_R, pixel_upper2_L, pixel_upper2_R,wsh
    #pixel_main_L = get_pixel(1000,856) #lower left branch
    #pixel_main_R = get_pixel(1580,856) #lower right branch
    pixel_upper1_L = get_pixel(984,590) #middle left branch  #1000,656
    pixel_upper1_R = get_pixel(1570,590) #middle right branch
    #pixel_upper2_L = get_pixel(1000,456) #upper left branch
    #pixel_upper2_R = get_pixel(1580,456) #upper left branch

    #start_time = time.time()
    #locate_branch(DELAY)
    #print("--- %s seconds ---" % (time.time() - start_time))
    wsh.SendKeys("{LEFT}")
    while True:
        if queue.empty():
            locate_branch()
        elif queue.get():
            dash()
        
def dash():
    if direction:
        wsh.SendKeys("{LEFT}")
        wsh.SendKeys("{LEFT}")
        wsh.SendKeys("{LEFT}")
        wsh.SendKeys("{LEFT}")
    else:
        wsh.SendKeys("{RIGHT}")
        wsh.SendKeys("{RIGHT}")
        wsh.SendKeys("{RIGHT}")
        wsh.SendKeys("{RIGHT}")
        
def locate_branch():
    global pixel_main_L, pixel_main_R, pixel_upper1_L, pixel_upper1_R, pixel_upper2_L, pixel_upper2_R, direction,wsh
    if direction:
        current_pixel = get_pixel(984,590) #upper left branch check
        if pixel_upper1_L!=current_pixel:
            wsh.SendKeys("{LEFT}")
            time.sleep(0.003)
            wsh.SendKeys("{RIGHT}")
            direction = False
        elif READ_UPPER_PIXELS:
            current_pixel_upper1 = get_pixel(1000,656) #middle left branch check
            if pixel_upper1_L==current_pixel_upper1:
                current_pixel_upper2 = get_pixel(1000,456) #upper left branch check
                if pixel_upper2_L==current_pixel_upper2:
                    wsh.SendKeys("{LEFT}")
                    time.sleep(0.005)
                    wsh.SendKeys("{LEFT}")
                    time.sleep(0.005)
                    wsh.SendKeys("{LEFT}")
                    print("triple L")
                else:
                    #print("single L 1")
                    wsh.SendKeys("{LEFT}")
            else:
                #print("single L 2")
                wsh.SendKeys("{LEFT}")
        else:
            wsh.SendKeys("{LEFT}")
    else:
        current_pixel = get_pixel(1570,590) #lower right branch check
        if pixel_upper1_R!=current_pixel:
            wsh.SendKeys("{RIGHT}")
            time.sleep(0.003)
            wsh.SendKeys("{LEFT}")
            direction = True
        elif READ_UPPER_PIXELS:
            current_pixel_upper1 = get_pixel(1580,656) #middle right branch check
            if pixel_upper1_R==current_pixel_upper1:
                current_pixel_upper2 = get_pixel(1580,456) #upper right branch check
                if pixel_upper2_R==current_pixel_upper2:
                    wsh.SendKeys("{RIGHT}")
                    time.sleep(0.005)
                    wsh.SendKeys("{RIGHT}")
                    time.sleep(0.005)
                    wsh.SendKeys("{RIGHT}")
                    print("triple R")
                else:
                    #print("single R 1")
                    wsh.SendKeys("{RIGHT}")
            else:
                #print("single R 2")
                wsh.SendKeys("{RIGHT}")
        else:
            wsh.SendKeys("{RIGHT}")
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

    wsh = comclt.Dispatch("WScript.Shell")
    act = wsh.AppActivate("Timberman")
    if act:
        game_win = win32gui.FindWindow(None, "Timberman")
        win32gui.MoveWindow(game_win,1600,0,720,480,True)
    else:
        print("Timberman is not launched.")
        exit()
    
    #if wsh.locateCenterOnScreen('img/start.png'):
    #   wsh.SendKeys('enter')
    time.sleep(1)

    qu = Queue()
    proc = Process(target=execute_program,args=(qu,))
    proc.start()

    keyboard.on_press(on_press_reaction)
    
    while not exit_cmd:
        pass
    #else:
    #    print("No Start button detected.")
