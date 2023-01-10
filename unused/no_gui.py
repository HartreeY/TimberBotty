import win32com.client as comclt
import cv2 as cv
import keyboard, time, mss, win32gui
import numpy as np

READ_UPPER_PIXELS = False
DELAY = 0.13

direction = True #True = left
pixel_main_L = None
pixel_main_R = None
pixel_upper1_L = None
pixel_upper1_R = None
pixel_upper2_L = None
pixel_upper2_R = None

exit_cmd = False
wsh = None
paused = False

def get_pixel(left,top):
    with mss.mss() as sct:
        pic = sct.grab({'mon':1, 'top':top, 'left':left, 'width':1, 'height':1})
        return pic.pixel(0,0)

def start_pixels():
    global pixel_main_L, pixel_main_R, pixel_upper1_L, pixel_upper1_R, pixel_upper2_L, pixel_upper2_R,wsh
    #pixel_main_L = get_pixel(1000,856) #lower left branch
    #pixel_main_R = get_pixel(1580,856) #lower right branch
    pixel_upper1_L = get_pixel(1872,234) #middle left branch  #1000,656
    pixel_upper1_R = get_pixel(2050,234) #middle right branch
    #pixel_upper2_L = get_pixel(1000,456) #upper left branch
    #pixel_upper2_R = get_pixel(1580,456) #upper left branch

    #start_time = time.time()
    #locate_branch(DELAY)
    #print("--- %s seconds ---" % (time.time() - start_time))
        
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
        current_pixel = get_pixel(1872,234) #upper left branch check
        if pixel_upper1_L!=current_pixel:
            wsh.SendKeys("{LEFT}")
            time.sleep(DELAY)
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
        current_pixel = get_pixel(2050,234) #lower right branch check
        if pixel_upper1_R!=current_pixel:
            wsh.SendKeys("{RIGHT}")
            time.sleep(DELAY)
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
    global paused,exit_cmd
    match event.name:
        case '1':
            if paused:
                start_pixels()
                paused = False
        case '2':
            paused = True
        case '3':
            exit_cmd = True
        case '4':
            dash()
            
            
        


if __name__ == '__main__':

    wsh = comclt.Dispatch("WScript.Shell")
    act = wsh.AppActivate("Timberman")
    if act:
        game_win = win32gui.FindWindow(None, "Timberman")
        win32gui.MoveWindow(game_win,1600,0,720,480,True)
    else:
        print("Timberman is not launched.")
        exit()
    
    START_PICTURE_CV = cv.imread("img/start.png")
    with mss.mss() as sct:
        pic = sct.grab({'mon':1, 'top':0+415, 'left':1600+325, 'width':80, 'height':80})
        pic_cv = cv.cvtColor(np.array(pic), cv.COLOR_RGB2BGR)

    #cv.imshow('img',pic_cv)
    #cv.waitKey(0)
    #cv.destroyAllWindows()

    comparison = cv.matchTemplate(pic_cv, START_PICTURE_CV, cv.TM_CCOEFF_NORMED)
    if (comparison >= 0.8).any():
        wsh.SendKeys("{ENTER}")
        time.sleep(2)
        wsh.SendKeys("{LEFT}")
        start_pixels()
        
        keyboard.on_press(on_press_reaction)

        while not exit_cmd:
            if not paused:
                locate_branch()
        
    else:
        print("No Start button detected.")

    print("Exited normally at",time.strftime("%H:%M:%S"))