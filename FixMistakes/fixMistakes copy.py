COORDINATES = {
    "Inclui": (217, 168),    
    "Grava": (335, 154),     
    "Ok": (503, 161),        
    "NF": (734, 567),
    "DATA NF": (846, 567),    
    "OS": (241, 231) 
}

import time
import pandas as pd
import pyautogui
import sys
import keyboard

pause_execution = False

def on_pause_press(e):
    global pause_execution
    if e.event_type == keyboard.KEY_DOWN:
        pause_execution = not pause_execution
        if pause_execution:
            print("\nPaused. Press 'Pause Break' again to resume.")
        else:
            print("\nResumed.")

keyboard.hook_key('pause', on_pause_press)

def check_pause():
    while pause_execution:
        time.sleep(0.1)

def confirm():
    pyautogui.hotkey('alt', 'tab')
    user_input = input("Press Enter to Continue or Q to quit: ").strip().upper()
    if user_input == "Q":
        print("Exiting...")
        exit()
    else:
        pass
    pyautogui.hotkey('alt', 'tab')

def click_at_positions(os_number):
    check_pause()
    print(f"Processing OS: {os_number}")
    
    pyautogui.moveTo(COORDINATES["OS"])
    pyautogui.doubleClick(COORDINATES["OS"])
    pyautogui.write(str(os_number), interval=0.01)
    pyautogui.press('enter')
    time.sleep(2)
    
    pyautogui.moveTo(COORDINATES["NF"])
    pyautogui.doubleClick(COORDINATES["NF"])
    pyautogui.write('2977064', interval=0.01)
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(0.5)
    
    pyautogui.write('16062025', interval=0.01)
    pyautogui.press('tab')
    time.sleep(0.5)
    
    pyautogui.doubleClick(COORDINATES["Grava"])
    pyautogui.click(COORDINATES["Grava"])
    time.sleep(1.5)

if __name__ == "__main__":
    try:
        print("===== OS Processing Automation =====")
        print("This script will process OS numbers from 460 to 472")
        print("Press 'Pause Break' to pause/resume execution")
        print("Press Ctrl+C to stop the script at any time")
        
        print(f"\nReady to process OS numbers 460 to 472")
        input("Press Enter to begin...")
        
        print("\nStarting in 3 seconds...")
        time.sleep(3)
        
        for os_num in range(460, 473):
            click_at_positions(os_num)
            
            if os_num % 5 == 0:
                print(f"Completed {os_num} of 472 OS numbers")
                confirm()
                
    except KeyboardInterrupt:
        print("\nScript stopped by user")
    finally:
        keyboard.unhook_all()
        input("Press Enter to close...")
