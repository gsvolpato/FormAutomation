COORDINATES = {
    "Inclui": (217, 168),    
    "Grava": (335, 154),     
    "Ok": (503, 161),        
    "AnoEmpenho": (234, 560),
    "Empenho": (290, 563),   
    "OS": (241, 231) 
}

import time
import pandas as pd
import pyautogui
import sys
import keyboard



# Global pause variable
pause_execution = False

# Pause/Resume handler
def on_pause_press(e):
    global pause_execution
    if e.event_type == keyboard.KEY_DOWN:
        pause_execution = not pause_execution
        if pause_execution:
            print("\nPaused. Press 'Pause Break' again to resume.")
        else:
            print("\nResumed.")

# Register the pause key handler
keyboard.hook_key('pause', on_pause_press)

# Function to check if paused and wait if necessary
def check_pause():
    while pause_execution:
        time.sleep(0.1)


# 3 - Confirmation Breaks
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
    time.sleep(0.5)
    
    pyautogui.moveTo(COORDINATES["AnoEmpenho"])
    pyautogui.doubleClick(COORDINATES["AnoEmpenho"])
    pyautogui.write('2025', interval=0.01)
    pyautogui.press('tab')
    time.sleep(0.5)
    
    pyautogui.write('1615', interval=0.01)
    pyautogui.doubleClick(COORDINATES["Grava"])
    pyautogui.click(COORDINATES["Grava"])
    time.sleep(1.5)

if __name__ == "__main__":
    try:
        print("===== OS Processing Automation =====")
        print("This script will process OS numbers from your specified starting point to 364")
        print("Press 'Pause Break' to pause/resume execution")
        print("Press Ctrl+C to stop the script at any time")
        
        # Get the starting OS number from user
        while True:
            try:
                start_os = int(input("\nEnter the OS number to start from (1-364): "))
                if 1 <= start_os <= 364:
                    break
                else:
                    print("Invalid input. Please enter a number between 1 and 364.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")
        
        # Confirmation before starting
        print(f"\nReady to process OS numbers {start_os} to 364")
        input("Press Enter to begin...")
        
        print("\nStarting in 3 seconds...")
        time.sleep(3)  # Give user time to prepare
        
        for os_num in range(start_os, 365):
            click_at_positions(os_num)
            
            # Optional: add a confirmation after every N records
            if os_num % 50 == 0:
                print(f"Completed {os_num} of 364 OS numbers")
                confirm()
                
    except KeyboardInterrupt:
        print("\nScript stopped by user")
    finally:
        keyboard.unhook_all()
        input("Press Enter to close...")
