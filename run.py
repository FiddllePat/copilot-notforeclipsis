import os
import numpy as np
import cv2
import torch
import time
import win32api
import pyautogui
from pynput import keyboard
import asyncio
from io import BytesIO
from PIL import Image
import ctypes
import win32gui
import win32ui
import win32con

def clear():
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')

def print_banner(fps, visible_objects):
    banner = f"""
  ______     _ _           _      
 |  ____|   | (_)         (_)     
 | |__   ___| |_ _ __  ___ _ ___  
 |  __| / __| | | '_ \/ __| / __| 
 | |___| (__| | | |_) \__ \ \__ \ 
 |______\___|_|_| .__/|___/_|___/ 
  / ____|       | |(_) |     | |  
 | |     ___  _ |_| _| | ___ | |_ 
 | |    / _ \| '_ \| | |/ _ \| __|
 | |___| (_) | |_) | | | (_) | |_ 
  \_____\___/| .__/|_|_|\___/ \__|
      |___ \ | |                  
 __   ____) ||_|                  
 \ \ / /__ <                      
  \ V /___) |                     
   \_/|____/                      
                                  
                                  
Frames Per Second : {fps:.2f}
Currently Visible: {len(visible_objects)} objects
{visible_objects}
"""
    return banner

print("Loading model...")
model = torch.hub.load('ultralytics/yolov5', 'custom', path='v3.pt')
print("Model loaded.")
clear()

running = False
track_other_outputs = False
tracked_classes = [0, 2, 3]
current_frame = None

def on_press(key):
    global running, track_other_outputs
    if key == keyboard.KeyCode(char='['):
        running = not running
        if running:
            print("Loop started.")
        else:
            print("Loop stopped.")
    elif key == keyboard.KeyCode(char='='):
        track_other_outputs = not track_other_outputs
        if track_other_outputs:
            print("Tracking all other outputs.")
        else:
            print("Tracking only specified classes.")

listener = keyboard.Listener(on_press=on_press)
listener.start()

last_state = None
start_time = time.time()
frame_count = 0

center_point = win32api.GetCursorPos()

def capture_screenshot(left, top, width, height):
    # Calculate the top-left and bottom-right coordinates of the region around the cursor
    rect_top_left = (max(0, left - width // 2), max(0, top - height // 2))
    rect_bottom_right = (min(pyautogui.size()[0], left + width // 2), min(pyautogui.size()[1], top + height // 2))
    region_width = rect_bottom_right[0] - rect_top_left[0]
    region_height = rect_bottom_right[1] - rect_top_left[1]
    
    # Capture the screenshot of the region around the cursor
    hdesktop = win32gui.GetDesktopWindow()
    desktop_dc = win32gui.GetWindowDC(hdesktop)
    img_dc = win32ui.CreateDCFromHandle(desktop_dc)
    mem_dc = img_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(img_dc, region_width, region_height)
    mem_dc.SelectObject(bitmap)
    mem_dc.BitBlt((0, 0), (region_width, region_height), img_dc, (rect_top_left[0], rect_top_left[1]), win32con.SRCCOPY)
    bmpinfo = bitmap.GetInfo()
    bmpstr = bitmap.GetBitmapBits(True)
    img = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1)
    win32gui.DeleteObject(bitmap.GetHandle())
    mem_dc.DeleteDC()
    img_dc.DeleteDC()
    win32gui.ReleaseDC(hdesktop, desktop_dc)
    
    return img

async def main_loop():
    global current_frame, frame_count, last_state, center_point

    while listener.running:
        if running:
            cursor_pos = win32api.GetCursorPos()
            if cursor_pos != center_point:
                center_point = cursor_pos
            width, height = 200, 200  # Define the size of the region around the cursor
            screenshot = capture_screenshot(cursor_pos[0], cursor_pos[1], width, height)
            screen = np.array(screenshot)

            results = model(screen)

            visible_objects = []

            if track_other_outputs:
                for result in results.xyxy[1:]:
                    if int(result[5]) in tracked_classes:
                        center_x = int((result[0] + result[2]) / 2) + cursor_pos[0] - width // 2
                        center_y = int((result[1] + result[3]) / 2) + cursor_pos[1] - height // 2
                        visible_objects.append((center_x, center_y, result[4].item()))
                        pyautogui.moveTo(center_x, center_y)
            else:
                for result in results.xyxy[0]:
                    if int(result[5]) in tracked_classes:
                        center_x = int((result[0] + result[2]) / 2) + cursor_pos[0] - width // 2
                        center_y = int((result[1] + result[3]) / 2) + cursor_pos[1] - height // 2
                        visible_objects.append((center_x, center_y, result[4].item()))
                        pyautogui.moveTo(center_x, center_y)

            frame_count += 1
            elapsed_time = time.time() - start_time
            fps = frame_count / elapsed_time

            current_state = (fps, visible_objects)
            if current_state != last_state:
                clear()
                print(print_banner(fps, visible_objects))
                last_state = current_state

            current_frame = screen

asyncio.run(main_loop())
