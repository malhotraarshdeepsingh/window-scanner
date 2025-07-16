import pyautogui
import pytesseract
import cv2
import numpy as np
import time
import re
import win32gui
import win32con
import win32com.client

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def find_window_by_partial_title(substring):
    def callback(hwnd, result):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if substring.lower() in title.lower():
                result.append(hwnd)
    result = []
    win32gui.EnumWindows(callback, result)
    return result[0] if result else None

while True:
    print("\n🕐 Waiting 5 seconds before screenshot...")
    time.sleep(5)

    window_title_partial = "Interactive Brokers"
    hwnd = find_window_by_partial_title(window_title_partial)

    if not hwnd:
        print("❌ Window not found!")
    else:
        print("✅ Window found!")

        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(2)

        rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = rect
        width = right - left
        height = bottom - top

        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        screenshot.save("window_screenshot.png")

        image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(thresh, config=custom_config)

        lines = text.split('\n')
        parsed_data = []
        for line in lines:
            line = line.strip()
            if not line or len(line) < 3:
                continue
            symbols = re.findall(r'\b[A-Z]{2,5}\b', line)
            prices = re.findall(r'\d+\.\d+', line)
            if symbols and prices:
                parsed_data.append((symbols[0], prices))

        print("\n📈 Parsed Stock Data (Symbol → Prices):\n")
        for symbol, prices in parsed_data:
            print(f"{symbol}: {', '.join(prices)}")

    print("\n🔁 Waiting 10 sec before next run...\n")
    time.sleep(10)