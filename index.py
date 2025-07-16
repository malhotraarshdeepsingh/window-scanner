import pygetwindow as gw
import pyautogui
import pytesseract
import cv2
import numpy as np
import time
import re

# Tesseract Path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

while True:
    print("\n🕐 Waiting 5 seconds before screenshot...")
    time.sleep(5)

    # Get window by title
    window = None
    for w in gw.getWindowsWithTitle('Interactive Brokers'):  # change title as needed
        if w.visible:
            window = w
            break

    if not window:
        print("❌ Window not found!")
    else:
        # Bring to front
        window.activate()
        time.sleep(1)

        # Capture region
        left, top, width, height = window.left, window.top, window.width, window.height
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        screenshot.save("window_screenshot.png")

        # Process image
        image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

        # OCR
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(thresh, config=custom_config)

        # Parse stock data
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

        # Display
        print("\n📈 Parsed Stock Data (Symbol → Prices):\n")
        for symbol, prices in parsed_data:
            print(f"{symbol}: {', '.join(prices)}")

    # ⏱ Wait 1 minute before next run
    print("\n🔁 Waiting 10 sec before next run...\n")
    time.sleep(10)
