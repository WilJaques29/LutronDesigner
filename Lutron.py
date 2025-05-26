import sys
import pytesseract
from PIL import ImageGrab, Image, ImageEnhance, ImageOps
import pyautogui
import threading
import csv
import time
import tkinter as tk
import keyboard
from pdf2image import convert_from_path
import pyperclip
import re
import string
import tkinter as tk
from tkinter import filedialog
# Initialize the array with each element being a list of the specified details

import pandas as pd

# Load the XLS file
file_path = r"C:\Users\Wiltj\OneDrive - Maverick Lite\_Maverick Lite Client\_WSH\Behrens, 469 E Flamingo_8280\Loads_Rooms_new.xls"  # Update this path if needed
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\tesseract.exe"

room_data = {}
last_checked_room = None
zone_data = {}
device_data = {}
equipment_data = {}
shade_data = []
fixture_data = {}
keypadTargets = {
        "Place:": (0,0),
        "seeTouch": (0,0),
        "controls": (0, 0),
        "Sensors": (0, 0),
        "Ceiling Occ RF": (0,0),
        "Wall Keypads": (0, 0),
        "Next": (0,0)
    }
loadsTargets = {
    "Place:": (0,0),
    "Add": (0,0),
    "Edit Fixture Type": (0,0),
    "Next": (0,0),
}
fixtureTargets = {
    "Done": (0,0),
    # if I need to add lutron lights
    "Lutron Lights": (0,0),
}

shadeTargets = {
    "Add shade group": (0,0),
    "Shade Group 1": (0,0),
    "Place:": (0,0),
    "Next": (0,0),
}

equipmentTargets = {
    #repeater
    "Hybrid": (0,0),
    "Clear": (0,0),
    "Place:": (0,0),
    "Next": (0,0),
    "Panels": (0,0),
    #device tab
    "Devices": (0,0),
    "LV-21": (0,0),
    # shade panel
    "Smart": (0,0),
    # if I need to add lutron lights
    "Lutron Lights": (0,0),
}

# for comparing excel files

added_rooms = set()
removed_rooms = set()
changed_rooms = set()

added_zones = set()
removed_zones = set()
changed_zones = set()

added_keypads = set()
removed_keypads = set()
changed_keypads = set()

added_repeaters = set()
removed_repeaters = set()
changed_repeaters = set()

def checkingLoads(room_num):
    """Ensure all expected loads for the given room number are visible on screen."""
    screenshot = ImageGrab.grab()
    text_data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)

    # Combine all OCR words into a searchable set
    screen_text = set(word.strip().lower() for word in text_data["text"] if word.strip())

    # Get all zone_names for the room
    expected_loads = []
    for zone, data in zone_data.items():
        if extract_room_number(zone) == room_num:
            expected_loads.append(data["zone_name"].strip().lower())

    # Check which loads are missing
    missing = [name for name in expected_loads if name not in screen_text]

    if missing:
        print(f"❌ Error: Missing loads in Room {room_num}: {missing}")
    else:
        print(f"✅ All loads visible on screen for Room {room_num}")

def findPreviousFloor(floor):
    '''finding previous floor'''
    # Take full screenshot
    screenshot = ImageGrab.grab()
    # Run OCR
    floor_lower = floor.lower()
    text_data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
    # Loop through OCR results and look for "Mech (102)"
    for i in range(len(text_data['text'])):
        word = text_data['text'][i].strip()
        if not word:
            continue
        if floor_lower in word.lower():
            x = text_data['left'][i] + text_data['width'][i] // 2
            y = text_data['top'][i] + text_data['height'][i] // 2
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.5)
            return
    print("Error: Previous Floor not found.", floor)

def load_excel_file(file_path, ketraLights):
    ketra = False
    df = pd.read_excel(file_path)

    # Clean up column names
    df.columns = [col.strip() for col in df.columns]

    # Global zone_data dictionary
    global zone_data
    zone_data = {}

    for _, row in df.iterrows():
        zone_raw = row.get("ZONE")
        if zone_raw and pd.notna(zone_raw):
            zone = str(zone_raw).strip()
            fixture_count_raw = row.get("FIXTURE,COUNT", "")
            fixture_count = str(fixture_count_raw).strip()
            zone_name = str(row.get("ZONE_NAME", "")).strip()
            fixtures = []
            ketra = False

            if zone[0] != "S":
                # Prompt if no fixture data
                while not fixture_count or fixture_count.lower() == 'nan':
                    print(f"Missing fixture for ZONE {zone} (ZONE_NAME: '{zone_name}').")
                    fixture_count = input("Enter fixture and count (e.g., AA,3 or AA,3 BB,2): ").strip()
            else:
                shade_data.append(zone)
                continue
            # Parse multiple fixture,count pairs
            entries = fixture_count.split()
            for entry in entries:
                if ',' in entry:
                    pair = entry.split(',')
                    if len(pair) != 2:
                        print(f"❌ Malformed entry '{entry}' in ZONE {zone}. Skipping.")
                        continue

                    fixture = pair[0].strip()
                    count = pair[1].strip()

                    if not fixture:
                        fixture = input(f"Enter fixture for ZONE {zone} (ZONE_NAME: '{zone_name}'): ").strip()

                    if not count:
                        count = input(f"Enter count for ZONE {zone} (ZONE_NAME: '{zone_name}'): ").strip()

                    if fixture in ketraLights:
                        ketra = True
                        print("Ketra Lights:", fixture)

                    fixtures.append({
                        "fixture": fixture,
                        "count": int(count) if count.isdigit() else count
                    })

            if zone:
                zone_data[zone] = {
                    "fixtures": fixtures,
                    "zone_name": zone_name,
                    "location": (0, 0),
                    "Ketra": ketra
                }


    # Room data: Room number as key
    for _, row in df.iterrows():
        room_bottom = row.get("ROOM_NAME_BOTTOM", "")
        room_main = row.get("ROOM_NAME", "")
        room_3rd = row.get("ROOM_NAME_3RD_LINE", "")

        if pd.isna(room_3rd):
            continue
        room_key = int(room_3rd)

        name_parts = [str(part).strip() for part in [room_main, room_bottom] if pd.notna(part) and str(part).strip()]
        full_name = " ".join(name_parts)

        room_data[room_key] = {"Room_Name": full_name, "location": (0,0)}

    # Device data: DEVICE_ID as key
    for _, row in df.iterrows():
        device_id = row.get("DEVICE_ID", "")
        panel_id = row.get("LCP-01")
        keypad_name = row.get("KEYPAD_NAME", "")
        if pd.isna(device_id):
            if pd.isna(panel_id):
                continue
            else:
                pnlParts = panel_id.split("-")
                if pnlParts[0] == "SP":
                    if len(pnlParts) < 2 or not pnlParts[1].isdigit():
                        print(f"Skipping invalid device ID format: {panel_id}")
                        continue
                    else:
                        suffix = pnlParts[2]
                        if "&" in suffix:
                            letters = suffix.split("&")
                            for letter in letters:
                                entry = f"{pnlParts[0]}-{pnlParts[1]}-{letter}"
                                equipment_data[entry] = entry
                        else:
                            equipment_data[panel_id] = panel_id
                else:
                    print("Lighting Panel we can't insert ",panel_id)
            continue
        parts = device_id.split("-")
        if len(parts) < 2 or not parts[1].isdigit():
            print(f"Skipping invalid device ID format: {device_id}")
            continue
        # Differientiating device name and keypad name
        if parts[0] == "K":
            device_data[device_id] = keypad_name
        elif parts[0] == "R":
            equipment_data[device_id] = keypad_name
        elif parts[0] == "D":
            device_data[device_id] = keypad_name
        elif parts[0] == "MS":
            device_data[device_id] = device_id
        elif parts[0] == "OS":
            device_data[device_id] = device_id

        def get_room_number_from_id(identifier):
            match = re.search(r'-([0-9]+)-', identifier)
            return int(match.group(1)) if match else None

        # Check device_data
        for device_id in device_data:
            room_number = get_room_number_from_id(device_id)
            if room_number not in room_data:
                print(f"❌ Invalid device: {device_id} (No room {room_number})")
                sys.exit(1)

        # Check equipment_data
        for equip_id in equipment_data:
            room_number = get_room_number_from_id(equip_id)
            if room_number not in room_data:
                print(f"❌ Invalid equipment: {equip_id} (No room {room_number})")
                sys.exit(1)

        # Check zone_data
        for zone in zone_data:
            match = re.match(r'^(\d+)', zone)
            if match:
                room_number = int(match.group(1))
                if room_number not in room_data:
                    print(f"❌ Invalid zone: {zone} (No room {room_number})")
                    sys.exit(1)

def insert_rooms():
    '''inserting rooms into lutron'''
    x, y = keypadTargets["Place:"]
    pyautogui.moveTo(x, y)
    time.sleep(.1)
    pyautogui.click()
    if keypadTargets["Next"] == (0,0):
        # got to find next
        screenshot = ImageGrab.grab()
        data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
        for i in range(len(data['text'])):
            word = data['text'][i].strip()
            if not word:
                continue
            if "Next" in word:
                x = data['left'][i] + data['width'][i] // 2
                y = data['top'][i] + data['height'][i] // 2
                keypadTargets["Next"] = (x, y)
    time.sleep(.1)
    pyautogui.press('down')

    def enter_text(text):
        text = text.lower().title()
        pyperclip.copy(text)  # Copy to clipboard
        pyautogui.press('f2')  # Enter edit mode
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'v')  # Paste from clipboard
        pyautogui.press('enter')
        time.sleep(0.3)

    # Organize rooms by floor (1xx, 2xx, etc.)
    floors = {}
    for room_num, info in room_data.items():
        floor_key = str(room_num)[0]
        floors.setdefault(floor_key, []).append((room_num, info["Room_Name"]))

    time.sleep(1)

    floor_names = {
        '1': 'First Floor',
        '2': 'Second Floor',
        '3': 'Third Floor',
        '6': "Exterior"
    }

    sorted_floors = sorted(floors.keys())
    county = 0
    print("floors: ", sorted_floors)
    for i, floor_key in enumerate(sorted_floors):
        county += 1
        floor_name = floor_names.get(floor_key, f"Floor {floor_key}")
        # Enter floor name
        enter_text(floor_name)

        time.sleep(0.5)
        # Insert child room under floor
        count = 0
        for j, (room_num, room_name) in enumerate(sorted(floors[floor_key])):
            count += 1
            full_name = f"{room_name} {room_num}"
            if count == 1:
                pyautogui.press('insert')
            else:
                pyautogui.hotkey('ctrl', 'insert')
            enter_text(full_name)
        if i < len(sorted_floors) -1:
            findPreviousFloor(floor_name.split(" ")[0].strip())
            pyautogui.hotkey('ctrl', 'insert')
            time.sleep(0.5)

    print("Room insertion complete.")

def goToRoom(room_number):
    # Get the room name from the dictionary

    screenshot = ImageGrab.grab()

    data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)

    for i in range(len(data['text'])):
        word = data['text'][i].strip()
        if not word:
            continue
        if room_number in word:
            x = data['left'][i] + data['width'][i] // 2
            y = data['top'][i] + data['height'][i] // 2
            pyautogui.moveTo(x, y)
            time.sleep(.3)
            pyautogui.click()
            time.sleep(.3)

def getAllKeypadPoints():
    '''getting all points from lutron'''
    screenshot = ImageGrab.grab()
    data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
    found = {}

    for i in range(len(data['text']) - 1):
        w1, w2 = data['text'][i:i+2]
        phrase = f"{w1.strip()} {w2.strip()}"
        if phrase == "Wall Keypads":
            x = data['left'][i] + data['width'][i] // 2
            y = data['top'][i] + data['height'][i] // 2
            found["Wall Keypads"] = (x, y)

    for i in range(len(data['text'])):
        word = data['text'][i].strip()
        if not word:
            continue
        if "Sensors" in word:
            x = data['left'][i] + data['width'][i] // 2
            y = data['top'][i] + data['height'][i] // 2
            found["Sensors"] = (x, y)

    if found["Wall Keypads"] != (0,0):
        x, y = found["Wall Keypads"]
        pyautogui.moveTo(x, y)
        time.sleep(.1)
        pyautogui.click()
        time.sleep(.3)
        screenshot = ImageGrab.grab()
        data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
        for i in range(len(data['text'])):
            word = data['text'][i].strip()
            if not word:
                continue

            for label in keypadTargets.keys():
                if label in word:
                    x = data['left'][i] + data['width'][i] // 2
                    y = data['top'][i] + data['height'][i] // 2
                    if label == "seeTouch":
                        x -= 15
                        y -= 100
                    found[label] = (x, y)

    if found["Sensors"] != (0,0):
        x, y = found["Sensors"]
        pyautogui.moveTo(x, y)
        time.sleep(.1)
        pyautogui.click()
        time.sleep(.3)
        screenshot = ImageGrab.grab()
        data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)

        for i in range(len(data['text']) - 2):
            w1, w2, w3 = data['text'][i:i+3]
            phrase = f"{w1.strip()} {w2.strip()} {w3.strip()}"
            if phrase == "Ceiling Occ RF":
                x = data['left'][i] + data['width'][i] // 2
                y = data['top'][i] + data['height'][i] // 2
                x -= 15
                y -= 105
                found["Ceiling Occ RF"] = (x, y)

    # Update keypadTargets with found coordinates
    for label in keypadTargets:
        if label in found:
            keypadTargets[label] = found[label]

def searchForRoom(room_number):
    '''Check if given room number is visible on screen using OCR'''

    screenshot = ImageGrab.grab()
    data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)

    for word in data["text"]:
        word = word.strip()
        if word.isdigit() and int(word) == room_number:
            print(f"Room {room_number} found on screen.")
            return True

    print(f"Error: Room {room_number} not found on screen.")
    if room_number not in room_data:
        print(f"Error: Room {room_number} not found in dict.")
    return False

def insert_keypads():
    '''This is for entering all the keypads in'''
    x, y = keypadTargets["Place:"]
    pyautogui.moveTo(x, y)
    time.sleep(.1)
    pyautogui.click()
    time.sleep(.3)
    x, y = keypadTargets["Next"]
    if x == 0 or y == 0:
        # got to find next
        screenshot = ImageGrab.grab()
        data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
        for i in range(len(data['text'])):
            word = data['text'][i].strip()
            if not word:
                continue
            if "Next" in word:
                x = data['left'][i] + data['width'][i] // 2
                y = data['top'][i] + data['height'][i] // 2
                keypadTargets["Next"] = (x, y)
    pyautogui.moveTo(x, y)
    time.sleep(.1)
    pyautogui.click()
    time.sleep(.1)
    pyautogui.press('up')
    time.sleep(0.1)



    def enter_text(text):
        pyperclip.copy(text)  # Copy to clipboard
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'v')  # Paste from clipboard
        time.sleep(0.3)

    def get_current_room_number(room_number = 0):
        time.sleep(0.2)
        pyautogui.press('f2')
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.2)
        room_text = pyperclip.paste().strip()
        global last_checked_room
        if last_checked_room != room_number:
            last_checked_room = room_number
        else:
            goToRoom(str(room_number))
            return room_number
        if room_text:
            parts = room_text.split()
            if parts[-1].isdigit():
                return int(parts[-1])
        print("Error Getting Room Number")
        return -1

    def extract_room_number(key):
        parts = key.split("-")
        if len(parts) >= 2 and parts[1].isdigit():
            return int(parts[1])
        return float('inf')  # fallback so non-matching keys sort last
    current_room_number = get_current_room_number()
    pyautogui.press('enter')

    for device_id in sorted(device_data.keys(), key = extract_room_number):
        keypad_name = device_data[device_id]
        # print("device ", device_id, "keypad ", keypad_name)
        parts = device_id.split("-")
        if len(parts) < 2 or not parts[1].isdigit():
            print(f"Skipping invalid device ID format: {device_id}")
            continue

        if parts[0] == "D":
            print("Can't Enter Dimmers Right now ", device_id)
            continue

        room_number = int(parts[1])

        # Check if room exists
        if room_number not in room_data:
            print(f"Room {room_number} not found in room_data. Skipping {device_id}.")
            continue

        # Click on the keypad field (must already be in keypadTargets)
        if keypadTargets["seeTouch"] == (0,0):
            print("Error: 'keypad' target not defined in keypadTargets.")
            continue

        while current_room_number != room_number:
            x, y = keypadTargets["Next"]
            pyautogui.moveTo(x, y)
            time.sleep(.1)
            pyautogui.click()
            time.sleep(.1)
            pyautogui.press('up')
            time.sleep(0.1)
            current_room_number = get_current_room_number(room_number)
            pyautogui.press("enter")

        # Insert keypad
        if parts[0] == "K":
            x, y = keypadTargets["Wall Keypads"]
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            x, y = keypadTargets["seeTouch"]
            x += 35
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            x -= 35
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            # print("Inserting keypad:", keypad_name)
            enter_text(keypad_name)
            time.sleep(0.5)
            pyautogui.press('tab')
            time.sleep(0.5)
        elif parts[0] == "MS" or parts[0] == "OS":
            x, y = keypadTargets["Sensors"]
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            x, y = keypadTargets["Ceiling Occ RF"]
            x += 35
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            x -= 35
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)

        # print("inserting device ", device_id)
        enter_text(device_id)
        pyautogui.press('enter')
        time.sleep(0.5)

        last_room_number = room_number

    print("All keypads inserted.")

def getRoomLocations():
    '''getting room locations from lutron'''

    # # Take full screenshot
    screenshot = ImageGrab.grab()
    # Run OCR
    text_data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
    # print("OCR results: ", text_data)
    # Loop through OCR results and look for "Mech (102)"
    for i in range(len(text_data['text'])):
        word = text_data['text'][i].strip()
        if not word:
            continue

        # Try to extract a number from the word
        for room_key in room_data.keys():
            if str(room_key) in word:
                x = text_data['left'][i] + text_data['width'][i] // 2
                y = text_data['top'][i] + text_data['height'][i] // 2
                room_data[room_key]["location"] = (x, y)
                break  # Optional: stop after first match per word

def keypadChecker():
    '''Checks for duplicate keypad names within the same room'''
    # Step 1: Group keypads by room number
    room_keypads = {}
    for device_id, keypad_name in device_data.items():
        parts = device_id.split("-")
        if len(parts) < 2 or not parts[1].isdigit():
            print(f"Invalid device ID format: {device_id}")
            continue
        room_number = int(parts[1])
        room_keypads.setdefault(room_number, []).append((device_id, keypad_name))

    # Step 2: Check each room for duplicate names
    for room_number, keypads in room_keypads.items():
        name_counts = {}
        for device_id, name in keypads:
            if name not in name_counts:
                name_counts[name] = 1
                # Leave as-is
            else:
                count = name_counts[name]
                new_name = f"{name} {count}"
                print(f"Renaming duplicate in Room {room_number}: '{name}' → '{new_name}'")
                device_data[device_id] = new_name
                name_counts[name] += 1

def extract_room_number(zone):
    """Extracts the numeric room number from a zone like '100A'."""
    match = re.match(r'^(\d+)', zone)
    return int(match.group(1)) if match else None

def getAllLoadPoints():
    '''getting all points from lutron'''
    x, y = keypadTargets["controls"]
    pyautogui.moveTo(x, y)
    time.sleep(.1)
    pyautogui.press("down")
    time.sleep(.1)
    pyautogui.press("enter")
    time.sleep(.5)


    screenshot = ImageGrab.grab()
    data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)

    found = {}
    for i in range(len(data['text'])):
        word = data['text'][i].strip()
        if not word:
            continue
        for label in loadsTargets.keys():
            if label == word:
                x = data['left'][i] + data['width'][i] // 2
                y = data['top'][i] + data['height'][i] // 2
                found[label] = (x, y)

    # Update loadsTargets
    for label in loadsTargets:
        if label in found:
            loadsTargets[label] = found[label]

    if loadsTargets["Next"] == (0,0):
        x, y = loadsTargets["Place:"]
        pyautogui.moveTo(x, y)
        time.sleep(.1)
        pyautogui.click()
        time.sleep(.3)


        screenshot = ImageGrab.grab()

        data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)

        found = {}
        for i in range(len(data['text'])):
            word = data['text'][i].strip()
            if not word:
                continue
            for label in loadsTargets.keys():
                if label == word:
                    x = data['left'][i] + data['width'][i] // 2
                    y = data['top'][i] + data['height'][i] // 2
                    found[label] = (x, y)

        # Update loadsTargets
        for label in loadsTargets:
            if label in found:
                loadsTargets[label] = found[label]

def loadChecker():
    '''Checks to ensure that there are no duplicate zone names in same room'''
    room_loads = {}

    # Step 1: Group zone_names by room number
    for zone, data in zone_data.items():
        room_number = extract_room_number(zone)
        if room_number is None:
            continue

        room_loads.setdefault(room_number, []).append((zone, data))

    # Step 2: Check for duplicate zone_names (non-Ketra only)
    for room_number, loads in room_loads.items():
        name_count = {}
        for zone, data in loads:
            name = data["zone_name"]

            if name not in name_count:
                name_count[name] = 1
            else:
                new_name = f"{name} {name_count[name]}"
                print(f"Renaming duplicate in Room {room_number}: '{name}' → '{new_name}'")
                zone_data[zone]["zone_name"] = new_name
                name_count[name] += 1

def insertLoads():
    '''inserting all the loads into lutron'''
    time1 = .2
    time2 = .5
    # clicking the place
    x, y = loadsTargets["Place:"]
    pyautogui.moveTo(x, y)
    time.sleep(time1)
    pyautogui.click()
    time.sleep(time1)
    pyautogui.press("down")
    time.sleep(time1)
    def normalize(text):
        # Lowercase, remove hyphens and all non-alphanumeric except underscore
        return re.sub(r'[^a-z0-9_]', '', text.lower().replace("-", ""))

    def findingKetra(loadTag, retry = 0):
        time.sleep(1)
        screenshot = ImageGrab.grab()

        data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
        x = y = 0
        words = data['text']
        norm_tag = normalize(loadTag)
        for i in range(len(words)):
            for j in range(1, 4):  # try 1- to 3-word combinations
                phrase = " ".join(words[i:i + j])
                if normalize(phrase) == norm_tag:
                    x = data['left'][i] + data['width'][i] // 2
                    y = data['top'][i] + data['height'][i] // 2
                    return x, y

        print("❌ Ketra Load not found:", loadTag)
        if retry < 3:
            return findingKetra(loadTag, retry + 1)
        return 0, 0

    def enter_text(text):
        pyperclip.copy(text)  # Copy to clipboard
        pyautogui.press("shift")
        pyautogui.press("shift")
        time.sleep(time2)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(time2)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(time2)
        pyautogui.press("tab")
        time.sleep(time2)

    def enterFixture(fixture):
        """entering in fixture"""
        pyautogui.write(fixture, interval=0.1)
        time.sleep(time2)
        pyautogui.press("tab")

    def get_current_room_number(room_number = 0):
        time.sleep(time2)
        pyautogui.press('f2')
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.2)
        room_text = pyperclip.paste().strip()
        global last_checked_room
        if last_checked_room != room_text:
            last_checked_room = room_text
        else:
            goToRoom(str(room_number))
            return room_number
        if room_text:
            parts = room_text.split()
            if parts[-1].isdigit():
                return int(parts[-1])
        print("Error Getting Room Number")
        return -1

    def extract_leading_number(value):
        match = re.match(r'^(\d+)', value)
        return int(match.group(1)) if match else None

    current_room_number = get_current_room_number()
    pyautogui.press('enter')
    if loadsTargets["Add"] == (0,0):
        print("Error: 'add' target not defined in LoadTargets.")
        return

    for load_num in sorted(zone_data.keys()):
        room_number = extract_leading_number(load_num)
        if room_number == 110:
            time1 = .3
            time2 = .7
        elif room_number == -1:
            x, y = loadsTargets["Next"]
            if x == 0 or y == 0:
                x,y = keypadTargets["Next"]
                if x != 0 or y != 0:
                    loadsTargets["Next"] = (x,y)
            pyautogui.moveTo(x, y)
            time.sleep(time1)
            pyautogui.click()
            time.sleep(time1)
            pyautogui.press('up')
            time.sleep(time1)
            current_room_number = get_current_room_number(room_number)
            pyautogui.press('enter')

        while room_number != current_room_number:
            # checkingLoads(room_number)
            x, y = loadsTargets["Next"]
            if x == 0 or y == 0:
                x,y = keypadTargets["Next"]
                if x != 0 or y != 0:
                    loadsTargets["Next"] = (x,y)
            pyautogui.moveTo(x, y)
            time.sleep(time1)
            pyautogui.click()
            time.sleep(time1)
            pyautogui.press('up')
            time.sleep(time1)
            current_room_number = get_current_room_number(room_number)
            pyautogui.press('enter')


        x, y = loadsTargets["Add"]
        pyautogui.moveTo(x, y)
        time.sleep(time2)
        pyautogui.click()
        time.sleep(time2)
        if zone_data[load_num]["Ketra"]:
            '''if light is ketra'''
            pyautogui.press("left")
            enter_text(zone_data[load_num]["zone_name"])
            pyautogui.press("right")
            time.sleep(.2)
            pyautogui.press("right")
            time.sleep(.2)
            loadTag = load_num + "-1"
            enter_text(loadTag)
            pyautogui.press("left")
            time.sleep(.2)
            pyautogui.press("left")
            time.sleep(.2)
            pyautogui.press("left")
            time.sleep(.2)
            enterFixture(zone_data[load_num]["fixtures"][0]["fixture"])
            time.sleep(.2)
            pyautogui.press("enter")
            time.sleep(.2)
            #getting ketra loadNum
            x,y = (0,0)
            x,y = findingKetra(zone_data[load_num]["zone_name"])
            if x == 0 or y == 0:
                continue
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.2)
            pyautogui.press("enter")
            time.sleep(.2)
            # This is where I am and I need to enter the rest of the ketra fixtures here
            tagCounter = 2
            current_fixture = 1
            for fixture_entry in zone_data[load_num]["fixtures"]:
                fixture = fixture_entry["fixture"]
                count = fixture_entry["count"]
                if current_fixture == 1:
                    count -= 1
                    current_fixture += 1
                for numOfEachFixtures in range(count):
                    enter_text(zone_data[load_num]["zone_name"])
                    time.sleep(.5)
                    enterFixture(fixture)
                    time.sleep(.2)
                    pyautogui.press("right")
                    time.sleep(.2)
                    loadTag = (f"{load_num}-{tagCounter}")
                    enter_text(loadTag)
                    time.sleep(.2)
                    pyautogui.press("left")
                    time.sleep(.2)
                    pyautogui.press("left")
                    time.sleep(.2)
                    pyautogui.press("left")
                    time.sleep(.2)
                    pyautogui.press("left")
                    time.sleep(.2)
                    pyautogui.press("enter")
                    time.sleep(.2)
                    tagCounter += 1


        else:
            pyautogui.press("left")
            enter_text(zone_data[load_num]["zone_name"])
            enterFixture(zone_data[load_num]["fixtures"][0]["fixture"])
            enter_text(zone_data[load_num]["fixtures"][0]["count"])
            enter_text(load_num)
            if zone_data[load_num]["fixtures"][0]["fixture"] != "EH" and zone_data[load_num]["fixtures"][0]["fixture"] != "Fan":
                time.sleep(time1)
                pyautogui.write("y")
                pyautogui.press("enter")

    print("All loads inserted.")

def insertFixtureTypes():
    x, y = loadsTargets["Edit Fixture Types"]
    print("edit fixture type X and Y ", x, " ", y)
    pyautogui.moveTo(x, y)
    time.sleep(.1)
    pyautogui.click()
    time.sleep(.5)

    screenshot = ImageGrab.grab()
    # Run OCR
    found ={}
    data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
    # Loop through OCR results and look for "Mech (102)"
    for i in range(len(data['text']) - 2):
        w1, w2, w3 = data['text'][i:i+3]
        phrase = f"{w1.strip()} {w2.strip()} {w3.strip()}"
        if phrase == "Lutron Lights":
            x = data['left'][i] + data['width'][i] // 2
            y = data['top'][i] + data['height'][i] // 2
            found["Lutron Lights"] = (x, y)

    # Second pass for single-word labels
    for i in range(len(data['text'])):
        word = data['text'][i].strip()
        if not word:
            continue
        for label in fixtureTargets.keys():
            if label == word:
                x = data['left'][i] + data['width'][i] // 2
                y = data['top'][i] + data['height'][i] // 2
                found[label] = (x, y)

    # Update loadsTargets
    for label in fixtureTargets:
        if label in found:
            fixtureTargets[label] = found[label]

    # can just start typing

def compareSheets(ketraLights):
    '''This is used to compare the difference between two sheets to get the changes'''

    #These are the old values
    room_data_old = room_data.copy()
    zone_data_old = zone_data.copy()
    device_data_old = device_data.copy()
    repeater_data_old = equipment_data.copy()

    # Prompt user to select the new Excel file
    file_path = input("Enter path to new Excel file: ").strip().strip('"')


    if not file_path:
        print("No file selected. Exiting.")
        return

    #Loading in new excel file
    load_excel_file(file_path,ketraLights)

    def compare_dicts(old_dict, new_dict):
        old_keys = set(old_dict.keys())
        new_keys = set(new_dict.keys())

        added = new_keys - old_keys
        removed = old_keys - new_keys
        changed = [key for key in old_keys & new_keys if old_dict[key] != new_dict[key]]

        return added, removed, changed

    added_rooms, removed_rooms, changed_rooms = compare_dicts(room_data_old, room_data)
    added_zones, removed_zones, changed_zones = compare_dicts(zone_data_old, zone_data)
    added_keypads, removed_keypads, changed_keypads = compare_dicts(device_data_old, device_data)
    added_repeaters, removed_repeaters, changed_repeaters = compare_dicts(repeater_data_old, repeater_data)

    # --- Output results ---

    print("\n--- 🧾 Sheet Comparison Result ---")

    def print_section(title, keys, old_dict=None, new_dict=None):
        if not keys:
            return
        print(f"\n{title}")
        for key in sorted(keys):
            if old_dict and new_dict:
                print(f"  ~ {key}:\n    Old: {old_dict[key]}\n    New: {new_dict[key]}")
            elif new_dict:
                print(f"  + {key}: {new_dict[key]}")
            elif old_dict:
                print(f"  - {key}: {old_dict[key]}")

    print_section("🟢 New Rooms:", added_rooms, new_dict=room_data)
    print_section("🟢 New Zones:", added_zones, new_dict=zone_data)
    print_section("🟢 New Keypads:", added_keypads, new_dict=device_data)
    print_section("🟢 New Repeaters:", added_repeaters, new_dict=equipment_data)

    print_section("🔴 Removed Rooms:", removed_rooms, old_dict=room_data_old)
    print_section("🔴 Removed Zones:", removed_zones, old_dict=zone_data_old)
    print_section("🔴 Removed Keypads:", removed_keypads, old_dict=device_data_old)
    print_section("🔴 Removed Repeaters:", removed_repeaters, old_dict=repeater_data_old)

    print_section("🟡 Modified Rooms:", changed_rooms, old_dict=room_data_old, new_dict=room_data)
    print_section("🟡 Modified Zones:", changed_zones, old_dict=zone_data_old, new_dict=zone_data)
    print_section("🟡 Modified Keypads:", changed_keypads, old_dict=device_data_old, new_dict=device_data)
    print_section("🟡 Modified Repeaters:", changed_repeaters, old_dict=repeater_data_old, new_dict=equipment_data)

    if not any([
        added_rooms, removed_rooms, changed_rooms,
        added_zones, removed_zones, changed_zones,
        added_keypads, removed_keypads, changed_keypads,
        added_repeaters, removed_repeaters, changed_repeaters
    ]):
        print("✅ No differences found between the sheets.")

def updateLutron(new):
    '''updating Lutron with the new values'''

def getAllShadePoints():
    '''getting equipment points'''

    screenshot = ImageGrab.grab()
    data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
    found = {}

    for i in range(len(data['text']) - 2):
        w1, w2, w3 = data['text'][i:i+3]
        phrase = f"{w1.strip()} {w2.strip()} {w3.strip()}"
        if phrase == "Shade Group 1":
            x = data['left'][i] + data['width'][i] // 2
            y = data['top'][i] + data['height'][i] // 2
            found["Shade Group 1"] = (x, y)

        if phrase == "Add shade group":
            x = data['left'][i] + data['width'][i] // 2
            y = data['top'][i] + data['height'][i] // 2
            found["Add shade group"] = (x, y)
            x, y = shadeTargets["Add shade group"]

    for i in range(len(data['text'])):
        word = data['text'][i].strip()
        if not word:
            continue

        for label in shadeTargets.keys():
            if label in word:
                x = data['left'][i] + data['width'][i] // 2
                y = data['top'][i] + data['height'][i] // 2
                found[label] = (x, y)

    # Update keypadTargets with found coordinates
    for label in shadeTargets:
        if label in found:
            shadeTargets[label] = found[label]

def insertShades():
    '''inserting the repeaters into the equipments page'''
    time.sleep(1)
    x, y = keypadTargets["controls"]
    pyautogui.moveTo(x, y)
    time.sleep(.1)
    pyautogui.press("down")
    time.sleep(.1)
    pyautogui.press("enter")
    time.sleep(.5)
    getAllShadePoints()
    x, y = shadeTargets["Place:"]
    pyautogui.moveTo(x, y)
    time.sleep(.2)
    pyautogui.click()
    time.sleep(.2)
    x, y = shadeTargets["Next"]
    if x == 0 or y == 0:
        x,y = keypadTargets["Next"]
        if x != 0 or y != 0:
            shadeTargets["Next"] = (x,y)
    pyautogui.moveTo(x, y)
    time.sleep(.1)
    pyautogui.click()
    time.sleep(.1)
    pyautogui.press("up")
    time.sleep(.1)

    def enter_text(text):
        pyperclip.copy(text)  # Copy to clipboard
        pyautogui.press("shift")
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'v')  # Paste from clipboard
        time.sleep(0.3)
        pyautogui.press("enter")
        time.sleep(0.3)

    def get_current_room_number(room_number = 0):
        time.sleep(1)
        pyautogui.press('f2')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        room_text = pyperclip.paste().strip()
        global last_checked_room
        if last_checked_room != room_text:
            last_checked_room = room_text
        else:
            goToRoom(str(room_number))
            return room_number
        if room_text:
            parts = room_text.split()
            if parts[-1].isdigit():
                return int(parts[-1])
        return None

    current_room_number = get_current_room_number()
    pyautogui.press('enter')

    only_A_shades = []
    other_shades = []
    for shade_id in shade_data:
        parts = shade_id.split("-")
        last_part = parts[2]

        # Check if last part is a single letter A–Z
        if len(last_part) == 1 and last_part.upper() == "A":
            only_A_shades.append(shade_id)
        else:
            other_shades.append(shade_id)
    for shade_id in sorted(only_A_shades):
        parts = shade_id.split("-")
        if len(parts) < 2 or not parts[1].isdigit():
            print(f"Skipping invalid shade ID format: {shade_id}")
            continue

        room_number = int(parts[1])

        while room_number != current_room_number:
            x, y = shadeTargets["Next"]
            pyautogui.moveTo(x, y)
            time.sleep(.2)
            pyautogui.click()
            time.sleep(.2)
            pyautogui.press('up')
            time.sleep(.2)
            current_room_number = get_current_room_number(room_number)
            pyautogui.press('enter')

        x, y = shadeTargets["Add shade group"]
        pyautogui.moveTo(x, y)
        time.sleep(.2)
        pyautogui.click()
        for shades in sorted(other_shades):
            cur_room = shades.split("-")
            if int(cur_room[1]) == room_number:
                time.sleep(.2)
                pyautogui.click()
        time.sleep(1)
        getAllShadePoints()
        x, y = shadeTargets["Shade Group 1"]
        pyautogui.moveTo(x, y)
        time.sleep(.2)
        pyautogui.click()
        time.sleep(.2)
        enter_text(shade_id)
        for shades in sorted(other_shades):
            cur_room = shades.split("-")
            if int(cur_room[1]) == room_number:
                enter_text(shades)

    print("All Shades inserted.")

def gettingAllEquipmentPoints():
    '''getting equipment points'''
    time.sleep(1)
    x, y = keypadTargets["controls"]
    pyautogui.moveTo(x, y)
    time.sleep(.1)
    pyautogui.press("down")
    time.sleep(.1)
    pyautogui.press("enter")
    time.sleep(.5)


    screenshot = ImageGrab.grab()
    data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
    found = {}

    for i in range(len(data['text'])):
        word = data['text'][i].strip()
        if not word:
            continue
        for label in equipmentTargets.keys():
            if label == word:
                x = data['left'][i] + data['width'][i] // 2
                y = data['top'][i] + data['height'][i] // 2
                found[label] = (x, y)

    if found["Panels"] != (0,0):
        x,y = found["Panels"]
        pyautogui.moveTo(x, y)
        time.sleep(1)
        pyautogui.click()
        time.sleep(1)
        screenshot = ImageGrab.grab()
        data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
        for i in range(len(data['text']) - 1):
            w1, w2 = data['text'][i:i+2]
            phrase = f"{w1.strip()} {w2.strip()}"
            if phrase == "10 w/TBs":
                x = data['left'][i] + data['width'][i] // 2
                y = data['top'][i] + data['height'][i] // 2
                x -= 35
                y -= 115
                found["10 w/TBs"] = (x, y)

            if phrase == "8 w/TBs":
                x = data['left'][i] + data['width'][i] // 2
                y = data['top'][i] + data['height'][i] // 2
                x -= 35
                y -= 115
                found["8 w/TBs"] = (x, y)
        # Second pass for single-word labels
        for i in range(len(data['text'])):
            word = data['text'][i].strip()
            if not word:
                continue
            for label in equipmentTargets.keys():
                if label == word:
                    x = data['left'][i] + data['width'][i] // 2
                    y = data['top'][i] + data['height'][i] // 2
                    if label == "Smart":
                        x -= 35
                        y -= 115
                    found[label] = (x, y)
    # print("found: ", found)
    if found["Devices"] != (0,0):
        x,y = found["Devices"]
        pyautogui.moveTo(x, y)
        time.sleep(1)
        pyautogui.click()
        time.sleep(1)
        screenshot = ImageGrab.grab()

        data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
        for i in range(len(data['text'])):
            word = data['text'][i].strip()
            if not word:
                continue
            for label in equipmentTargets.keys():
                if label == word:
                    x = data['left'][i] + data['width'][i] // 2
                    y = data['top'][i] + data['height'][i] // 2
                    if label == "Hybrid":
                        x -= 15
                        y -= 100
                    if label == "Clear":
                        x -= 35
                        y -= 115
                    found[label] = (x, y)


    # Update loadsTargets
    for label in equipmentTargets:
        if label in found:
            equipmentTargets[label] = found[label]

def insertEquipment():
    '''inserting the repeaters into the equipments page'''
    x, y = equipmentTargets["Place:"]
    pyautogui.moveTo(x, y)
    time.sleep(.2)
    pyautogui.click()
    time.sleep(.2)
    pyautogui.press("down")
    time.sleep(.2)

    def enter_text(text):
        pyperclip.copy(text)  # Copy to clipboard
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'v')  # Paste from clipboard
        time.sleep(0.3)

    def get_current_room_number(room_number = 0):
        time.sleep(1)
        pyautogui.press('f2')
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.2)
        room_text = pyperclip.paste().strip()
        global last_checked_room
        if last_checked_room != room_text:
            last_checked_room = room_text
        else:
            goToRoom(str(room_number))
            return room_number
        if room_text:
            parts = room_text.split()
            if parts[-1].isdigit():
                return int(parts[-1])
        return None

    def extract_room_number(key):
        parts = key.split("-")
        if len(parts) >= 2 and parts[1].isdigit():
            return int(parts[1])
        return float('inf')  # fallback so non-matching keys sort last

    current_room_number = get_current_room_number()
    pyautogui.press('enter')
    for device_id in sorted(equipment_data.keys(), key = extract_room_number):
        parts = device_id.split("-")
        if len(parts) < 2 or not parts[1].isdigit():
            print(f"Skipping invalid device ID format: {device_id}")
            continue

        room_number = int(parts[1])

        # Click on the keypad field (must already be in keypadTargets)
        if equipmentTargets["Hybrid"] == (0,0):
            print("Error: 'Repeater' target not found.")
            continue

        if room_number not in room_data:
            print("Room for this device does not exist ", device_id)
            continue

        if current_room_number not in room_data:
            x, y = equipmentTargets["Next"]
            if x == 0 or y == 0:
                x,y = keypadTargets["Next"]
                if x != 0 or y != 0:
                    shadeTargets["Next"] = (x,y)
            pyautogui.moveTo(x, y)
            time.sleep(.2)
            pyautogui.click()
            time.sleep(.2)
            pyautogui.press('up')
            time.sleep(.2)
            current_room_number = get_current_room_number(room_number)
            pyautogui.press('enter')

        while room_number != current_room_number:
            # checkingLoads(room_number)
            x, y = equipmentTargets["Next"]
            if x == 0 or y == 0:
                x,y = keypadTargets["Next"]
                if x != 0 or y != 0:
                    shadeTargets["Next"] = (x,y)
            pyautogui.moveTo(x, y)
            time.sleep(.2)
            pyautogui.click()
            time.sleep(.2)
            pyautogui.press('up')
            time.sleep(.2)
            current_room_number = get_current_room_number(room_number)
            pyautogui.press('enter')

        # Insert equipment
        if parts[0] == "R":
            x, y = equipmentTargets["Devices"]
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            x, y = equipmentTargets["Hybrid"]
            x += 35
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            x -= 35
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
        elif parts[0] == "GW":
            x, y = equipmentTargets["Devices"]
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            x, y = equipmentTargets["Clear"]
            x += 35
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            x -= 35
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
        elif parts[0] == "SP":
            x, y = equipmentTargets["Panels"]
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            x, y = equipmentTargets["Smart"]
            x += 35
            pyautogui.moveTo(x, y)
            pyautogui.click()
            time.sleep(.3)
            x -= 35
            pyautogui.moveTo(x, y)
            time.sleep(.3)
            pyautogui.click()
            time.sleep(.5)
            device_id = "Power Supply " + device_id

        enter_text(device_id)
        pyautogui.press('enter')
        time.sleep(0.5)

    print("All equipment inserted.")

def prompt_file_selection():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title="Select a file", filetypes=[("Excel Files", "*.xlsx *.xls")])
    return file_path

if __name__ == "__main__":
    # Example usage
    ketraLights = ["AK", "AL"]
    print("Select the original xls")
    # input()
    file_path = prompt_file_selection()
    if file_path:
        print("Selected file:", file_path)
    else:
        print("No file selected.")

    load_excel_file(file_path, ketraLights)

    print("Press '1' if you are creating program.")
    print("Press '2' if you are editing existing program.")
    choice = input("Enter your choice: ")
    if choice == '1':
        '''This is for creating a new program'''
        time.sleep(2)
        getAllKeypadPoints()
        insert_rooms()
        getRoomLocations()
        keypadChecker()
        insert_keypads()
        loadChecker()
        getAllLoadPoints()
        insertLoads()
        insertShades()
        gettingAllEquipmentPoints()
        insertEquipment()


    elif choice == '2':
        '''This is for editing existing program'''
        compareSheets(ketraLights)
        updateLutron()
    else:
        print("Invalid choice. Exiting.")
        sys.exit(1)
