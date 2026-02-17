"""
CES Touch Product Export - GUI Automation
Uses PyAutoGUI image matching with coordinate fallback.
Screen resolution: 1440x900

Flow:
1. Kill CES Touch if running
2. Launch CES Touch, wait for login screen
3. Click ADMIN -> enter password 2448 -> Accept
4. MANAGER SCREEN -> BACK OFFICE -> Utilities -> System Menu
5. Import / Export -> Products -> Select directory -> wait for export
6. Exit dialog -> Sales / Review -> Sales Mode (back to login)

SETUP:
  Place button images in C:\BulkPricing\btn_images\
  pip install pyautogui Pillow pywin32
"""

import pyautogui
import subprocess
import time
import os
import sys

# Safety settings
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

# ============================================================================
# CONFIGURATION
# ============================================================================

CES_TOUCH_EXE = r"C:\touch\touch.exe"
CES_TOUCH_DIR = r"C:\Touch"
CES_PROCESS_NAME = "touch.exe"
ADMIN_PASSWORD = "2448"
STARTUP_WAIT = 60
EXPORT_WAIT = 120

# Button images directory
BTN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "btn_images")

# Image matching confidence (lower = more lenient)
CONFIDENCE = 0.7

# Fallback coordinates (1440x900) - confirmed from actual screen
FALLBACK = {
    'btn_admin':          (499, 100),
    'btn_manager_screen': (72, 745),
    'btn_back_office':    (158, 185),
    'btn_utilities':      (1321, 67),
    'btn_system_menu':    (1322, 198),
    'btn_import_export':  (794, 444),
    'btn_products':       (721, 700),
    'btn_select':         (841, 344),
    'btn_exit_dialog':    (1216, 819),
    'btn_sales_review':   (106, 67),
    'btn_sales_mode':     (103, 195),
}

# Check if OpenCV is available for confidence matching
try:
    import cv2
    HAS_OPENCV = True
except ImportError:
    HAS_OPENCV = False

# ============================================================================

def log(msg):
    timestamp = time.strftime('%H:%M:%S')
    print(f"  [{timestamp}] {msg}")


def locate_image(img_path):
    """Locate image on screen, with or without OpenCV confidence"""
    if HAS_OPENCV:
        return pyautogui.locateCenterOnScreen(img_path, confidence=CONFIDENCE)
    else:
        return pyautogui.locateCenterOnScreen(img_path)


def find_and_click(name, timeout=10, pause_after=1.5):
    """Try to find button by image, fall back to coordinates"""
    img_path = os.path.join(BTN_DIR, f"{name}.png")
    
    # Try image matching first
    if os.path.exists(img_path):
        log(f"Looking for {name}...")
        start = time.time()
        while time.time() - start < timeout:
            try:
                location = locate_image(img_path)
                if location:
                    log(f"Found {name} at ({location.x}, {location.y}) - clicking")
                    pyautogui.click(location)
                    time.sleep(pause_after)
                    return True
            except pyautogui.ImageNotFoundException:
                pass
            except Exception as e:
                log(f"Image search error: {e}")
                break
            time.sleep(0.5)
        log(f"Image match failed for {name}, trying fallback...")
    
    # Fallback to coordinates
    if name in FALLBACK and FALLBACK[name] is not None:
        x, y = FALLBACK[name]
        log(f"Using fallback coords for {name}: ({x}, {y})")
        pyautogui.click(x, y)
        time.sleep(pause_after)
        return True
    
    log(f"[ERROR] Cannot find {name} - no image match and no fallback coords")
    return False


def image_gone(name, timeout=120):
    """Wait for an image to disappear from screen"""
    img_path = os.path.join(BTN_DIR, f"{name}.png")
    if not os.path.exists(img_path):
        time.sleep(30)
        return True
    
    start = time.time()
    while time.time() - start < timeout:
        try:
            if HAS_OPENCV:
                location = pyautogui.locateOnScreen(img_path, confidence=CONFIDENCE)
            else:
                location = pyautogui.locateOnScreen(img_path)
            if location is None:
                return True
        except pyautogui.ImageNotFoundException:
            return True
        except:
            pass
        time.sleep(2)
    return False


def kill_ces_touch():
    """Force close CES Touch if it's running"""
    log("Checking if CES Touch is running...")
    try:
        result = subprocess.run(
            ['tasklist', '/FI', f'IMAGENAME eq {CES_PROCESS_NAME}'],
            capture_output=True, text=True, timeout=10
        )
        if CES_PROCESS_NAME.lower() in result.stdout.lower():
            log("CES Touch is running - force closing...")
            subprocess.run(
                ['taskkill', '/F', '/IM', CES_PROCESS_NAME],
                capture_output=True, timeout=10
            )
            time.sleep(3)
            log("CES Touch closed")
        else:
            log("CES Touch not running")
    except Exception as e:
        log(f"[WARN] Could not check/kill process: {e}")


def launch_ces_touch():
    """Launch CES Touch and wait for login screen"""
    log(f"Launching CES Touch: {CES_TOUCH_EXE}")
    
    if not os.path.exists(CES_TOUCH_EXE):
        log(f"[ERROR] CES Touch not found at: {CES_TOUCH_EXE}")
        return False
    
    subprocess.Popen([CES_TOUCH_EXE], cwd=CES_TOUCH_DIR)
    
    # Wait for ADMIN button to appear via image match
    log("Waiting for CES Touch login screen...")
    admin_img = os.path.join(BTN_DIR, "btn_admin.png")
    if os.path.exists(admin_img):
        found = False
        for i in range(STARTUP_WAIT):
            try:
                loc = locate_image(admin_img)
                if loc:
                    log(f"Login screen detected after {i+1}s")
                    found = True
                    time.sleep(2)
                    break
            except:
                pass
            time.sleep(1)
        if not found:
            log(f"Login screen not detected after {STARTUP_WAIT}s, proceeding anyway...")
    else:
        log(f"No btn_admin image - waiting fixed {STARTUP_WAIT}s...")
        time.sleep(STARTUP_WAIT)
    
    return True


def enter_password():
    """Type the admin password using keyboard"""
    log("Entering password via keyboard...")
    time.sleep(0.5)
    pyautogui.typewrite(ADMIN_PASSWORD, interval=0.15)
    time.sleep(0.5)
    
    # Try clicking Accept button, fall back to Enter key
    if not find_and_click('btn_accept', timeout=3, pause_after=2.0):
        log("Trying Enter key instead...")
        pyautogui.press('enter')
        time.sleep(2.0)


def wait_for_export():
    """Wait for the export to complete"""
    export_file = r"C:\Touch\IMP-EXP\sku_0002.xls"
    
    log("Waiting for export to complete...")
    
    old_mtime = 0
    if os.path.exists(export_file):
        old_mtime = os.path.getmtime(export_file)
    
    # Watch for "Exporting" message to disappear
    txt_img = os.path.join(BTN_DIR, "txt_exporting.png")
    if os.path.exists(txt_img):
        log("Watching for export message to disappear...")
        time.sleep(2)
        image_gone('txt_exporting', timeout=EXPORT_WAIT)
        time.sleep(2)
    
    # Check file was updated
    start = time.time()
    while time.time() - start < 30:
        if os.path.exists(export_file):
            new_mtime = os.path.getmtime(export_file)
            if new_mtime > old_mtime:
                time.sleep(3)
                log(f"Export complete! File updated.")
                return True
        time.sleep(2)
    
    if os.path.exists(export_file):
        log("[WARN] File exists but may not have been updated")
        return True
    
    log("[ERROR] Export file not found")
    return False


def export_products():
    """Run the full GUI automation sequence"""
    
    print("\n" + "=" * 60)
    print("  CES TOUCH PRODUCT EXPORT - GUI AUTOMATION")
    print("=" * 60)
    
    # Check for button images
    if os.path.exists(BTN_DIR):
        images = [f for f in os.listdir(BTN_DIR) if f.endswith('.png')]
        log(f"Found {len(images)} button images in {BTN_DIR}")
    else:
        log(f"[WARN] No btn_images folder at {BTN_DIR}")
        os.makedirs(BTN_DIR, exist_ok=True)
    
    # Step 0: Kill and relaunch
    kill_ces_touch()
    
    if not launch_ces_touch():
        return False
    
    try:
        # Step 1: Click ADMIN
        log("=== Step 1: Login as ADMIN ===")
        if not find_and_click('btn_admin', timeout=15, pause_after=2.0):
            return False
        
        # Step 2: Enter password
        log("=== Step 2: Enter password ===")
        enter_password()
        
        # Step 3: MANAGER SCREEN
        log("=== Step 3: Go to Manager Screen ===")
        if not find_and_click('btn_manager_screen', timeout=10, pause_after=2.0):
            return False
        
        # Step 4: BACK OFFICE
        log("=== Step 4: Go to Back Office ===")
        if not find_and_click('btn_back_office', timeout=10, pause_after=2.0):
            return False
        
        # Step 5: Utilities
        log("=== Step 5: Open Utilities ===")
        if not find_and_click('btn_utilities', timeout=10, pause_after=2.0):
            return False
        
        # Step 6: System Menu
        log("=== Step 6: Open System Menu ===")
        if not find_and_click('btn_system_menu', timeout=10, pause_after=2.0):
            return False
        
        # Step 7: Import / Export
        log("=== Step 7: Open Import / Export ===")
        if not find_and_click('btn_import_export', timeout=10, pause_after=2.5):
            return False
        
        # Step 8: Products
        log("=== Step 8: Click Products to export ===")
        if not find_and_click('btn_products', timeout=10, pause_after=2.5):
            return False
        
        # Step 9: Select directory
        log("=== Step 9: Confirm export directory ===")
        if not find_and_click('btn_select', timeout=10, pause_after=2.0):
            return False
        
        # Step 10: Wait for export
        log("=== Step 10: Waiting for export ===")
        if not wait_for_export():
            log("[WARN] Export may have failed, continuing...")
        
        # Step 11: Exit dialog
        log("=== Step 11: Close Import/Export dialog ===")
        if not find_and_click('btn_exit_dialog', timeout=10, pause_after=2.0):
            return False
        
        # Step 12: Sales / Review
        log("=== Step 12: Navigate to Sales / Review ===")
        if not find_and_click('btn_sales_review', timeout=10, pause_after=2.0):
            return False
        
        # Step 13: Sales Mode
        log("=== Step 13: Return to login screen ===")
        if not find_and_click('btn_sales_mode', timeout=10, pause_after=1.5):
            return False
        
        log("[DONE] Export complete - CES Touch is back at login screen")
        return True
        
    except pyautogui.FailSafeException:
        log("[ABORT] Mouse moved to corner - failsafe triggered!")
        return False
    except Exception as e:
        log(f"[ERROR] Automation failed: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = export_products()
    sys.exit(0 if success else 1)
