"""
Button Image Capture Helper
Run this on the Windows PC to capture button screenshots for PyAutoGUI image matching.

Instructions:
1. Navigate CES Touch to the screen with the button you want to capture
2. Run: python capture_buttons.py
3. It will take a full screenshot and save it
4. Then you can crop buttons from it, OR use the interactive mode
"""

import pyautogui
import os
import sys
import time

BTN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "btn_images")
os.makedirs(BTN_DIR, exist_ok=True)

def capture_full():
    """Take a full screenshot"""
    fname = os.path.join(BTN_DIR, "full_screenshot.png")
    print(f"Taking screenshot in 3 seconds...")
    time.sleep(3)
    img = pyautogui.screenshot()
    img.save(fname)
    print(f"Saved: {fname} ({img.size[0]}x{img.size[1]})")
    return img

def capture_region(name):
    """Interactively capture a button region"""
    print(f"\nCapturing: {name}")
    print("Move your mouse to the TOP-LEFT corner of the button and press Enter...")
    input()
    x1, y1 = pyautogui.position()
    print(f"  Top-left: ({x1}, {y1})")
    
    print("Now move to the BOTTOM-RIGHT corner and press Enter...")
    input()
    x2, y2 = pyautogui.position()
    print(f"  Bottom-right: ({x2}, {y2})")
    
    # Take screenshot and crop
    img = pyautogui.screenshot(region=(x1, y1, x2-x1, y2-y1))
    fname = os.path.join(BTN_DIR, f"{name}.png")
    img.save(fname)
    print(f"  Saved: {fname} ({x2-x1}x{y2-y1} pixels)")

def capture_all():
    """Guide through capturing all buttons"""
    buttons = [
        ('btn_admin', 'Login screen: ADMIN button (orange)'),
        ('btn_accept', 'Password dialog: Accept button (green)'),
        ('txt_enter_password', 'Password dialog: "Enter Password" text'),
        ('btn_manager_screen', 'Sales screen: MANAGER SCREEN button'),
        ('btn_back_office', 'Manager screen: BACK OFFICE button'),
        ('btn_utilities', 'Back office: Utilities button (green)'),
        ('btn_system_menu', 'Utilities: System Menu button (green)'),
        ('btn_import_export', 'System menu: Import / Export button'),
        ('btn_products', 'Export dialog: Products button'),
        ('btn_select', 'Directory dialog: Select button'),
        ('txt_exporting', 'Export: "Exporting Products file..." text'),
        ('btn_exit_dialog', 'Import/Export dialog: Exit button (red)'),
        ('btn_sales_review', 'Top nav: Sales / Review button'),
        ('btn_sales_mode', 'Sales submenu: Sales Mode button'),
    ]
    
    print("=" * 60)
    print("  BUTTON IMAGE CAPTURE")
    print("=" * 60)
    print(f"\nThis will capture {len(buttons)} button images.")
    print("For each button, navigate CES Touch to the right screen,")
    print("then position your mouse at the corners when prompted.\n")
    
    for name, desc in buttons:
        print(f"\n--- {desc} ---")
        inp = input("Ready? (y to capture, s to skip, q to quit): ").strip().lower()
        if inp == 'q':
            break
        elif inp == 's':
            continue
        capture_region(name)
    
    print(f"\nDone! Images saved to: {BTN_DIR}")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == 'full':
        capture_full()
    else:
        capture_all()
