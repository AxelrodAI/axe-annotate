"""
Diagnostic script to test Excel behavior when alt-tabbing between applications.
This tests the scenario where hotkeys stop working after switching to another app and back.

Usage:
    python diagnose_alt_tab.py --auto
"""
import pythoncom
import time
import sys
import os
import argparse
import io

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

parser = argparse.ArgumentParser(description='Diagnose alt-tab issues')
parser.add_argument('--auto', action='store_true', help='Run in non-interactive mode')
args = parser.parse_args()


def check_foreground_window():
    """Check what application is currently in the foreground."""
    try:
        import win32gui
        import win32process
        
        hwnd = win32gui.GetForegroundWindow()
        window_title = win32gui.GetWindowText(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        
        return {
            "hwnd": hwnd,
            "title": window_title,
            "pid": pid,
            "is_excel": "excel" in window_title.lower()
        }
    except ImportError:
        return {"error": "win32gui not available"}
    except Exception as e:
        return {"error": str(e)}


def test_excel_connection():
    """Test if Excel connection works regardless of foreground state."""
    try:
        import win32com.client as win32
        
        # Method 1: GetActiveObject (gets existing Excel)
        try:
            excel = win32.GetActiveObject("Excel.Application")
            active_workbook = excel.ActiveWorkbook
            active_sheet = excel.ActiveSheet
            selection = excel.Selection
            
            return {
                "method": "GetActiveObject",
                "success": True,
                "workbook": active_workbook.Name if active_workbook else None,
                "sheet": active_sheet.Name if active_sheet else None,
                "selection": selection.Address if selection else None,
                "excel_visible": excel.Visible,
                "excel_hwnd": excel.Hwnd
            }
        except Exception as e:
            return {
                "method": "GetActiveObject",
                "success": False,
                "error": str(e)
            }
            
    except ImportError:
        return {"error": "win32com not available"}


def test_xlwings_connection():
    """Test xlwings connection."""
    try:
        import xlwings as xw
        
        if len(xw.apps) == 0:
            return {"success": False, "error": "No Excel apps found"}
        
        app = xw.apps.active
        if app is None:
            return {"success": False, "error": "No active app"}
        
        book = app.books.active
        sheet = book.sheets.active if book else None
        selection = app.selection
        
        return {
            "success": True,
            "pid": app.pid,
            "workbook": book.name if book else None,
            "sheet": sheet.name if sheet else None,
            "selection": selection.address if selection else None
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def test_add_comment():
    """Test if we can add a comment to Excel."""
    try:
        import win32com.client as win32
        
        excel = win32.GetActiveObject("Excel.Application")
        selection = excel.Selection
        
        if selection is None:
            return {"success": False, "error": "No selection"}
        
        # Clear existing comment
        try:
            selection.ClearComments()
        except:
            pass
        
        # Add test comment
        timestamp = time.strftime('%H:%M:%S')
        test_comment = f"Alt-Tab Test @ {timestamp}"
        selection.AddComment(test_comment)
        
        # Verify
        comment = selection.Comment
        if comment:
            text = comment.Text()
            return {"success": True, "comment": text}
        else:
            return {"success": False, "error": "Comment not found after adding"}
            
    except Exception as e:
        return {"success": False, "error": str(e)}


def diagnose():
    print("=" * 60)
    print("       Alt-Tab Focus Change Diagnostic")
    print("=" * 60)
    
    pythoncom.CoInitialize()
    
    iterations = 5
    results = []
    
    for i in range(iterations):
        print(f"\n--- Test {i+1}/{iterations} ---")
        
        # Check foreground window
        fg = check_foreground_window()
        print(f"  Foreground: {fg.get('title', 'unknown')[:40]}")
        print(f"  Is Excel: {fg.get('is_excel', 'unknown')}")
        
        # Pump COM messages (critical!)
        pythoncom.PumpWaitingMessages()
        
        # Test connections
        print("  Testing connections...")
        
        win32_result = test_excel_connection()
        print(f"    win32com: {'OK' if win32_result.get('success') else 'FAIL'}")
        if not win32_result.get('success'):
            print(f"      Error: {win32_result.get('error')}")
        
        xw_result = test_xlwings_connection()
        print(f"    xlwings:  {'OK' if xw_result.get('success') else 'FAIL'}")
        if not xw_result.get('success'):
            print(f"      Error: {xw_result.get('error')}")
        
        # Test adding comment
        comment_result = test_add_comment()
        print(f"    Comment:  {'OK' if comment_result.get('success') else 'FAIL'}")
        if not comment_result.get('success'):
            print(f"      Error: {comment_result.get('error')}")
        
        results.append({
            "iteration": i + 1,
            "foreground_is_excel": fg.get('is_excel'),
            "win32_ok": win32_result.get('success'),
            "xlwings_ok": xw_result.get('success'),
            "comment_ok": comment_result.get('success')
        })
        
        # Wait and pump messages
        if i < iterations - 1:
            if not args.auto:
                input("\n  >>> ALT-TAB to another app, then ALT-TAB back to Excel, then press Enter...")
            else:
                print("  [Auto Mode] Waiting 2s (simulating focus change)...")
                # Pump messages during wait
                for _ in range(20):
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.1)
    
    # Summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    
    all_passed = all(r['comment_ok'] for r in results)
    print(f"\nTotal tests: {len(results)}")
    print(f"All passed: {'YES' if all_passed else 'NO'}")
    
    if not all_passed:
        print("\nFailed iterations:")
        for r in results:
            if not r['comment_ok']:
                print(f"  - Test {r['iteration']}: Excel in foreground: {r['foreground_is_excel']}")
    
    pythoncom.CoUninitialize()
    
    return all_passed


if __name__ == "__main__":
    if not args.auto:
        print("INSTRUCTIONS:")
        print("1. Make sure Excel is open with a cell selected")
        print("2. During the test, you'll alt-tab between apps")
        print("3. This tests if the tool works after focus changes")
        print()
        input("Press Enter to start...")
    else:
        print("[Auto Mode] Running immediately...")
    
    success = diagnose()
    
    if not args.auto:
        input("\nPress Enter to exit...")
    else:
        print(f"\n[Auto Mode] Done. Result: {'PASS' if success else 'FAIL'}")
