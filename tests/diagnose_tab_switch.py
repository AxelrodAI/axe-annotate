"""
Diagnostic script to test Excel tab/workbook switching behavior.
This specifically tests the scenario where hotkeys stop working after switching tabs.

Usage:
    python diagnose_tab_switch.py        # Interactive mode
    python diagnose_tab_switch.py --auto # Non-interactive mode
"""
import pythoncom
import time
import sys
import argparse
import io

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

parser = argparse.ArgumentParser(description='Diagnose tab switching issues')
parser.add_argument('--auto', action='store_true', help='Run in non-interactive mode')
args = parser.parse_args()


def diagnose():
    print("=" * 60)
    print("       Tab/Workbook Switching Diagnostic")
    print("=" * 60)
    
    print("\nInitializing COM...")
    pythoncom.CoInitialize()
    
    try:
        import xlwings as xw
        print(f"xlwings version: {xw.__version__}")
    except Exception as e:
        print(f"FATAL: Cannot import xlwings: {e}")
        return
    
    # Test 1: Basic connection
    print("\n" + "-" * 40)
    print("[Test 1] Basic Connection Check")
    print("-" * 40)
    
    try:
        num_apps = len(xw.apps)
        print(f"  Excel instances found: {num_apps}")
        
        if num_apps == 0:
            print("  FAIL: No Excel running!")
            return
        
        app = xw.apps.active
        print(f"  Active app PID: {app.pid}")
        print(f"  Excel version: {app.api.Version}")
        
        # List ALL open workbooks
        print(f"\n  Open workbooks ({len(app.books)}):")
        for i, book in enumerate(app.books):
            is_active = " <-- ACTIVE" if book == app.books.active else ""
            print(f"    {i+1}. {book.name}{is_active}")
            
    except Exception as e:
        print(f"  FAIL: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # Test 2: Selection tracking across apps
    print("\n" + "-" * 40)
    print("[Test 2] Selection State Inspection")
    print("-" * 40)
    
    try:
        # Method A: xw.apps.active.selection (xlwings wrapper)
        print("\n  Method A: xw.apps.active.selection")
        try:
            sel_a = xw.apps.active.selection
            print(f"    Address: {sel_a.address}")
            print(f"    Sheet: {sel_a.sheet.name}")
            print(f"    Book: {sel_a.sheet.book.name}")
        except Exception as e:
            print(f"    FAILED: {e}")
        
        # Method B: Direct COM API
        print("\n  Method B: app.api.Selection (Direct COM)")
        try:
            app = xw.apps.active
            api_sel = app.api.Selection
            print(f"    Address: {api_sel.Address}")
            print(f"    Parent Sheet: {api_sel.Worksheet.Name}")
            print(f"    Parent Book: {api_sel.Worksheet.Parent.Name}")
        except Exception as e:
            print(f"    FAILED: {e}")
        
        # Method C: ActiveCell
        print("\n  Method C: app.api.ActiveCell (Direct COM)")
        try:
            app = xw.apps.active
            active_cell = app.api.ActiveCell
            print(f"    Address: {active_cell.Address}")
            print(f"    Parent Sheet: {active_cell.Worksheet.Name}")
            print(f"    Parent Book: {active_cell.Worksheet.Parent.Name}")
        except Exception as e:
            print(f"    FAILED: {e}")
            
        # Method D: ActiveWorkbook + ActiveSheet
        print("\n  Method D: ActiveWorkbook.ActiveSheet (Direct COM)")
        try:
            app = xw.apps.active
            active_book = app.api.ActiveWorkbook
            active_sheet = app.api.ActiveSheet
            print(f"    Active Workbook: {active_book.Name}")
            print(f"    Active Sheet: {active_sheet.Name}")
        except Exception as e:
            print(f"    FAILED: {e}")
            
    except Exception as e:
        print(f"  FAIL: {e}")
        import traceback
        traceback.print_exc()
    
    # Test 3: Simulated tab switch detection
    print("\n" + "-" * 40)
    print("[Test 3] Fresh Reference Acquisition")
    print("-" * 40)
    print("  (Testing if we can get fresh references after state change)")
    
    for i in range(3):
        print(f"\n  Iteration {i+1}:")
        try:
            # Clear any cached references by getting fresh ones
            fresh_app = xw.apps.active
            fresh_book = fresh_app.books.active
            fresh_sheet = fresh_book.sheets.active
            fresh_selection = fresh_app.selection
            
            # Read properties to force COM resolution
            book_name = fresh_book.name
            sheet_name = fresh_sheet.name
            sel_addr = fresh_selection.address
            
            print(f"    Book: {book_name}")
            print(f"    Sheet: {sheet_name}")
            print(f"    Selection: {sel_addr}")
            
            # Try to add a comment
            test_comment = f"Test {i+1} @ {time.strftime('%H:%M:%S')}"
            try:
                fresh_selection.api.ClearComments()
            except:
                pass
            time.sleep(0.05)
            fresh_selection.api.AddComment(test_comment)
            print(f"    Comment added: OK")
            
        except Exception as e:
            print(f"    FAILED: {e}")
        
        if i < 2:
            if not args.auto:
                input(f"\n  >>> SWITCH to a different tab/workbook in Excel, then press Enter...")
            else:
                print(f"  [Auto Mode] Waiting 1s (simulate tab switch)...")
                time.sleep(1)
    
    # Test 4: Check for lingering COM issues
    print("\n" + "-" * 40)
    print("[Test 4] COM State Check")
    print("-" * 40)
    
    try:
        import win32com.client as win32
        
        # Try to get Excel via win32com directly (bypassing xlwings)
        print("\n  Checking via win32com.client.GetActiveObject...")
        try:
            excel = win32.GetActiveObject("Excel.Application")
            print(f"    Excel.Application obtained")
            print(f"    ActiveWorkbook: {excel.ActiveWorkbook.Name}")
            print(f"    ActiveSheet: {excel.ActiveSheet.Name}")
            print(f"    ActiveCell: {excel.ActiveCell.Address}")
        except Exception as e:
            print(f"    Note: {e}")
            
        # Try Dispatch (creates new connection)
        print("\n  Checking via win32com.client.Dispatch (fresh connection)...")
        try:
            excel2 = win32.Dispatch("Excel.Application")
            print(f"    Excel.Application dispatched")
            print(f"    ActiveWorkbook: {excel2.ActiveWorkbook.Name}")
            print(f"    ActiveSheet: {excel2.ActiveSheet.Name}")
        except Exception as e:
            print(f"    Note: {e}")
            
    except ImportError:
        print("  win32com not available (pywin32 not installed)")
    except Exception as e:
        print(f"  Error: {e}")
    
    print("\n" + "=" * 60)
    print("Diagnostic Complete!")
    print("=" * 60)
    
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    if not args.auto:
        print("INSTRUCTIONS:")
        print("1. Open Excel with at least 2 workbooks OR 2 sheets")
        print("2. Select a cell in the first sheet/workbook")
        print("3. Press Enter to start the diagnostic")
        print()
        input("Press Enter when ready...")
    else:
        print("[Auto Mode] Running immediately...")
    
    diagnose()
    
    if not args.auto:
        input("\nPress Enter to exit...")
    else:
        print("[Auto Mode] Done.")
