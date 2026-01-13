"""
Debug script to identify Excel annotation issues.
Run this while Excel is open with a cell selected.

Usage:
    python debug_excel.py         # Interactive mode (waits for Enter)
    python debug_excel.py --auto  # Non-interactive mode (for agents/automation)
"""
import pythoncom
import time
import sys
import argparse

# Fix encoding for Windows console
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Parse arguments early
parser = argparse.ArgumentParser(description='Debug Excel annotation issues')
parser.add_argument('--auto', action='store_true', help='Run in non-interactive mode (no input prompts)')
args = parser.parse_args()

def run_debug():
    print("=" * 60)
    print("         Excel Annotation Debug Tool")
    print("=" * 60)
    print("\nInitializing COM...")
    pythoncom.CoInitialize()
    
    print("\n[Step 1] Importing xlwings...")
    try:
        import xlwings as xw
        print(f"  OK - xlwings version: {xw.__version__}")
    except Exception as e:
        print(f"  FAIL - Failed to import xlwings: {e}")
        return
    
    print("\n[Step 2] Checking for Excel instances...")
    try:
        num_apps = len(xw.apps)
        print(f"  Found {num_apps} Excel instance(s)")
        if num_apps == 0:
            print("  FAIL - No Excel running. Please open Excel first!")
            return
    except Exception as e:
        print(f"  FAIL - Error checking apps: {e}")
        return
    
    print("\n[Step 3] Getting active app...")
    try:
        app = xw.apps.active
        print(f"  OK - App PID: {app.pid}")
        print(f"  OK - Version: {app.api.Version}")
        print(f"  OK - Visible: {app.visible}")
    except Exception as e:
        print(f"  FAIL - Error getting app: {e}")
        return
    
    print("\n[Step 4] Getting active workbook...")
    try:
        book = app.books.active
        if book:
            print(f"  OK - Workbook: {book.name}")
            print(f"  OK - Full path: {book.fullname}")
        else:
            print("  FAIL - No active workbook!")
            return
    except Exception as e:
        print(f"  FAIL - Error getting workbook: {e}")
        return
    
    print("\n[Step 5] Getting active sheet...")
    try:
        sheet = book.sheets.active
        if sheet:
            print(f"  OK - Sheet: {sheet.name}")
            print(f"  OK - Index: {sheet.index}")
        else:
            print("  FAIL - No active sheet!")
            return
    except Exception as e:
        print(f"  FAIL - Error getting sheet: {e}")
        return
    
    print("\n[Step 6] Getting selection (CRITICAL)...")
    try:
        # Method 1: Through app.selection
        selection = app.selection
        print(f"  Method 1 (app.selection):")
        print(f"    Address: {selection.address}")
        print(f"    Row: {selection.row}, Column: {selection.column}")
        print(f"    Value: {selection.value}")
    except Exception as e:
        print(f"  FAIL - Method 1 failed: {e}")
    
    try:
        # Method 2: Through API directly
        api_selection = app.api.Selection
        print(f"  Method 2 (app.api.Selection):")
        print(f"    Address: {api_selection.Address}")
        print(f"    Row: {api_selection.Row}, Column: {api_selection.Column}")
    except Exception as e:
        print(f"  FAIL - Method 2 failed: {e}")
    
    try:
        # Method 3: Through ActiveCell
        active_cell = app.api.ActiveCell
        print(f"  Method 3 (app.api.ActiveCell):")
        print(f"    Address: {active_cell.Address}")
        print(f"    Row: {active_cell.Row}, Column: {active_cell.Column}")
    except Exception as e:
        print(f"  FAIL - Method 3 failed: {e}")
    
    print("\n[Step 7] Testing note/comment addition...")
    try:
        # Use the selection we got
        selection = app.selection
        cell_api = selection.api
        
        # Try to clear comments
        print("  Clearing existing comments...")
        try:
            cell_api.ClearComments()
            print("    OK - ClearComments succeeded")
        except Exception as e:
            print(f"    Note: ClearComments: {e}")
        
        time.sleep(0.1)
        
        # Add a test comment
        test_comment = f"Debug Test @ {time.strftime('%H:%M:%S')}"
        print(f"  Adding comment: '{test_comment}'...")
        try:
            cell_api.AddComment(test_comment)
            print("    OK - AddComment succeeded!")
        except Exception as e:
            print(f"    FAIL - AddComment FAILED: {e}")
            
            # Try alternative method
            print("  Trying alternative NoteText method...")
            try:
                cell_api.NoteText(test_comment)
                print("    OK - NoteText succeeded!")
            except Exception as e2:
                print(f"    FAIL - NoteText also FAILED: {e2}")
        
        # Verify comment was added
        time.sleep(0.1)
        print("  Verifying comment...")
        try:
            comment = cell_api.Comment
            if comment:
                text = comment.Text()
                print(f"    OK - Comment verified: '{text[:50]}...'")
            else:
                print("    ? Comment object is None")
        except Exception as e:
            print(f"    ? Verification issue: {e}")
            
    except Exception as e:
        print(f"  FAIL - Comment test failed: {e}")
    
    print("\n[Step 8] Testing with a NEW cell (simulating your issue)...")
    try:
        # Find an empty cell in the next row
        current_row = app.selection.row
        current_col = app.selection.column
        new_row = current_row + 10  # Go 10 rows down to find empty space
        
        print(f"  Selecting new cell at row {new_row}, col {current_col}...")
        
        # Select the new cell
        new_cell = sheet.range((new_row, current_col))
        new_cell.select()
        
        time.sleep(0.2)  # Give Excel time to update
        
        # Now re-read the selection
        print("  Re-reading selection after selecting new cell...")
        fresh_selection = app.selection
        print(f"    Address: {fresh_selection.address}")
        print(f"    Row: {fresh_selection.row}")
        
        # Try to add comment to the new cell
        print("  Adding comment to NEW cell...")
        fresh_selection.api.ClearComments()
        fresh_selection.api.AddComment(f"NEW CELL Test @ {time.strftime('%H:%M:%S')}")
        print("    OK - SUCCESS on NEW cell!")
        
        # Go back to original cell
        sheet.range((current_row, current_col)).select()
        
    except Exception as e:
        print(f"  FAIL - New cell test failed: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "=" * 60)
    print("Debug complete! Check the output above for issues.")
    print("=" * 60)
    
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    if not args.auto:
        input("Make sure Excel is open with a cell selected, then press Enter...")
    else:
        print("[Auto Mode] Skipping input prompt, running immediately...")
    
    run_debug()
    
    if not args.auto:
        input("\nPress Enter to exit...")
    else:
        print("[Auto Mode] Debug complete.")
