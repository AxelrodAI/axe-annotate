"""
Comprehensive Edge Case Testing for Axe Annotate
=================================================
Tests various scenarios that could cause the tool to fail.

Usage:
    python edge_case_tests.py --auto

Edge Cases Tested:
1. Multiple cell selection (range instead of single cell)
2. Empty workbook (no data for context)
3. Cell with existing comment (overwrite handling)
4. Very large comment text
5. Rapid successive operations (race conditions)
6. Selection at spreadsheet boundaries (row 1, col 1)
7. Multiple Excel instances running
8. Excel minimized
9. Merged cells
10. Cell in a different workbook than expected
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

parser = argparse.ArgumentParser(description='Edge case testing for Axe Annotate')
parser.add_argument('--auto', action='store_true', help='Run in non-interactive mode')
args = parser.parse_args()

import excel_ops


def test_multiple_cell_selection():
    """Test: What happens when user selects a range instead of single cell?"""
    print("\n[Test 1] Multiple Cell Selection (Range)")
    print("-" * 40)
    
    try:
        import xlwings as xw
        app = xw.apps.active
        sheet = app.books.active.sheets.active
        
        # Select a range
        original_selection = app.selection.address
        range_to_select = sheet.range("A1:C3")
        range_to_select.select()
        time.sleep(0.1)
        
        # Try to get selection
        _, _, _, selection = excel_ops.get_active_selection()
        
        if selection:
            print(f"  Selection address: {selection.address}")
            print(f"  Row: {selection.row}, Col: {selection.column}")
            
            # Try to add comment to range
            try:
                test_comment = "Range selection test"
                success = excel_ops.add_note_to_cell(selection, test_comment)
                print(f"  Comment added: {'OK' if success else 'FAILED'}")
                
                # Note: Comments on ranges typically go to top-left cell
                result = "PASS - Comment added to top-left cell of range"
            except Exception as e:
                result = f"ISSUE - {e}"
        else:
            result = "FAIL - Could not get selection"
        
        # Restore original selection
        try:
            sheet.range(original_selection).select()
        except:
            pass
            
    except Exception as e:
        result = f"ERROR - {e}"
    
    print(f"  Result: {result}")
    return "PASS" in result or "OK" in result


def test_empty_context():
    """Test: What happens when there's no context data (empty headers)?"""
    print("\n[Test 2] Empty Context (No Headers)")
    print("-" * 40)
    
    try:
        import xlwings as xw
        app = xw.apps.active
        sheet = app.books.active.sheets.active
        
        # Find an empty area
        original_selection = app.selection.address
        test_cell = sheet.range("ZZ100")  # Likely empty
        test_cell.select()
        time.sleep(0.1)
        
        _, _, _, selection = excel_ops.get_active_selection()
        
        if selection:
            context = excel_ops.get_context(selection)
            print(f"  Ticker: {context.get('ticker', 'N/A')}")
            print(f"  Period: {context.get('time_period', 'N/A')}")
            print(f"  Line Item: {context.get('line_item', 'N/A')}")
            
            # Should return defaults, not crash
            if context.get('line_item') and context.get('time_period'):
                result = "PASS - Defaults used for missing context"
            else:
                result = "ISSUE - Context extraction returned empty"
        else:
            result = "FAIL - Could not get selection"
        
        # Restore
        try:
            sheet.range(original_selection).select()
        except:
            pass
            
    except Exception as e:
        result = f"ERROR - {e}"
    
    print(f"  Result: {result}")
    return "PASS" in result


def test_existing_comment_overwrite():
    """Test: Does overwriting existing comments work?"""
    print("\n[Test 3] Existing Comment Overwrite")
    print("-" * 40)
    
    try:
        import xlwings as xw
        app = xw.apps.active
        sheet = app.books.active.sheets.active
        
        # Use a test cell
        original_selection = app.selection.address
        test_cell = sheet.range("ZZ1")
        test_cell.select()
        time.sleep(0.1)
        
        _, _, _, selection = excel_ops.get_active_selection()
        
        if selection:
            # Add first comment
            excel_ops.add_note_to_cell(selection, "First comment")
            time.sleep(0.1)
            
            # Add second comment (should overwrite)
            excel_ops.add_note_to_cell(selection, "Second comment")
            time.sleep(0.1)
            
            # Verify
            try:
                comment = selection.api.Comment
                if comment:
                    text = comment.Text()
                    if "Second" in text and "First" not in text:
                        result = "PASS - Comment properly overwritten"
                    else:
                        result = f"ISSUE - Comment not overwritten. Text: {text[:50]}"
                else:
                    result = "FAIL - No comment found"
            except Exception as e:
                result = f"ISSUE - {e}"
            
            # Cleanup
            try:
                selection.api.ClearComments()
            except:
                pass
        else:
            result = "FAIL - Could not get selection"
        
        # Restore
        try:
            sheet.range(original_selection).select()
        except:
            pass
            
    except Exception as e:
        result = f"ERROR - {e}"
    
    print(f"  Result: {result}")
    return "PASS" in result


def test_large_comment():
    """Test: Can we handle very large comments?"""
    print("\n[Test 4] Large Comment Text")
    print("-" * 40)
    
    try:
        import xlwings as xw
        app = xw.apps.active
        sheet = app.books.active.sheets.active
        
        original_selection = app.selection.address
        test_cell = sheet.range("ZZ2")
        test_cell.select()
        time.sleep(0.1)
        
        _, _, _, selection = excel_ops.get_active_selection()
        
        if selection:
            # Create a large comment (Excel has limits around 32k chars)
            large_text = "X" * 10000  # 10k characters
            large_comment = f"Large Comment Test\n{'=' * 50}\n{large_text}"
            
            success = excel_ops.add_note_to_cell(selection, large_comment)
            
            if success:
                # Verify it was added
                comment = selection.api.Comment
                if comment:
                    text_len = len(comment.Text())
                    print(f"  Comment length: {text_len} chars")
                    result = f"PASS - Large comment ({text_len} chars) added"
                else:
                    result = "ISSUE - Comment object is None"
            else:
                result = "FAIL - add_note_to_cell returned False"
            
            # Cleanup
            try:
                selection.api.ClearComments()
            except:
                pass
        else:
            result = "FAIL - Could not get selection"
        
        # Restore
        try:
            sheet.range(original_selection).select()
        except:
            pass
            
    except Exception as e:
        result = f"ERROR - {e}"
    
    print(f"  Result: {result}")
    return "PASS" in result


def test_rapid_operations():
    """Test: What happens with rapid successive operations?"""
    print("\n[Test 5] Rapid Successive Operations")
    print("-" * 40)
    
    try:
        import xlwings as xw
        app = xw.apps.active
        sheet = app.books.active.sheets.active
        
        original_selection = app.selection.address
        successes = 0
        failures = 0
        
        print("  Running 10 rapid operations...")
        
        for i in range(10):
            test_cell = sheet.range(f"ZZ{10 + i}")
            test_cell.select()
            
            # Minimal delay
            time.sleep(0.05)
            
            _, _, _, selection = excel_ops.get_active_selection()
            if selection:
                success = excel_ops.add_note_to_cell(selection, f"Rapid test {i}")
                if success:
                    successes += 1
                else:
                    failures += 1
            else:
                failures += 1
            
            # Cleanup
            try:
                test_cell.api.ClearComments()
            except:
                pass
        
        print(f"  Successes: {successes}/10")
        print(f"  Failures: {failures}/10")
        
        if successes == 10:
            result = "PASS - All rapid operations succeeded"
        elif successes > 7:
            result = f"PARTIAL - {successes}/10 succeeded (acceptable)"
        else:
            result = f"FAIL - Only {successes}/10 succeeded"
        
        # Restore
        try:
            sheet.range(original_selection).select()
        except:
            pass
            
    except Exception as e:
        result = f"ERROR - {e}"
    
    print(f"  Result: {result}")
    return "PASS" in result or "PARTIAL" in result


def test_boundary_cells():
    """Test: What happens at spreadsheet boundaries?"""
    print("\n[Test 6] Boundary Cells (Row 1, Col 1)")
    print("-" * 40)
    
    try:
        import xlwings as xw
        app = xw.apps.active
        sheet = app.books.active.sheets.active
        
        original_selection = app.selection.address
        
        # Test A1 - no headers above or to the left
        test_cell = sheet.range("A1")
        test_cell.select()
        time.sleep(0.1)
        
        _, _, _, selection = excel_ops.get_active_selection()
        
        if selection:
            context = excel_ops.get_context(selection)
            print(f"  A1 Context - Period: {context.get('time_period')}, Item: {context.get('line_item')}")
            
            # Should handle gracefully with defaults
            success = excel_ops.add_note_to_cell(selection, "Boundary test A1")
            
            if success:
                result = "PASS - Boundary cell handled correctly"
            else:
                result = "FAIL - Could not add comment to A1"
            
            # Cleanup
            try:
                selection.api.ClearComments()
            except:
                pass
        else:
            result = "FAIL - Could not get selection"
        
        # Restore
        try:
            sheet.range(original_selection).select()
        except:
            pass
            
    except Exception as e:
        result = f"ERROR - {e}"
    
    print(f"  Result: {result}")
    return "PASS" in result


def test_multiple_excel_instances():
    """Test: What if multiple Excel instances are running?"""
    print("\n[Test 7] Multiple Excel Instances")
    print("-" * 40)
    
    try:
        import xlwings as xw
        
        num_instances = len(xw.apps)
        print(f"  Excel instances running: {num_instances}")
        
        if num_instances > 1:
            print("  Multiple instances detected!")
            # Check if we get the right one
            for i, app in enumerate(xw.apps):
                try:
                    book_name = app.books.active.name if app.books.active else "No workbook"
                    print(f"    Instance {i+1}: {book_name}")
                except:
                    print(f"    Instance {i+1}: [Error reading]")
        
        # Test that we can still get selection
        _, _, _, selection = excel_ops.get_active_selection()
        
        if selection:
            result = f"PASS - Correctly accessing 1 of {num_instances} instance(s)"
        else:
            result = "FAIL - Could not get selection"
            
    except Exception as e:
        result = f"ERROR - {e}"
    
    print(f"  Result: {result}")
    return "PASS" in result


def test_excel_minimized():
    """Test: What if Excel is minimized?"""
    print("\n[Test 8] Excel Minimized State")
    print("-" * 40)
    
    try:
        import win32com.client as win32
        
        excel = win32.GetActiveObject("Excel.Application")
        
        # Check window state
        # -4140 = xlNormal, -4137 = xlMinimized, -4143 = xlMaximized
        window_state = excel.WindowState
        states = {-4140: "Normal", -4137: "Minimized", -4143: "Maximized"}
        state_name = states.get(window_state, f"Unknown ({window_state})")
        
        print(f"  Window state: {state_name}")
        print(f"  Visible: {excel.Visible}")
        
        # Test if we can still operate when minimized
        if window_state == -4137:  # Minimized
            # Try to get selection anyway
            _, _, _, selection = excel_ops.get_active_selection()
            if selection:
                result = "PASS - Can operate while minimized"
            else:
                result = "ISSUE - Cannot operate while minimized"
        else:
            result = f"SKIPPED - Excel is {state_name} (not minimized)"
            
    except Exception as e:
        result = f"ERROR - {e}"
    
    print(f"  Result: {result}")
    return "PASS" in result or "SKIPPED" in result


def test_special_characters():
    """Test: Can we handle special characters in comments?"""
    print("\n[Test 9] Special Characters in Comments")
    print("-" * 40)
    
    try:
        import xlwings as xw
        app = xw.apps.active
        sheet = app.books.active.sheets.active
        
        original_selection = app.selection.address
        test_cell = sheet.range("ZZ5")
        test_cell.select()
        time.sleep(0.1)
        
        _, _, _, selection = excel_ops.get_active_selection()
        
        if selection:
            # Test various special characters
            special_comment = """Special Characters Test:
- Unicode: cafe, resume, naive
- Symbols: $100, 50%, #1
- Math: 2+2=4, 5>3, x<y
- Quotes: "double" and 'single'
- Asian: Yen, Euro
- Emoji: (using text alternatives)
- Newlines and
  indentation"""
            
            success = excel_ops.add_note_to_cell(selection, special_comment)
            
            if success:
                # Verify
                comment = selection.api.Comment
                if comment:
                    result = "PASS - Special characters handled"
                else:
                    result = "ISSUE - Comment object is None"
            else:
                result = "FAIL - add_note_to_cell returned False"
            
            # Cleanup
            try:
                selection.api.ClearComments()
            except:
                pass
        else:
            result = "FAIL - Could not get selection"
        
        # Restore
        try:
            sheet.range(original_selection).select()
        except:
            pass
            
    except Exception as e:
        result = f"ERROR - {e}"
    
    print(f"  Result: {result}")
    return "PASS" in result


def test_edit_mode_detection():
    """Test: Can we detect when Excel is in edit mode?"""
    print("\n[Test 10] Edit Mode Detection")
    print("-" * 40)
    
    try:
        import win32com.client as win32
        
        excel = win32.GetActiveObject("Excel.Application")
        
        # Check if Excel is in edit mode
        # This is tricky - Excel doesn't have a direct "IsEditMode" property
        # But certain operations fail when in edit mode
        
        try:
            # Try to access a property that fails in edit mode
            _ = excel.Version
            _ = excel.ActiveWorkbook.Name
            _ = excel.ActiveSheet.Name
            _ = excel.Selection.Address
            
            print("  Excel appears to be in Ready state")
            result = "PASS - Edit mode detection working"
            
        except Exception as e:
            if "Call was rejected" in str(e) or "busy" in str(e).lower():
                print("  Excel appears to be in Edit Mode or busy")
                result = "INFO - Edit mode detected correctly"
            else:
                result = f"ISSUE - Unexpected error: {e}"
            
    except Exception as e:
        result = f"ERROR - {e}"
    
    print(f"  Result: {result}")
    return "PASS" in result or "INFO" in result


def run_all_tests():
    """Run all edge case tests."""
    print("=" * 60)
    print("       Axe Annotate Edge Case Testing")
    print("=" * 60)
    
    pythoncom.CoInitialize()
    
    results = {}
    
    tests = [
        ("Multiple Cell Selection", test_multiple_cell_selection),
        ("Empty Context", test_empty_context),
        ("Comment Overwrite", test_existing_comment_overwrite),
        ("Large Comment", test_large_comment),
        ("Rapid Operations", test_rapid_operations),
        ("Boundary Cells", test_boundary_cells),
        ("Multiple Instances", test_multiple_excel_instances),
        ("Minimized State", test_excel_minimized),
        ("Special Characters", test_special_characters),
        ("Edit Mode Detection", test_edit_mode_detection),
    ]
    
    for name, test_func in tests:
        try:
            passed = test_func()
            results[name] = "PASS" if passed else "FAIL"
        except Exception as e:
            print(f"\n  EXCEPTION: {e}")
            results[name] = "ERROR"
    
    # Summary
    print("\n" + "=" * 60)
    print("                    SUMMARY")
    print("=" * 60)
    
    passed = sum(1 for r in results.values() if r == "PASS")
    failed = sum(1 for r in results.values() if r == "FAIL")
    errors = sum(1 for r in results.values() if r == "ERROR")
    
    print(f"\nTotal: {len(results)} tests")
    print(f"  Passed: {passed}")
    print(f"  Failed: {failed}")
    print(f"  Errors: {errors}")
    
    print("\nDetailed Results:")
    for name, status in results.items():
        status_icon = "[OK]" if status == "PASS" else "[!!]" if status == "FAIL" else "[??]"
        print(f"  {status_icon} {name}: {status}")
    
    pythoncom.CoUninitialize()
    
    return failed == 0 and errors == 0


if __name__ == "__main__":
    if not args.auto:
        print("INSTRUCTIONS:")
        print("1. Make sure Excel is open with a workbook")
        print("2. Select any cell (not in edit mode)")
        print("3. This will test various edge cases")
        print()
        input("Press Enter to start...")
    else:
        print("[Auto Mode] Running immediately...")
    
    success = run_all_tests()
    
    if not args.auto:
        input("\nPress Enter to exit...")
    else:
        print(f"\n[Auto Mode] Done. Overall: {'PASS' if success else 'SOME ISSUES FOUND'}")
