"""Stress test for Excel annotation operations.

Usage:
    python stress_test_excel.py         # Interactive full test
    python stress_test_excel.py --quick # Quick connection test
    python stress_test_excel.py --auto  # Non-interactive mode (for agents)
"""
import sys
import os
# Add parent directory to path to import excel_ops and data_fetcher
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import excel_ops
import data_fetcher
import time
import pythoncom
import argparse

# Parse arguments early
parser = argparse.ArgumentParser(description='Stress test Excel annotation')
parser.add_argument('--quick', action='store_true', help='Run quick connection test only')
parser.add_argument('--auto', action='store_true', help='Run in non-interactive mode (no input prompts)')
args = parser.parse_args()

def stress_test():
    """
    Stress test that runs multiple annotation cycles to verify reliability.
    """
    print("=" * 60)
    print("           Axe Annotate Stress Test")
    print("=" * 60)
    print("\nPREREQUISITES:")
    print("  1. Excel must be OPEN")
    print("  2. A workbook must be ACTIVE")
    print("  3. You must NOT be editing a cell (press Esc first)")
    print("  4. Select any cell before starting")
    print("\n" + "=" * 60)
    
    if not args.auto:
        input("Press Enter to start the test...")
    else:
        print("[Auto Mode] Skipping input prompt...")
    print()

    # Initialize COM for this thread (critical for reliability!)
    pythoncom.CoInitialize()
    print("[Setup] COM initialized for this thread.\n")

    # First, run a health check
    print("[Health Check] Testing Excel connection...")
    success, message = excel_ops.test_connection()
    if not success:
        print(f"[Health Check] FAILED: {message}")
        print("\nPlease fix the issue and try again.")
        pythoncom.CoUninitialize()
        return
    print(f"[Health Check] PASSED: {message}\n")

    # Run the stress test
    num_iterations = 5
    successes = 0
    failures = 0

    for i in range(1, num_iterations + 1):
        print(f"--- Iteration {i}/{num_iterations} ---")
        try:
            # 1. Connect
            app, book, sheet, selection = excel_ops.get_active_selection()
            if not selection:
                print("  RESULT: FAIL - No selection available")
                failures += 1
                continue
                
            print(f"  Selection: {selection.address}")
            
            # 2. Context
            context = excel_ops.get_context(selection)
            line_item = context.get('line_item', 'N/A')
            time_period = context.get('time_period', 'N/A')
            print(f"  Context: {line_item} | {time_period}")
            
            # 3. Write Note
            comments = f"Stress Test Comment #{i}\nTimestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}"
            result = excel_ops.add_note_to_cell(selection, comments)
            
            if result:
                print("  RESULT: SUCCESS")
                successes += 1
            else:
                print("  RESULT: FAIL - Could not add note")
                failures += 1
            
        except Exception as e:
            print(f"  RESULT: ERROR - {e}")
            failures += 1
        
        # Small delay to mimic realistic user speed
        if i < num_iterations:
            time.sleep(0.5)
    
    # Cleanup COM
    pythoncom.CoUninitialize()
    
    # Summary
    print("\n" + "=" * 60)
    print("                    TEST SUMMARY")
    print("=" * 60)
    print(f"  Total Iterations: {num_iterations}")
    print(f"  Successes:        {successes}")
    print(f"  Failures:         {failures}")
    print(f"  Success Rate:     {(successes / num_iterations) * 100:.1f}%")
    print("=" * 60)
    
    if failures == 0:
        print("\n[PASS] ALL TESTS PASSED! The tool is working reliably.")
    else:
        print(f"\n[FAIL] {failures} test(s) failed. Review the output above for details.")


def quick_test():
    """
    Single iteration test for quick verification.
    """
    print("Quick Connection Test\n")
    
    pythoncom.CoInitialize()
    
    success, message = excel_ops.test_connection()
    print(f"Result: {'PASS' if success else 'FAIL'}")
    print(f"Details: {message}")
    
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    if args.quick:
        quick_test()
    else:
        try:
            stress_test()
        except KeyboardInterrupt:
            print("\nTest interrupted by user.")
        except Exception as e:
            print(f"\nFatal Test Error: {e}")
        finally:
            if not args.auto:
                input("\nPress Enter to exit...")
            else:
                print("[Auto Mode] Stress test complete.")
