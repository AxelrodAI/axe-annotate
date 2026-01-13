"""Test that simulates the real hotkey workflow - multiple annotations in sequence.
This tests the worker queue pattern used by main.py.

Usage:
    python test_queue.py        # Interactive mode
    python test_queue.py --auto # Non-interactive mode (for agents)
"""
import pythoncom
import time
import threading
import queue
import io
import sys
import os
import argparse

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Add parent directory to path to import excel_ops and data_fetcher
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Parse arguments early
parser = argparse.ArgumentParser(description='Test queue-based annotation workflow')
parser.add_argument('--auto', action='store_true', help='Run in non-interactive mode (no input prompts)')
args = parser.parse_args()

# Import our modules (after setting up path)
import excel_ops
import data_fetcher

# Simulated task queue (same as main.py)
task_queue = queue.Queue()
results = []

def worker_loop():
    """Worker thread that processes annotation tasks."""
    print("[Worker] Starting, initializing COM...")
    pythoncom.CoInitialize()
    
    while True:
        try:
            task = task_queue.get(timeout=1)
            if task is None:
                break
                
            task_id, mode = task
            print(f"\n[Worker] Processing task #{task_id}...")
            
            try:
                # Step 1: Get selection (this is where it might fail on 2nd+ call)
                app, book, sheet, selection = excel_ops.get_active_selection()
                
                if not selection:
                    results.append((task_id, False, "No selection"))
                    print(f"[Worker] Task #{task_id} FAILED: No selection")
                    task_queue.task_done()
                    continue
                
                addr = selection.address
                print(f"[Worker] Task #{task_id}: Selection = {addr}")
                
                # Step 2: Get context
                context = excel_ops.get_context(selection)
                print(f"[Worker] Task #{task_id}: Context = {context.get('line_item')} / {context.get('time_period')}")
                
                # Step 3: Dummy comment (skip data_fetcher to speed up)
                comment = f"Test Annotation #{task_id}\nCell: {addr}\nTime: {time.strftime('%H:%M:%S')}"
                
                # Step 4: Add comment
                success = excel_ops.add_note_to_cell(selection, comment)
                
                if success:
                    results.append((task_id, True, addr))
                    print(f"[Worker] Task #{task_id} SUCCESS: Annotated {addr}")
                else:
                    results.append((task_id, False, "add_note_to_cell returned False"))
                    print(f"[Worker] Task #{task_id} FAILED: Could not add note")
                    
            except Exception as e:
                results.append((task_id, False, str(e)))
                print(f"[Worker] Task #{task_id} ERROR: {e}")
            
            # Cooldown (same as main.py)
            time.sleep(0.2)
            task_queue.task_done()
            
        except queue.Empty:
            continue
        except Exception as e:
            print(f"[Worker] Loop error: {e}")
    
    pythoncom.CoUninitialize()
    print("[Worker] Stopped.")


def run_test():
    print("=" * 60)
    print("     Multi-Annotation Queue Test")
    print("=" * 60)
    print("\nThis tests the same worker queue pattern as main.py.")
    print("Select different cells while the test runs to simulate real usage.\n")
    
    # Start worker thread
    worker = threading.Thread(target=worker_loop, daemon=True)
    worker.start()
    
    print("Worker started. Submitting 5 annotation tasks...")
    print(">>> IMPORTANT: Click different cells in Excel between each task!\n")
    
    for i in range(1, 6):
        print(f"--- Submitting Task #{i} ---")
        if not args.auto:
            print("    (Click a different cell in Excel NOW)")
            time.sleep(2)  # Give user time to click a cell
        else:
            time.sleep(0.5)  # Shorter delay in auto mode
        
        task_queue.put((i, "v1"))
        
        # Wait for task to complete
        task_queue.join()
        print()
    
    # Stop worker
    task_queue.put(None)
    worker.join(timeout=2)
    
    # Summary
    print("\n" + "=" * 60)
    print("                   TEST SUMMARY")
    print("=" * 60)
    succeeded = [r for r in results if r[1]]
    failed = [r for r in results if not r[1]]
    
    print(f"\nTotal Tasks: {len(results)}")
    print(f"Succeeded:   {len(succeeded)}")
    print(f"Failed:      {len(failed)}")
    
    if failed:
        print("\nFailed tasks:")
        for task_id, _, reason in failed:
            print(f"  - Task #{task_id}: {reason}")
    
    if len(succeeded) == len(results):
        print("\n*** ALL TESTS PASSED! ***")
    else:
        print("\n*** SOME TESTS FAILED - See details above ***")
    
    print("=" * 60)


if __name__ == "__main__":
    if not args.auto:
        input("Open Excel with a workbook, select a cell, then press Enter to start...")
    else:
        print("[Auto Mode] Skipping input prompt, running immediately...")
    
    run_test()
    
    if not args.auto:
        input("\nPress Enter to exit...")
    else:
        print("[Auto Mode] Queue test complete.")
