"""
Axe Annotate - Main Entry Point
================================
A tool for Hedge Fund Analysts to annotate Excel cells with contextual comments.

Architecture Overview:
- Main thread: Registers keyboard hotkeys and waits for Esc to quit
- Worker thread: Handles all Excel COM operations via a task queue
- Task queue: Decouples hotkey handlers from COM operations

Why a separate worker thread?
- COM operations must happen in a thread that initializes COM
- Hotkey callbacks run in the keyboard library's thread
- We queue tasks and process them in our COM-initialized worker

Usage:
    python main.py

Hotkeys:
    Ctrl+Shift+m  - Auto-annotate selected cell
    Ctrl+Shift+2  - Prompt for custom annotation
    Ctrl+Shift+h  - Health check (verify Excel connection)
    Esc           - Quit the application
"""

import excel_ops
import data_fetcher
import keyboard
import time
import threading
import queue
import pythoncom
import tkinter as tk
from tkinter import simpledialog

# =============================================================================
# GLOBAL STATE
# =============================================================================

# Task queue for communication between hotkey handlers and worker thread
# Format: (mode, payload) where mode is "v1" or "v2", payload is optional prompt
task_queue = queue.Queue()

# Shutdown flag for graceful termination
shutdown_flag = threading.Event()


# =============================================================================
# WORKER THREAD
# =============================================================================

def worker_loop():
    """
    Persistent worker thread that processes Excel operations.
    
    IMPORTANT COM NOTES:
    1. CoInitialize() must be called at thread start
    2. CoUninitialize() must be called at thread end
    3. PumpWaitingMessages() keeps COM alive while idle (critical for tab switching!)
    
    The worker runs in a loop:
    1. Pump COM messages (keeps references fresh when user switches tabs)
    2. Check for tasks in queue
    3. Process task: get selection -> get context -> fetch data -> add note
    4. Repeat until shutdown
    """
    print("[Worker] Thread Started. Initializing COM...")
    pythoncom.CoInitialize()
    
    # Verify Excel connection at startup
    success, message = excel_ops.test_connection()
    if success:
        print(f"[Worker] Excel connection verified: {message}")
    else:
        print(f"[Worker] Warning: {message}")
        print("[Worker] The tool will still run - ensure Excel is open when using hotkeys.")
    
    while not shutdown_flag.is_set():
        try:
            # CRITICAL: Pump COM messages while waiting
            # This prevents stale references when user switches Excel tabs/workbooks
            pythoncom.PumpWaitingMessages()
            
            # Check for tasks (short timeout for responsive message pumping)
            try:
                task = task_queue.get(timeout=0.1)
            except queue.Empty:
                continue
            
            # Sentinel value signals shutdown
            if task is None:
                break
            
            mode, payload = task
            print(f"\n[Worker] Processing Task: {mode}")
            
            # --- CORE ANNOTATION LOGIC ---
            try:
                # Step 1: Get fresh Excel references
                app, book, sheet, selection = excel_ops.get_active_selection()
                if not selection:
                    print("[Worker] No active selection. Ensure Excel is open and a cell is selected.")
                    print("[Worker] Tip: Press Esc in Excel if you're editing a cell.")
                    task_queue.task_done()
                    continue

                # Step 2: Extract context from cell position
                context = excel_ops.get_context(selection)
                ticker = context.get("ticker", "UNKNOWN")
                period = context.get("time_period", "Current")
                line_item = context.get("line_item", "General")
                cell_addr = context.get("cell_address", "?")
                print(f"[Worker] Context: {ticker} | {period} | {line_item} | Cell: {cell_addr}")

                # Step 3: Fetch annotation content
                comments = data_fetcher.fetch_comments(ticker, period, line_item)
                
                # Step 4: Add custom prompt for V2 mode
                if mode == "v2" and payload:
                    comments += f"\n\n--- ANALYST PROMPT ---\nQ: {payload}\nA: (AI Generated Answer...)"

                # Step 5: Write comment to cell
                success = excel_ops.add_note_to_cell(selection, comments)
                if success:
                    print(f"[Worker] SUCCESS: Annotation added to {cell_addr}")
                else:
                    print("[Worker] FAILED: Could not add annotation.")
                
            except Exception as e:
                print(f"[Worker] Error: {e}")
            
            # Cooldown before next operation
            time.sleep(0.2)
            task_queue.task_done()
            print("[Worker] Ready for next annotation...")
            
        except Exception as e:
            print(f"[Worker] Fatal Loop Error: {e}")
            
    pythoncom.CoUninitialize()
    print("[Worker] Thread Stopped.")


# =============================================================================
# HOTKEY HANDLERS
# =============================================================================

def on_hotkey_v1():
    """
    Ctrl+Shift+m: Auto-Annotate
    Queues a V1 task (no user prompt, automatic context extraction).
    """
    print("\n-> Hotkey V1 Pressed (Ctrl+Shift+m)")
    task_queue.put(("v1", None))


def on_hotkey_v2():
    """
    Ctrl+Shift+2: Prompt + Annotate
    Shows a dialog for custom prompt, then queues a V2 task.
    
    Note: The dialog runs in a temporary thread to avoid blocking.
    """
    print("\n-> Hotkey V2 Pressed (Ctrl+Shift+2)")
    
    def ui_step():
        try:
            # Create temporary tkinter root for dialog
            dialog_root = tk.Tk()
            dialog_root.withdraw()
            dialog_root.attributes("-topmost", True)
            
            prompt = simpledialog.askstring("Axe Annotate", "Enter your prompt:", parent=dialog_root)
            dialog_root.destroy()
            
            if prompt:
                task_queue.put(("v2", prompt))
            else:
                print("[UI] Prompt cancelled or empty.")
        except Exception as e:
            print(f"[UI] Error: {e}")

    threading.Thread(target=ui_step).start()


def on_health_check():
    """
    Ctrl+Shift+h: Health Check
    Tests the Excel connection without modifying anything.
    """
    print("\n-> Health Check Requested (Ctrl+Shift+h)")
    
    def check_step():
        pythoncom.CoInitialize()
        success, message = excel_ops.test_connection()
        pythoncom.CoUninitialize()
        
        if success:
            print(f"[Health] CONNECTED: {message}")
        else:
            print(f"[Health] NOT READY: {message}")
    
    threading.Thread(target=check_step).start()


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    """
    Application entry point.
    Sets up worker thread, registers hotkeys, and waits for exit.
    """
    print("=" * 55)
    print("         Axe Annotate v2.2 (Clean Edition)")
    print("=" * 55)
    print()
    print("  SHORTCUTS:")
    print("  " + "-" * 45)
    print("  Ctrl+Shift+m   Auto-Annotate selected cell")
    print("  Ctrl+Shift+2   Custom Prompt + Annotate")
    print("  Ctrl+Shift+h   Check Excel connection")
    print("  Esc            Quit")
    print("  " + "-" * 45)
    print()
    print("  TIPS:")
    print("  * Open Excel before using hotkeys")
    print("  * Press Esc in Excel if editing a cell")
    print("  * Select target cell before pressing hotkeys")
    print()
    print("=" * 55)

    # Start worker thread (daemon=True means it dies with main thread)
    worker = threading.Thread(target=worker_loop, daemon=True)
    worker.start()

    # Register global hotkeys
    keyboard.add_hotkey('ctrl+shift+m', on_hotkey_v1)
    keyboard.add_hotkey('ctrl+shift+2', on_hotkey_v2)
    keyboard.add_hotkey('ctrl+shift+h', on_health_check)

    print("\n[Status] Ready and listening for hotkeys...\n")

    # Wait for Esc key to exit
    keyboard.wait('esc')
    
    # Graceful shutdown
    print("\n[Shutdown] Stopping...")
    shutdown_flag.set()
    task_queue.put(None)  # Sentinel to wake up worker
    worker.join(timeout=2)
    print("[Shutdown] Goodbye!")


if __name__ == "__main__":
    main()
