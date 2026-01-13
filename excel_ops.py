"""
Excel Operations Module
========================
Handles all Excel COM interactions via xlwings and win32com.

This module provides:
- get_active_selection(): Get fresh Excel app, book, sheet, and selection
- get_context(): Extract context (ticker, period, line item) from cell position
- add_note_to_cell(): Add a comment/note to a cell
- test_connection(): Verify Excel is accessible

IMPORTANT - COM Reference Freshness:
------------------------------------
When users switch Excel tabs or workbooks, cached references become stale.
This module uses multiple strategies to ensure fresh references:

1. win32com.client.GetActiveObject() - Bypasses xlwings caching
2. Direct API calls (app.api.ActiveSheet.Name) - Gets current state
3. Selection sheet verification - Detects and corrects stale selections
4. Retry logic with exponential backoff - Handles transient failures

All functions are designed to be called from a thread that has initialized COM
via pythoncom.CoInitialize().
"""

import xlwings as xw
import time

# =============================================================================
# CONFIGURATION
# =============================================================================

MAX_RETRIES = 3                # Number of retry attempts for COM operations
RETRY_DELAY_BASE = 0.3         # Base delay in seconds (uses exponential backoff)


# =============================================================================
# INTERNAL HELPERS
# =============================================================================

def _force_excel_refresh(app):
    """
    Forces Excel to update its internal state.
    
    This helps when:
    - New rows/columns have just been added
    - User has switched sheets/workbooks
    - Excel's internal state is out of sync
    
    Args:
        app: xlwings App object
    """
    try:
        # Toggle ScreenUpdating to force refresh
        app.api.ScreenUpdating = False
        app.api.ScreenUpdating = True
        
        # Recalculate formulas
        try:
            app.api.Calculate()
        except Exception:
            pass
            
    except Exception as e:
        # Non-fatal, continue anyway
        print(f"[Excel] Note: Refresh failed ({e}), continuing...")


def _is_likely_label(value):
    """
    Determines if a cell value is likely a label/header vs numeric data.
    
    Returns True for text labels, False for numbers/empty cells.
    Used in context extraction to find headers.
    """
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return False
    s_val = str(value).strip()
    if s_val == "":
        return False
    # Check if it's a formatted number (e.g., "$1,234", "50%")
    try:
        float(s_val.replace(",", "").replace("$", "").replace("%", ""))
        return False
    except ValueError:
        return True


def _safe_read_cell(sheet, row, col):
    """
    Safely reads a cell value, returning None on any error.
    
    Args:
        sheet: xlwings Sheet object
        row: 1-indexed row number
        col: 1-indexed column number
    
    Returns:
        Cell value or None if error
    """
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


# =============================================================================
# PUBLIC API
# =============================================================================

def _wait_for_excel_ready(excel_api, timeout=2.0):
    """
    Waits for Excel to be in a ready state.
    
    After alt-tabbing, Excel might need a moment to fully restore.
    This function waits until Excel responds properly.
    
    Args:
        excel_api: COM Excel.Application object
        timeout: Maximum wait time in seconds
    
    Returns:
        bool: True if Excel is ready, False if timeout
    """
    start_time = time.time()
    
    while (time.time() - start_time) < timeout:
        try:
            # Try multiple checks to ensure Excel is responsive
            _ = excel_api.Version
            _ = excel_api.Ready  # Excel's Ready property
            
            # Check if we can access the active workbook
            wb = excel_api.ActiveWorkbook
            if wb is not None:
                _ = wb.Name
                return True
                
        except Exception:
            pass
        
        time.sleep(0.1)
    
    return False


def get_active_selection(max_retries=MAX_RETRIES):
    """
    Gets fresh references to the active Excel app, workbook, sheet, and selection.
    
    CRITICAL: This function handles multiple problematic scenarios:
    1. Tab switching within Excel (switching sheets/workbooks)
    2. Alt-tabbing between applications (Excel loses and regains focus)
    3. Stale COM references after extended idle periods
    
    Strategy:
    1. Use win32com.GetActiveObject() for truly fresh Excel connection
    2. Wait for Excel to be in ready state (handles alt-tab recovery)
    3. Match to xlwings App by window handle (Hwnd)
    4. Get active workbook/sheet by NAME from API (not .active property)
    5. Verify selection is on the correct sheet, correct if needed
    6. Retry with exponential backoff on failures
    
    Returns:
        tuple: (app, book, sheet, selection) or (None, None, None, None) on failure
    """
    last_error = None
    
    for attempt in range(max_retries):
        try:
            # --- Step 1: Get fresh Excel reference via win32com ---
            # This bypasses xlwings internal caching
            excel_api = None
            try:
                import win32com.client as win32
                
                # Use GetActiveObject to get the running Excel instance
                excel_api = win32.GetActiveObject("Excel.Application")
                
                # CRITICAL: Wait for Excel to be ready after potential alt-tab
                if not _wait_for_excel_ready(excel_api, timeout=1.0):
                    # Excel might be busy, give it more time and retry
                    print("[Excel] Waiting for Excel to be ready...")
                    time.sleep(0.3)
                    if not _wait_for_excel_ready(excel_api, timeout=1.0):
                        raise ConnectionError("Excel is not responding. Please try again.")
                
                # Verify workbook exists
                if excel_api.ActiveWorkbook is None:
                    raise ConnectionError("No active workbook. Please open a workbook.")
                if excel_api.ActiveSheet is None:
                    raise ConnectionError("No active sheet found.")
                    
                # Find matching xlwings App by window handle
                target_hwnd = excel_api.Hwnd
                app = None
                for xw_app in xw.apps:
                    try:
                        if xw_app.api.Hwnd == target_hwnd:
                            app = xw_app
                            break
                    except Exception:
                        continue
                
                if app is None:
                    # Fallback: use active app
                    if len(xw.apps) > 0:
                        app = xw.apps.active
                    else:
                        raise ConnectionError("No xlwings app found.")
                    
            except ImportError:
                # win32com not available
                if len(xw.apps) == 0:
                    raise ConnectionError("No Excel running. Please open Excel first.")
                app = xw.apps.active
            except Exception as e:
                # win32com failed, use xlwings fallback
                if len(xw.apps) == 0:
                    raise ConnectionError("No Excel running. Please open Excel first.")
                app = xw.apps.active
                if "GetActiveObject" not in str(e) and "not responding" not in str(e):
                    print(f"[Excel] Using xlwings fallback ({type(e).__name__}: {e})")
            
            if app is None:
                raise ConnectionError("No active Excel application found.")
            
            # --- Step 2: Force refresh and verify ready state ---
            _force_excel_refresh(app)
            time.sleep(0.05)
            
            # Verify Excel is responsive
            try:
                _ = app.api.Version
            except Exception as e:
                raise ConnectionError(f"Excel busy or in Edit Mode. Press Esc first. ({e})")
            
            # --- Step 3: Get workbook by name (not .active) ---
            try:
                # Use API directly for most current state
                api_book = app.api.ActiveWorkbook
                if api_book is None:
                    raise ConnectionError("No active workbook.")
                book_name = api_book.Name
                book = app.books[book_name]
            except KeyError:
                # Book not in xlwings cache, try to get it
                book = app.books.active
            except Exception as e:
                book = app.books.active
                if book is None:
                    raise ConnectionError(f"Cannot access workbook: {e}")
            
            if book is None:
                raise ConnectionError("No active workbook.")
            
            # --- Step 4: Get sheet by name (not .active) ---
            try:
                api_sheet = app.api.ActiveSheet
                if api_sheet is None:
                    raise ConnectionError("No active sheet.")
                sheet_name = api_sheet.Name
                sheet = book.sheets[sheet_name]
            except KeyError:
                sheet = book.sheets.active
            except Exception as e:
                sheet = book.sheets.active
                if sheet is None:
                    raise ConnectionError(f"Cannot access sheet: {e}")
                
            if sheet is None:
                raise ConnectionError("No active sheet.")
            
            # --- Step 5: Get selection with stale reference detection ---
            try:
                # Try xlwings selection first
                selection = app.selection
                addr = selection.address
                row = selection.row
                col = selection.column
                
                # CRITICAL: Check if selection is on a different sheet (stale ref)
                try:
                    sel_sheet = selection.sheet.name
                    actual_sheet = app.api.ActiveSheet.Name
                    
                    if sel_sheet != actual_sheet:
                        print(f"[Excel] Stale selection detected! Correcting...")
                        print(f"[Excel]   Selection was on: {sel_sheet}")
                        print(f"[Excel]   Active sheet is: {actual_sheet}")
                        
                        # Get correct selection via API
                        api_sel = app.api.Selection
                        if api_sel is not None:
                            row = api_sel.Row
                            col = api_sel.Column
                            selection = sheet.range((row, col))
                            addr = selection.address
                        print(f"[Excel] Corrected selection: {addr}")
                    else:
                        print(f"[Excel] Selection: {addr} (Row {row}, Col {col})")
                except Exception as sheet_check_error:
                    # Sheet check failed but we have a selection, use it
                    print(f"[Excel] Sheet verification skipped ({sheet_check_error})")
                    print(f"[Excel] Selection: {addr} (Row {row}, Col {col})")
                    
            except Exception as e:
                # xlwings selection failed, try direct API
                try:
                    api_sel = app.api.Selection
                    if api_sel is None:
                        raise ConnectionError("No selection in Excel.")
                    row = api_sel.Row
                    col = api_sel.Column
                    selection = sheet.range((row, col))
                    addr = selection.address
                    print(f"[Excel] Selection (via API fallback): {addr}")
                except Exception as api_e:
                    raise ConnectionError(f"Cannot read selection: {e}. API fallback also failed: {api_e}")
            
            return app, book, sheet, selection
            
        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                delay = RETRY_DELAY_BASE * (2 ** attempt)
                print(f"[Excel] Attempt {attempt + 1} failed: {e}. Retry in {delay:.1f}s...")
                time.sleep(delay)
            else:
                print(f"[Excel] All {max_retries} attempts failed. Error: {e}")
    
    return None, None, None, None


def get_context(selection):
    """
    Extracts context from the cell's position in the spreadsheet.
    
    Assumes a typical financial spreadsheet layout:
    - Row 1: Time period headers (Q1 2024, Q2 2024, etc.)
    - Column A: Line item labels (Revenue, Net Income, etc.)
    - Cell A1: Ticker symbol (AAPL, MSFT, etc.)
    
    Search Strategy:
    - Time Period: Search UP in the same column for first text label
    - Line Item: Search LEFT in the same row for first text label
    - Ticker: Read cell A1
    
    Args:
        selection: xlwings Range object
    
    Returns:
        dict with keys: ticker, time_period, line_item, cell_address
    """
    if not selection:
        return {}

    try:
        sheet = selection.sheet
        row_idx = selection.row
        col_idx = selection.column
        cell_addr = selection.address
    except Exception as e:
        print(f"[Context] Error reading selection: {e}")
        return {
            "ticker": "UNKNOWN",
            "time_period": "Unknown Period",
            "line_item": "Unknown Item",
            "cell_address": "?"
        }

    # Search LEFT for Line Item
    line_item = None
    for c in range(col_idx - 1, 0, -1):
        val = _safe_read_cell(sheet, row_idx, c)
        if _is_likely_label(val):
            line_item = str(val).strip()
            print(f"[Context] Line item found in col {c}: '{line_item}'")
            break
    if not line_item:
        line_item = "Unknown Line Item"

    # Search UP for Time Period
    time_period = None
    for r in range(row_idx - 1, 0, -1):
        val = _safe_read_cell(sheet, r, col_idx)
        if _is_likely_label(val):
            time_period = str(val).strip()
            print(f"[Context] Time period found in row {r}: '{time_period}'")
            break
    if not time_period:
        time_period = "Unknown Period"

    # Ticker from A1
    ticker = "UNKNOWN"
    ticker_val = _safe_read_cell(sheet, 1, 1)
    if ticker_val and _is_likely_label(ticker_val):
        ticker = str(ticker_val).strip()
    
    print(f"[Context] Result: Ticker='{ticker}', Period='{time_period}', Item='{line_item}'")

    return {
        "ticker": ticker,
        "time_period": time_period,
        "line_item": line_item,
        "cell_address": cell_addr
    }


def add_note_to_cell(selection, note_text, max_retries=MAX_RETRIES):
    """
    Adds a comment/note to the selected cell.
    
    Handles existing comments by clearing them first.
    If a range is selected, adds comment to the first cell.
    Uses retry logic for reliability.
    
    Args:
        selection: xlwings Range object (can be single cell or range)
        note_text: String content for the comment
        max_retries: Number of retry attempts
    
    Returns:
        bool: True on success, False on failure
    """
    if not selection:
        print("[Excel] Cannot add note: No selection provided.")
        return False

    for attempt in range(max_retries):
        try:
            cell_api = selection.api
            
            # Handle multi-cell ranges: comments can only be added to single cells
            # Get the first cell in the range
            try:
                # Check if it's a multi-cell selection
                if selection.count > 1:
                    # Get just the first cell
                    first_cell = selection[0, 0]  # Top-left cell
                    cell_api = first_cell.api
                    print(f"[Excel] Multi-cell selection detected, using first cell: {first_cell.address}")
            except Exception:
                # If count fails, just use the selection as-is
                pass
            
            # Clear existing comment
            try:
                cell_api.ClearComments()
            except Exception:
                try:
                    if cell_api.Comment:
                        cell_api.Comment.Delete()
                except Exception:
                    pass
            
            time.sleep(0.05)
            
            # Add new comment
            cell_api.AddComment(note_text)
            
            time.sleep(0.05)
            return True
            
        except Exception as e:
            if attempt < max_retries - 1:
                delay = RETRY_DELAY_BASE * (2 ** attempt)
                print(f"[Excel] Note failed (attempt {attempt + 1}): {e}. Retry in {delay:.1f}s...")
                time.sleep(delay)
            else:
                print(f"[Excel] Failed to add note after {max_retries} attempts: {e}")
    
    return False


def test_connection():
    """
    Quick health check to verify Excel connection.
    
    Returns:
        tuple: (success: bool, message: str)
    """
    try:
        if len(xw.apps) == 0:
            return False, "No Excel running. Please open Excel."
        
        app = xw.apps.active
        if app is None:
            return False, "No active Excel instance."
        
        try:
            version = app.api.Version
        except Exception as e:
            return False, f"Excel not responding: {e}"
        
        try:
            book = app.books.active
            if book is None:
                return False, "No workbook open."
            workbook_name = book.name
        except Exception:
            return False, "Cannot access workbook."
        
        try:
            sheet_name = book.sheets.active.name
        except Exception:
            sheet_name = "?"
        
        return True, f"Excel {version} - '{workbook_name}' / {sheet_name}"
        
    except Exception as e:
        return False, f"Connection error: {e}"
