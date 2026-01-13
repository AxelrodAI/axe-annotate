# DEBUG_HISTORY.md - Debugging History and Solutions

This document chronicles the bugs found and fixed in Axe Annotate, along with their root causes and solutions. Future agents should consult this before debugging similar issues.

---

## Table of Contents
1. [Tab Switching Bug](#1-tab-switching-bug)
2. [Alt-Tab Focus Loss Bug](#2-alt-tab-focus-loss-bug)
3. [Multi-Cell Selection Bug](#3-multi-cell-selection-bug)
4. [Agent Hang Bug](#4-agent-hang-bug)
5. [Edge Cases Tested](#5-edge-cases-tested)

---

## 1. Tab Switching Bug

### Symptom
Hotkeys stop working after user switches between sheets/workbooks within Excel.

### Root Cause
xlwings internally caches COM references. When the user switches to a different sheet or workbook:
- `xw.apps.active` might still point to the old reference
- `app.selection` might return the selection from the previous sheet
- The stale reference causes operations to fail silently or target the wrong cell

### Solution
1. **Use `win32com.client.GetActiveObject()`** to bypass xlwings caching and get a truly fresh Excel reference.

2. **Match xlwings App by window handle** (`Hwnd`) to ensure we're using the correct Excel instance.

3. **Get workbook/sheet by NAME** from the COM API rather than using `.active` properties:
   ```python
   book_name = app.api.ActiveWorkbook.Name
   book = app.books[book_name]
   ```

4. **Detect stale selections** by comparing the selection's sheet to the active sheet:
   ```python
   sel_sheet = selection.sheet.name
   actual_sheet = app.api.ActiveSheet.Name
   if sel_sheet != actual_sheet:
       # Correct the selection
   ```

5. **Pump COM messages** in the worker loop while idle:
   ```python
   pythoncom.PumpWaitingMessages()
   ```

### Files Modified
- `excel_ops.py`: `get_active_selection()` function
- `main.py`: `worker_loop()` function

### Test Script
```bash
python tests/diagnose_tab_switch.py --auto
```

---

## 2. Alt-Tab Focus Loss Bug

### Symptom
Hotkeys stop working after user alt-tabs to another application (browser, VS Code, etc.) and alt-tabs back to Excel.

### Root Cause
When Excel loses and regains focus:
- Excel needs a moment to fully restore its state
- COM references may be briefly invalid
- `ActiveWorkbook` might temporarily return `None`

### Solution
Added `_wait_for_excel_ready()` function that:
1. Checks `excel_api.Ready` property
2. Validates `excel_api.Version` is accessible
3. Confirms `ActiveWorkbook` is not None and accessible
4. Retries every 100ms up to a configurable timeout (default 2 seconds)

```python
def _wait_for_excel_ready(excel_api, timeout=2.0):
    """Waits for Excel to be in a ready state after focus change."""
    start_time = time.time()
    while (time.time() - start_time) < timeout:
        try:
            _ = excel_api.Version
            _ = excel_api.Ready
            wb = excel_api.ActiveWorkbook
            if wb is not None:
                _ = wb.Name
                return True
        except Exception:
            pass
        time.sleep(0.1)
    return False
```

### Files Modified
- `excel_ops.py`: Added `_wait_for_excel_ready()`, updated `get_active_selection()`

### Test Script
```bash
python tests/diagnose_alt_tab.py --auto
```

---

## 3. Multi-Cell Selection Bug

### Symptom
Tool fails when user selects a range of cells instead of a single cell.

### Root Cause
Excel's `AddComment` method only works on single cells. When a range is selected:
- `selection.api.AddComment()` fails with a COM exception
- Error code: `-2147352567` (Invalid parameter)

### Solution
Detect multi-cell selection and use only the first cell:
```python
if selection.count > 1:
    first_cell = selection[0, 0]  # Top-left cell
    cell_api = first_cell.api
    print(f"[Excel] Multi-cell selection, using: {first_cell.address}")
```

### Files Modified
- `excel_ops.py`: `add_note_to_cell()` function

### Test Script
```bash
python tests/edge_case_tests.py --auto
```

---

## 4. Agent Hang Bug

### Symptom
AI agents get stuck with "Sending termination request to com" when running debug scripts.

### Root Cause
Debug and test scripts had interactive `input()` prompts that wait for user input. AI agents cannot type into these prompts, causing the script to hang indefinitely.

### Solution
Added `--auto` flag to all test scripts:
```python
parser.add_argument('--auto', action='store_true', 
                    help='Run in non-interactive mode')
args = parser.parse_args()

if not args.auto:
    input("Press Enter to continue...")
else:
    print("[Auto Mode] Skipping input prompt...")
```

### Files Modified
- All files in `tests/` directory

### Usage
```bash
# For agents/automation:
python tests/debug_excel.py --auto

# For humans (interactive):
python tests/debug_excel.py
```

---

## 5. Edge Cases Tested

The following edge cases have been tested and verified working:

| Edge Case | Status | Notes |
|-----------|--------|-------|
| Multiple cell selection (range) | ✅ PASS | Uses first cell in range |
| Empty context (no headers) | ✅ PASS | Returns default values |
| Existing comment overwrite | ✅ PASS | Properly clears and replaces |
| Large comments (10k+ chars) | ✅ PASS | Excel handles up to ~32k chars |
| Rapid successive operations | ✅ PASS | 10/10 succeeded with 50ms delay |
| Boundary cells (A1) | ✅ PASS | Handles missing headers gracefully |
| Multiple Excel instances | ✅ PASS | Uses correct instance via Hwnd |
| Excel minimized | ✅ PASS | Can operate while minimized |
| Special characters | ✅ PASS | Unicode, symbols, newlines work |
| Edit mode detection | ✅ PASS | Detects when user is editing |

### Test Script
```bash
python tests/edge_case_tests.py --auto
```

---

## Diagnostic Commands Reference

| Command | Purpose |
|---------|---------|
| `python tests/debug_excel.py --auto` | Basic connection test |
| `python tests/stress_test_excel.py --auto` | Reliability test (5 iterations) |
| `python tests/test_queue.py --auto` | Worker queue pattern test |
| `python tests/diagnose_tab_switch.py --auto` | Tab switching test |
| `python tests/diagnose_alt_tab.py --auto` | Alt-tab focus test |
| `python tests/edge_case_tests.py --auto` | Edge case suite (10 tests) |

---

## Common Error Patterns

### Error: "No active selection"
**Cause**: User is in Excel's Edit Mode (cursor blinking in cell)
**Solution**: Press Esc in Excel, then retry

### Error: "Excel is busy or in Edit Mode"
**Cause**: Excel is processing or user is editing
**Solution**: Wait a moment, press Esc, retry

### Error: COM Exception -2147352567
**Cause**: Usually multi-cell selection for single-cell operation
**Solution**: Fixed in code - automatically uses first cell

### Error: "Call was rejected by callee"
**Cause**: Excel is showing a dialog (Save As, Print, etc.)
**Solution**: Close the dialog, retry

### Error: Stale selection detected
**Cause**: User switched tabs and xlwings cached old selection
**Solution**: Fixed in code - automatically corrects selection

---

## Architecture Notes for Future Agents

### Threading Model
```
Main Thread (keyboard listener)
    │
    ├─► Hotkey detected → Put task in queue
    │
Worker Thread (COM operations)
    │
    ├─► CoInitialize() at start
    ├─► PumpWaitingMessages() while idle
    ├─► Get task from queue
    ├─► get_active_selection() with fresh refs
    ├─► get_context() for headers
    ├─► fetch_comments() for data
    ├─► add_note_to_cell() to write
    └─► CoUninitialize() at shutdown
```

### Key Invariants
1. **Never cache COM objects** - Always get fresh references
2. **Pump messages while idle** - Prevents stale references
3. **Verify sheet matches** - Selection might be from wrong sheet
4. **Handle multi-cell selections** - Use first cell only
5. **Retry with backoff** - Transient failures are common

---

*Last updated: 2026-01-13*
*Version: 2.2 (Clean Edition)*
