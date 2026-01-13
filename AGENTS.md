# AGENTS.md - Documentation for AI Coding Agents

This file provides context and guidelines for AI agents working on this codebase.

> **IMPORTANT**: Before debugging any issues, read `DEBUG_HISTORY.md` first!
> It contains solutions to previously fixed bugs that may recur.

## Project Overview

**Axe Annotate** is a Windows tool that adds contextual annotations to Excel cells. It runs in the background, listens for keyboard shortcuts, and automatically adds comments to selected cells based on the cell's context (ticker symbol, time period, line item).

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                     main.py (Entry Point)                   │
│  - Registers keyboard hotkeys (Ctrl+Shift+m, etc.)          │
│  - Spawns worker thread for COM operations                  │
│  - Manages task queue between hotkey handlers and worker    │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                     excel_ops.py                            │
│  - All Excel/COM interactions via xlwings + win32com        │
│  - get_active_selection(): Gets fresh Excel references      │
│  - get_context(): Extracts ticker, period, line item        │
│  - add_note_to_cell(): Adds comments to cells               │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                     data_fetcher.py                         │
│  - Coordinate logic for data retrieval                      │
│  - Calls rag_ops.py for RAG pipeline                        │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│             rag_ops.py  &  edgar_ops.py                     │
│  - rag_ops: Search, Retrieve, and LLM Summarization         │
│  - edgar_ops: SEC.gov fetching (No API Key required)        │
└─────────────────────────────────────────────────────────────┘
```

## Critical Technical Details

### COM Threading (IMPORTANT!)

Windows COM (Component Object Model) is used to communicate with Excel. Key rules:

1. **COM must be initialized per-thread**: Call `pythoncom.CoInitialize()` at the start of any thread that uses COM, and `pythoncom.CoUninitialize()` at the end.

2. **Message pumping is required**: When a thread is idle but needs to keep COM references fresh, call `pythoncom.PumpWaitingMessages()` periodically. Without this, switching Excel tabs/workbooks can cause stale references.

3. **Always get fresh references**: Never cache `xw.apps.active`, `app.books.active`, etc. These can become stale. The `get_active_selection()` function handles this.

### Why Hotkeys Stop Working After Tab Switch

This was a major bug. The root cause:
- xlwings caches some internal references
- When user switches tabs/workbooks, these cached refs point to the old location
- Solution: Use `win32com.client.GetActiveObject()` for fresh COM refs, and verify selection sheet matches active sheet

### File Descriptions

| File | Purpose |
|------|---------|
| `rag_ops.py` | **RAG Pipeline**: Coordinates fetching + extracting + summarizing |
| `edgar_ops.py` | **SEC Client**: Fetches 10-Q/10-K text directly from SEC.gov |
| `debug_rag_pipeline.py` | **Diagnostics**: Tests web fetching & LLM in isolation |
|------|---------|
| `main.py` | Entry point. Hotkey registration, worker thread, task queue |
| `excel_ops.py` | All Excel COM interactions. Selection, context, comments |
| `data_fetcher.py` | Mock data source. Replace with real API calls |
| `run_axe_annotate.bat` | Windows launcher script |
| `requirements.txt` | Python dependencies |

### Test/Debug Files

| File | Purpose | Usage |
|------|---------|-------|
| `debug_excel.py` | Step-by-step Excel connection test | `python tests/debug_excel.py --auto` |
| `stress_test_excel.py` | Multi-iteration reliability test | `python tests/stress_test_excel.py --auto` |
| `test_queue.py` | Tests the worker queue pattern | `python tests/test_queue.py --auto` |
| `diagnose_tab_switch.py` | Tests tab/workbook switching | `python tests/diagnose_tab_switch.py --auto` |
| `diagnose_alt_tab.py` | Tests alt-tab between apps | `python tests/diagnose_alt_tab.py --auto` |

**Note**: All test scripts support `--auto` flag for non-interactive mode (required for AI agents).

## Common Issues & Solutions

### Issue: "No active selection" or similar errors
**Cause**: User is in Excel's Edit Mode (typing in a cell)
**Solution**: Press Esc in Excel before using hotkeys

### Issue: Hotkeys don't work after switching tabs within Excel
**Cause**: Stale COM references (fixed in current version)
**Solution**: The code now uses `win32com.client.GetActiveObject()` and message pumping

### Issue: Hotkeys don't work after alt-tabbing to other apps and back
**Cause**: Excel needs time to restore after focus change; COM references may be briefly invalid
**Solution**: The `_wait_for_excel_ready()` function now waits for Excel to be responsive before proceeding. The code uses `excel_api.Ready` property and validates workbook access.

### Issue: Script hangs when run by agent
**Cause**: `input()` prompts wait for user input
**Solution**: Use `--auto` flag on all test scripts

## Modifying the Code

### Adding a new hotkey
1. In `main.py`, create handler function: `def on_hotkey_new(): ...`
2. Register it: `keyboard.add_hotkey('ctrl+shift+x', on_hotkey_new)`
3. Add task to queue or run directly

### Changing data source
1. Edit `data_fetcher.py`
2. Replace `fetch_comments()` with real API calls
3. Keep the return format: string with formatted annotation text

### Debugging Excel issues
1. Run `python debug_excel.py --auto` to test connection
2. Run `python diagnose_tab_switch.py --auto` for tab-switch issues
3. Check that `pythoncom.CoInitialize()` is called in the thread

## Dependencies

- `xlwings`: Excel COM wrapper (high-level API)
- `pywin32`: Low-level Windows COM access (win32com.client)
- `keyboard`: Global hotkey detection
- `pythoncom`: COM threading support (part of pywin32)

## Running the Tool

```bash
# Option 1: Use the batch file (recommended)
run_axe_annotate.bat

# Option 2: Direct Python
python main.py
```

## Testing Changes

After making changes, verify with:
```bash
python debug_excel.py --auto         # Basic connection test
python stress_test_excel.py --auto   # Reliability test
python diagnose_tab_switch.py --auto # Tab switching test
```
