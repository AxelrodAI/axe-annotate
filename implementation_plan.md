# Axe Annotate - Implementation Plan

## Status: ✅ MVP Complete (v2.2)

## Overview
A tool for hedge fund analysts to automatically annotate Excel cells with insights from financial filings (10-Q, 8-K, Transcripts) based on the cell's context (Ticker, Time Period, Line Item).

## How It Works
1. **User selects a cell** in Excel (e.g., B2 containing Q1 2024 Revenue)
2. **User presses hotkey** (Ctrl+Shift+m for auto, Ctrl+Shift+2 for prompted)
3. **Tool extracts context**:
   - Ticker: From cell A1
   - Time Period: Searched UP in same column
   - Line Item: Searched LEFT in same row
4. **Tool fetches data** (currently mock, ready for real API integration)
5. **Tool adds cell comment** with the annotation

## Architecture

```
main.py          → Entry point, hotkeys, worker thread
excel_ops.py     → Excel COM operations (xlwings + win32com)
data_fetcher.py  → Data source (mock, replace for production)
tests/           → Debug and test scripts
```

## Tech Stack
- **Python 3.8+**
- **xlwings** - High-level Excel COM wrapper
- **pywin32** - Low-level COM access (fixes tab-switching issues)
- **keyboard** - Global hotkey detection

## Completed Features

### Phase 1: MVP ✅
- [x] Global hotkey listener (Ctrl+Shift+m, Ctrl+Shift+2, Ctrl+Shift+h)
- [x] Context extraction from Excel selection
- [x] Cell comment annotation
- [x] Mock data fetcher
- [x] Worker thread with COM message pumping
- [x] Retry logic with exponential backoff
- [x] Tab/workbook switching fix (win32com + stale ref detection)

### Phase 2: Reliability ✅
- [x] Fixed "hotkeys stop working after tab switch" bug
- [x] Non-interactive test mode (--auto flag) for CI/agents
- [x] Comprehensive debug/diagnostic tools
- [x] Clean codebase with documentation

## Future Enhancements

### Phase 3: Real Data Sources
- [ ] Connect to SEC EDGAR API for 10-K/10-Q filings
- [ ] Integrate earnings transcript APIs
- [ ] Add caching layer for fetched documents

### Phase 4: AI Enhancement
- [ ] GPT/Claude integration for smart summarization
- [ ] Semantic search across filings
- [ ] Custom prompts with AI-generated answers

### Phase 5: UI/UX
- [ ] System tray icon
- [ ] Configuration GUI
- [ ] Multiple annotation styles (comments, notes, sidebar)

## File Structure

```
Axe Annotate/
├── main.py              # Entry point
├── excel_ops.py         # Excel operations
├── data_fetcher.py      # Data source
├── requirements.txt     # Dependencies
├── run_axe_annotate.bat # Windows launcher
├── README.md            # User documentation
├── AGENTS.md            # AI agent documentation
├── DEBUG_HISTORY.md     # Bug history & solutions
├── implementation_plan.md # This file
└── tests/
    ├── run_tests.py           # Test runner
    ├── debug_excel.py         # Connection debugger
    ├── stress_test_excel.py   # Reliability test
    ├── test_queue.py          # Queue pattern test
    ├── diagnose_tab_switch.py # Tab switching test
    ├── diagnose_alt_tab.py    # Alt-tab focus test
    ├── edge_case_tests.py     # Edge case suite
    ├── test_logic.py          # Unit tests
    └── verify_connection.py   # Quick connection check
```
