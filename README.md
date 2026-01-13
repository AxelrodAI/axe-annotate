# Axe Annotate

A Windows tool for Hedge Fund Analysts to automatically annotate Excel cells with contextual comments from SEC filings and earnings transcripts.

## Features

| Hotkey | Description |
|--------|-------------|
| `Ctrl+Shift+m` | **Auto-Annotate**: Reads context from cell position and adds annotation |
| `Ctrl+Shift+2` | **Prompt-Annotate**: Opens dialog for custom prompt before annotating |
| `Ctrl+Shift+h` | **Health Check**: Verifies Excel connection |
| `Esc` | **Quit**: Exit the application |

## Quick Start

### Option 1: Batch File (Recommended)
```bash
run_axe_annotate.bat
```

### Option 2: Manual
```bash
pip install -r requirements.txt
python main.py
```

## How It Works

1. **Open Excel** with a workbook containing financial data
2. **Set up your spreadsheet**:
   - Cell A1: Ticker symbol (e.g., "AAPL")
   - Row 1: Time periods (e.g., "Q1 2024", "Q2 2024")
   - Column A: Line items (e.g., "Revenue", "Net Income")
3. **Select a data cell** (e.g., B2 for Q1 2024 Revenue)
4. **Press `Ctrl+Shift+m`**
5. Watch the annotation appear as a cell comment!

### Example Layout

|     | A       | B        | C        |
|-----|---------|----------|----------|
| 1   | AAPL    | Q1 2024  | Q2 2024  |
| 2   | Revenue | 50,000   | 52,000   |
| 3   | Net Inc | 10,000   | 11,000   |

Selecting B2 and pressing `Ctrl+Shift+m` will:
- Detect Ticker: AAPL (from A1)
- Detect Period: Q1 2024 (from B1)
- Detect Item: Revenue (from A2)
- Add a comment with relevant insights

## Project Structure

```
Axe Annotate/
├── main.py           # Entry point, hotkeys, worker thread
├── excel_ops.py      # Excel COM operations
├── data_fetcher.py   # Data source (mock, replace for production)
├── requirements.txt  # Python dependencies
├── run_axe_annotate.bat  # Windows launcher
├── AGENTS.md         # Documentation for AI agents
├── README.md         # This file
└── tests/            # Debug and test scripts
    ├── debug_excel.py
    ├── stress_test_excel.py
    ├── test_queue.py
    └── diagnose_tab_switch.py
```

## Troubleshooting

### Hotkeys don't work
- Make sure you're not in Excel's Edit Mode (press Esc in Excel first)
- Verify Excel is open with a workbook
- Run health check: `Ctrl+Shift+h`

### Hotkeys stop working after switching tabs
This was fixed in v2.2. If you experience this:
1. Run `python diagnose_tab_switch.py --auto` to diagnose
2. The fix uses `win32com.client.GetActiveObject()` for fresh references

### Annotations appear on wrong cell
- Make sure you select the cell BEFORE pressing the hotkey
- Wait a moment after switching tabs before using hotkeys

## Configuration

### Changing Hotkeys
Edit `main.py` and modify:
```python
keyboard.add_hotkey('ctrl+shift+m', on_hotkey_v1)
keyboard.add_hotkey('ctrl+shift+2', on_hotkey_v2)
```

### Connecting to Real Data Sources
Edit `data_fetcher.py` and replace `fetch_comments()` with API calls to:
- SEC EDGAR for 10-K/10-Q filings
- Earnings transcript APIs
- Your internal data sources

## Requirements

- Windows 10/11
- Python 3.8+
- Microsoft Excel (desktop version)

## Dependencies

- `xlwings` - Excel COM wrapper
- `pywin32` - Windows COM support
- `keyboard` - Global hotkey detection

## For AI Agents

See `AGENTS.md` for detailed technical documentation on:
- Architecture and COM threading
- Common issues and solutions
- How to modify the codebase
- Debug and test scripts

## License

Internal tool - not for distribution.
