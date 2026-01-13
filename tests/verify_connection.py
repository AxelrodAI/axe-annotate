import xlwings as xw
import sys

def verify_excel_state():
    print("Connecting to Active Excel Instance...")
    try:
        app = xw.apps.active
        book = app.books.active
        sheet = book.sheets.active
        selection = app.selection
        
        print(f"SUCCESS: Connected to '{book.name}'")
        print(f"Active Sheet: '{sheet.name}'")
        
        # Read user specified cells
        val_a1 = sheet.range("A1").value
        val_b1 = sheet.range("B1").value
        val_a2 = sheet.range("A2").value
        
        print(f"\n[Current State Check]")
        print(f"Cell A1 (Ticker)      : {val_a1}")
        print(f"Cell B1 (Time Period) : {val_b1}")
        print(f"Cell A2 (Line Item)   : {val_a2}")
        
        print(f"\n[Selection]")
        if selection:
            print(f"Address: {selection.address}")
            print(f"Value  : {selection.value}")
            
            # Verify context logic
            target_col = selection.column
            target_row = selection.row
            
            # Replicate get_context logic
            ctx_period = sheet.range((1, target_col)).value
            ctx_line = sheet.range((target_row, 1)).value
            
            print(f"\n[Inferred Context for Selection]")
            print(f"Target Period : {ctx_period}")
            print(f"Target Line   : {ctx_line}")
        else:
            print("No cell selected.")

    except Exception as e:
        print(f"\n[ERROR] Could not connect to Excel: {e}")
        print("Make sure an Excel file is OPEN and not in a modal state (e.g., editing a cell).")

if __name__ == "__main__":
    verify_excel_state()
