import unittest
from unittest.mock import MagicMock
import data_fetcher
import excel_ops

class TestAxeAnnotate(unittest.TestCase):
    
    def test_data_fetcher(self):
        # Test that data fetcher returns a string with the ticker and period
        ticker = "AAPL"
        period = "Q1 2024"
        line_item = "Revenue"
        
        result = data_fetcher.fetch_comments(ticker, period, line_item)
        print("\n[Test Data Fetcher] Result:\n", result)
        
        self.assertIn("AAPL", result)
        self.assertIn("Q1 2024", result)
        self.assertIn("Revenue", result)

    def test_context_extraction(self):
        # Mock xlwings selection object
        mock_selection = MagicMock()
        mock_sheet = MagicMock()
        mock_selection.sheet = mock_sheet
        mock_selection.row = 2
        mock_selection.column = 2 # Column B
        mock_selection.address = "$B$2"

        # Setup mock return values for context cells
        # Header (Time Period) at (1, 2)
        mock_sheet.range.side_effect = lambda *args: MagicMock(value="Q1 2024") if args[0] == (1, 2) else MagicMock(value="Revenue")
        
        # We need to handle the specific calls: range((1, col)), range((row, 1))
        def range_side_effect(coords):
            row, col = coords
            val = None
            if row == 1 and col == 2:
                val = "Q1 2024"
            elif row == 2 and col == 1:
                val = "Revenue"
            return MagicMock(value=val)

        mock_sheet.range.side_effect = range_side_effect

        context = excel_ops.get_context(mock_selection)
        print("\n[Test Context Extraction] Result:\n", context)
        
        self.assertEqual(context['time_period'], "Q1 2024")
        self.assertEqual(context['line_item'], "Revenue")

if __name__ == '__main__':
    unittest.main()
