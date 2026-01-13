"""
Data Fetcher Module
====================
Fetches annotation content for a given context (ticker, period, line item).

Currently uses MOCK DATA for demonstration.
Replace fetch_comments() with real API calls for production use.

Potential Data Sources:
- SEC EDGAR API: 10-K, 10-Q filings
- Transcript APIs: Earnings call transcripts  
- AlphaVantage/Yahoo Finance: Financial data
- Custom internal databases
"""

import time


def fetch_comments(ticker: str, period: str, line_item: str) -> str:
    """
    Fetches contextual comments for a given financial data point.
    
    In production, this would:
    1. Search SEC EDGAR for the 10-Q/10-K matching the period
    2. Search transcript APIs for earnings call mentions
    3. Extract relevant snippets using NLP/keyword matching
    
    Args:
        ticker: Stock symbol (e.g., "AAPL", "MSFT")
        period: Time period (e.g., "Q1 2024", "FY 2023")
        line_item: Financial metric (e.g., "Revenue", "Net Income")
    
    Returns:
        Formatted string with annotation content
    """
    print(f"[DataFetcher] Fetching: {ticker} | {period} | {line_item}")
    
    # Simulate network delay
    time.sleep(0.5)
    
    # Mock dataset - keyed by lowercase line item
    mock_data = {
        "revenue": [
            f"{ticker} reported strong revenue growth in {period} driven by demand.",
            "Segment A contributed 60% of total sales."
        ],
        "net income": [
            "Net income was impacted by one-time tax charges.",
            "Operational efficiency improved margins quarter-over-quarter."
        ],
        "net interest income": [
            f"Net interest income (NII) was a key driver for {ticker} in {period}.",
            "Higher rates supported NII expansion, offset by deposit costs.",
            "Management expects NII to stabilize in coming quarters."
        ],
        "eps": [
            f"Earnings per share for {period} exceeded analyst expectations.",
            "Share buybacks contributed to EPS growth."
        ],
        "operating expenses": [
            "Operating expenses were well-controlled despite inflation.",
            "Technology investments increased as planned."
        ]
    }
    
    # Case-insensitive lookup
    line_item_lower = (line_item or "").lower()
    
    if line_item_lower in mock_data:
        comments = mock_data[line_item_lower]
    else:
        # Generic fallback for unknown items
        company = ticker if ticker != "UNKNOWN" else "The company"
        comments = [
            f"Analyst Note: {line_item} for {company} in {period} aligns with consensus.",
            f"Management discussed {line_item} trends during the {period} earnings call.",
            f"See 10-Q {period} for detailed breakdown of {line_item}."
        ]
    
    # Format the output
    formatted = f"--- AXE ANNOTATE ---\n"
    formatted += f"Ticker: {ticker} | Period: {period}\n"
    formatted += f"Item: {line_item}\n"
    formatted += f"Source: Automated Insight\n\n"
    
    for comment in comments:
        formatted += f"• {comment}\n"
        
    return formatted
