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

from rag_ops import rag

def fetch_comments(ticker: str, period: str, line_item: str) -> str:
    """
    Fetches contextual comments using the RAG pipeline.
    
    Workflow:
    1. Search for transcript/filing URL (rag.find_transcript_url)
    2. Scrape content via Firecrawl (rag.fetch_content)
    3. Retrieve relevant sections (rag.retrieve_context)
    4. Format for display
    
    Args:
        ticker: Stock symbol (e.g., "AAPL", "MSFT")
        period: Time period (e.g., "Q1 2024", "FY 2023")
        line_item: Financial metric (e.g., "Revenue", "Net Income")
    
    Returns:
        Formatted string with annotation content
    """
    try:
        print(f"[DataFetcher] RAG Fetch: {ticker} | {period} | {line_item}")
        
        # 1. Fetch Content
        try:
            content = rag.get_filing_content(ticker, period)
        except Exception as e:
            return f"Error Fetching Filing: {str(e)}"
            
        if not content:
            return f"No filing data found for {ticker} ({period})."
        
        # 2. Retrieve Context
        # If line_item is unknown/generic, use broader terms
        query = line_item if line_item and line_item != "Unknown Line Item" else "Financial Highlights"
        
        try:
            raw_insights = rag.retrieve_context(content, query)
        except Exception as e:
            raw_insights = f"Error extracting context: {e}"
        
        # 3. Summarize with LLM (with timeout/fail safety)
        summary = "Summary unavailable (Time out or Error)."
        try:
            summary = rag.summarize_context(raw_insights, query)
        except Exception as e:
            summary = f"AI Summary Failed: {e}\n\nRaw Context:\n{raw_insights[:500]}..."
        
        # 4. Format Output
        formatted = f"--- AXE KEY INSIGHTS ---\n"
        formatted += f"Target: {ticker} | Period: {period}\n"
        formatted += f"Topic: {query}\n"
        formatted += f"Source: 10-Q/K (AI Summarized)\n\n"
        formatted += summary
        
        return formatted

    except Exception as e:
        print(f"[DataFetcher] Fatal Error: {e}")
        return f"System Error: {str(e)}"
