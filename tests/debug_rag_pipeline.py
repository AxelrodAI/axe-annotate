"""
RAG Pipeline Debugger
=====================
Diagnose issues with SEC fetching and LLM summarization independently.
"""
import sys
import os
import time

# Add parent dir to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import rag_ops
import edgar_ops

def test_edgar_connection(ticker="AAPL"):
    print(f"\n--- Testing SEC EDGAR Connection ({ticker}) ---")
    start = time.time()
    try:
        cik = edgar_ops.get_cik_from_ticker(ticker)
        print(f"1. Ticker Resolve: OK (CIK: {cik}) - {time.time()-start:.2f}s")
        
        if not cik: 
            print("FAIL: Could not resolve ticker.")
            return None
            
        text = edgar_ops.get_latest_filing_text(ticker, "10-Q")
        if text:
            print(f"2. Fetch 10-Q: OK ({len(text)} chars) - {time.time()-start:.2f}s")
            return text
        else:
            print("FAIL: Could not fetch filing text.")
            return None
    except Exception as e:
        print(f"FAIL: Exception during EDGAR fetch: {e}")
        return None

def test_llm_summarization(text, kpi="Revenue"):
    print(f"\n--- Testing LLM Summarization ({kpi}) ---")
    if not text:
        print("SKIP: No text to summarize.")
        return

    # Simulate retrieval first
    retrieved = rag_ops.rag.retrieve_context(text, kpi)
    print(f"1. Retrieval: Found {len(retrieved)} chars of context.")
    if "No specific comments" in retrieved:
        print("WARN: Retrieval found nothing relevant.")
    
    start = time.time()
    try:
        summary = rag_ops.rag.summarize_context(retrieved, kpi)
        print(f"2. Pollinations API: Returned {len(summary)} chars - {time.time()-start:.2f}s")
        print("\n=== SUMMARY OUTPUT ===")
        # Safe print for Windows console
        print(summary.encode('cp1252', errors='replace').decode('cp1252'))
        print("======================")
    except Exception as e:
        print(f"FAIL: LLM call failed: {e}")

if __name__ == "__main__":
    ticker = "NVDA" # Use a different ticker to test robustness
    print(f"Diagnosing RAG for {ticker}...")
    
    text = test_edgar_connection(ticker)
    test_llm_summarization(text, "Gross Margin")
