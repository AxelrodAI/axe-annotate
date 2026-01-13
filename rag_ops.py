"""
RAG Operations Module
=====================

This module handles the Retrieval Augmented Generation pipeline:
1. Search: Finding relevant document URLs (transcripts, filings)
2. Fetch: Scraping content using Firecrawl
3. Retrieve: Finding relevant sections based on the requested KPI/Line Item

Dependencies:
- firecrawl-py (pip install firecrawl-py)
"""

import os
import re
import time
import json

import os
import re
import time
import json
import edgar_ops  # Import our new module

# Placeholder for Firecrawl client
# try:
#     from firecrawl import FirecrawlApp
# except ImportError:
#     FirecrawlApp = None

class RAGPipeline:
    def __init__(self, api_key=None):
        self.api_key = api_key or os.getenv("FIRECRAWL_API_KEY")
        self.client = None
        
        # Initialize Firecrawl if key is available
        # if self.api_key and FirecrawlApp:
        #     self.client = FirecrawlApp(api_key=self.api_key)
        #     print("[RAG] Firecrawl initialized")
        # else:
        #     print("[RAG] Firecrawl not available - Using EDGAR Fallback")

    def get_filing_content(self, ticker, period):
        """
        Main entry point for fetching text.
        Prioritizes:
        1. Firecrawl (if key exists) -> Transcripts
        2. EDGAR (Free) -> 10-Q/10-K Filings
        3. Mock Data -> Fallback
        """
        # Strategy 1: Firecrawl (Disabled for now as user has no key)
        if self.client:
            url = self._find_transcript_url(ticker, period)
            return self._fetch_firecrawl(url)
            
        # Strategy 2: EDGAR (Free, Public)
        print(f"[RAG] Fetching SEC EDGAR filing for {ticker} ({period})...")
        # Map period "Q1 2024" to form type. 
        # Logic: If Q1-Q3 -> 10-Q. If Q4/FY -> 10-K.
        form_type = "10-K" if "Q4" in period or "FY" in period else "10-Q"
        
        text = edgar_ops.get_latest_filing_text(ticker, form_type)
        if text:
            return text
            
        # Strategy 3: Detailed Mock Data (Last Resort)
        print("[RAG] EDGAR failed or not found. Using Mock Data.")
        return self._get_mock_transcript(ticker, period)

    def find_transcript_url(self, ticker, period):
        """Deprecated: Logic moved to get_filing_content"""
        return f"https://www.seekingalpha.com/symbol/{ticker}/earnings/transcripts"

    def fetch_content(self, url_or_ticker):
        """
        Legacy wrapper for backward compatibility if needed, 
        but we prefer get_filing_content now.
        """
        pass

    def retrieve_context(self, text, query_kpi):
        """
        Simple retrieval: Find paragraphs containing keywords from the KPI.
        """
        if not text:
            return "No content available."

        print(f"[RAG] Retrieving insights for: '{query_kpi}'")
        
        # Split text into paragraphs (try to respect structure)
        # HTML cleaning in edgar_ops returns \n for tags, but we might have big blobs.
        # Let's split by double newline or period-space-space.
        paragraphs = re.split(r'\n\s*\n', text)
        if len(paragraphs) < 5: # If text is too dense
            paragraphs = text.split('. ')
            
        relevant_chunks = []
        
        # Normalize query
        keywords = query_kpi.lower().split()
        stopwords = {'revenue', 'income', 'profit', 'margin', 'sales', 'of', 'in', 'the', 'a', 'an', 'to', 'for', 'and', 'from'}
        search_terms = [k for k in keywords if k not in stopwords and len(k) > 2]
        
        if not search_terms:
            search_terms = keywords 
            
        print(f"[RAG] Search terms: {search_terms}")

        for p in paragraphs:
            p_lower = p.lower()
            # Score: +1 for each term found
            score = sum(1 for term in search_terms if term in p_lower)
            
            # Boost for strict whole-word matches or proximity?
            # For now, simple count is fine.
            
            if score > 0 and len(p) > 50: # Filter simplistic lines
                relevant_chunks.append((score, p.strip()))

        # Sort by relevance
        relevant_chunks.sort(key=lambda x: x[0], reverse=True)
        
        # Return top chunks
        top_chunks = []
        seen = set()
        for score, chunk in relevant_chunks[:4]:
            if chunk not in seen:
                # Truncate very long chunks
                if len(chunk) > 500:
                    chunk = chunk[:500] + "..."
                top_chunks.append(chunk)
                seen.add(chunk)
                if len(top_chunks) >= 3:
                    break
        
        if not top_chunks:
            return "No specific comments found for this item in the filing."
            
        return "\n\n".join([f"> {c}" for c in top_chunks])

    def _fetch_firecrawl(self, url):
        # ... existing firecrawl logic ...
        pass

    def _get_mock_transcript(self, ticker, period):
        """Returns a mock transcript text for testing."""
        return f"""
        (Mock Transcript for {ticker} {period})
        Speaker 1 (CEO): Good afternoon. We are pleased to report strong results.
        
        Our Total Revenue grew 15% year-over-year, driven by strong performance in our Cloud segment.
        
        Net Income was solid at $5 billion.
        
        The Cloud segment specifically saw a 30% increase in sales. We are seeing massive demand for our AI infrastructure.
        
        Operating margin improved by 200 basis points due to our operational efficiency initiatives.
        """

# Singleton instance for easy import
rag = RAGPipeline()
