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
        #     print("[RAG] Firecrawl not available (missing key or library)")

    def find_transcript_url(self, ticker, period):
        """
        Constructs a search query to find the transcript URL.
        In a real scenario, this would use a Search API (Google/Bing).
        For now, it returns a simulated URL or queries a known free source.
        """
        query = f"{ticker} {period} earnings transcript site:seekingalpha.com OR site:motleyfool.com"
        print(f"[RAG] Search Query: {query}")
        
        # Mock logic: Return a placeholder URL if we can't search
        # In a real agentic loop, I would use the `search_web` tool here.
        # But for the python script, we might need a distinct search API.
        return f"https://www.seekingalpha.com/symbol/{ticker}/earnings/transcripts"

    def fetch_content(self, url):
        """
        Uses Firecrawl to scrape the content of the URL.
        """
        if not self.client:
            print("[RAG] Mocking fetch (no Firecrawl client)...")
            return self._get_mock_transcript()

        try:
            print(f"[RAG] Firecrawling: {url}")
            # Real call would be:
            # scrape_result = self.client.scrape_url(url, params={'formats': ['markdown']})
            # return scrape_result['markdown']
            return self._get_mock_transcript() # Fallback for now
        except Exception as e:
            print(f"[RAG] Fetch error: {e}")
            return None

    def retrieve_context(self, text, query_kpi):
        """
        Simple retrieval: Find paragraphs containing keywords from the KPI.
        """
        if not text:
            return "No content available."

        print(f"[RAG] Retrieving insights for: '{query_kpi}'")
        
        # Split text into paragraphs (double newline usually)
        paragraphs = text.split("\n\n")
        relevant_chunks = []
        
        # Normalize query
        keywords = query_kpi.lower().split()
        # Remove common words
        stopwords = {'revenue', 'income', 'profit', 'margin', 'sales', 'of', 'in', 'the', 'a', 'an', 'to'}
        # Keep specialized words
        search_terms = [k for k in keywords if len(k) > 3]
        
        if not search_terms:
            search_terms = keywords # Fallback if everything was filtered
            
        print(f"[RAG] Search terms: {search_terms}")

        for p in paragraphs:
            p_lower = p.lower()
            # Simple scoring: count keyword matches
            score = sum(1 for term in search_terms if term in p_lower)
            
            if score > 0:
                relevant_chunks.append((score, p.strip()))

        # Sort by relevance (score)
        relevant_chunks.sort(key=lambda x: x[0], reverse=True)
        
        # Return top 3 unique chunks
        top_chunks = []
        seen = set()
        for score, chunk in relevant_chunks[:5]:
            if chunk not in seen:
                top_chunks.append(chunk)
                seen.add(chunk)
                if len(top_chunks) >= 3:
                    break
        
        if not top_chunks:
            return "No specific comments found for this item."
            
        return "\n\n".join([f"> {c}" for c in top_chunks])

    def _get_mock_transcript(self):
        """Returns a mock transcript text for testing."""
        return """
        Speaker 1 (CEO): Good afternoon. We are pleased to report strong results for Q4 2025.
        
        Our total revenue grew 15% year-over-year to $25 billion, driven by strong performance in our Cloud segment.
        
        The Cloud segment specifically saw a 30% increase in sales, reaching $10 billion. We are seeing massive demand for our AI infrastructure.
        
        Regarding iPhone sales, we faced some supply chain headwinds, but demand remains robust. iPhone revenue was flat at $40 billion.
        
        Operating margin improved by 200 basis points due to our operational efficiency initiatives.
        
        Speaker 2 (CFO): I will provide more details on the financials. Net income was $5 billion, an increase of 10%.
        
        For the FICC trading revenue, we saw a slight decline of 5% due to lower volatility in the markets. However, equity trading remained strong.
        
        We are updating our guidance for the next fiscal year. We expect revenue to grow between 10% and 12%.
        """

# Singleton instance for easy import
rag = RAGPipeline()
