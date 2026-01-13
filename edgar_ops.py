"""
EDGAR Operations Module
=======================

Provides free, no-registration access to SEC EDGAR filings (10-Q, 10-K).
Uses direct HTTP requests to SEC.gov endpoints.

Rules:
- Must use a proper User-Agent header (SEC requirement).
- Max 10 requests/sec (though we won't hit this).

Endpoints:
- Company Tickers: https://www.sec.gov/files/company_tickers.json
- Submissions: https://data.sec.gov/submissions/CIK{cik}.json
"""

import json
import time
import re
import urllib.request
import urllib.error

# User-Agent is MANDATORY for SEC.gov
# Format: "Sample Company Name AdminContact@sample.com"
HEADERS = {
    "User-Agent": "AxeAnnotate/2.2 (OpenSourceResearch; contact@axelrod.ai)",
    "Accept-Encoding": "gzip, deflate",
    "Host": "www.sec.gov" 
    # Note: data.sec.gov is used for submissions, www.sec.gov for tickers
}

def _make_request(url, host="www.sec.gov"):
    """Helper to make legitimate requests to SEC.gov"""
    req = urllib.request.Request(url, headers={
        "User-Agent": HEADERS["User-Agent"],
        "Assert-Encoding": "gzip, deflate",
        "Host": host
    })
    
    try:
        with urllib.request.urlopen(req) as response:
            data = response.read()
            # Handle gzip if needed? usually urllib handles it or returns bytes
            return data
    except urllib.error.HTTPError as e:
        print(f"[EDGAR] HTTP Error {e.code}: {url}")
        return None
    except Exception as e:
        print(f"[EDGAR] Request Error: {e}")
        return None

# Cache for Ticker -> CIK mapping
_TICKER_CACHE = {}

def get_cik_from_ticker(ticker):
    """
    Resolves a ticker symbol (e.g., AAPL) to its CIK number (0000320193).
    """
    ticker = ticker.upper().strip()
    
    # Check cache first
    if ticker in _TICKER_CACHE:
        return _TICKER_CACHE[ticker]
    
    print("[EDGAR] Fetching ticker map...")
    url = "https://www.sec.gov/files/company_tickers.json"
    data = _make_request(url, host="www.sec.gov")
    
    if not data:
        return None
        
    try:
        companies = json.loads(data)
        # Structure is {"0": {"cik_str": 320193, "ticker": "AAPL", "title": "..."}, ...}
        
        for key, val in companies.items():
            t = val["ticker"]
            cik = val["cik_str"]
            # Pad CIK to 10 digits
            cik_str = str(cik).zfill(10)
            _TICKER_CACHE[t] = cik_str
            
            if t == ticker:
                return cik_str
                
    except Exception as e:
        print(f"[EDGAR] Error parsing ticker map: {e}")
        
    return _TICKER_CACHE.get(ticker)

def get_latest_filing_text(ticker, form_type="10-Q"):
    """
    Fetches the full text of the latest filing of a specific type.
    """
    cik = get_cik_from_ticker(ticker)
    if not cik:
        print(f"[EDGAR] Could not find CIK for {ticker}")
        return None
        
    print(f"[EDGAR] Found CIK for {ticker}: {cik}")
    
    # Fetch submissions history
    # URL Format: https://data.sec.gov/submissions/CIK##########.json
    url = f"https://data.sec.gov/submissions/CIK{cik}.json"
    data = _make_request(url, host="data.sec.gov")
    
    if not data:
        return None
        
    try:
        history = json.loads(data)
        filings = history.get("filings", {}).get("recent", {})
        
        # Lists of metadata
        forms = filings.get("form", [])
        accession_nums = filings.get("accessionNumber", [])
        primary_docs = filings.get("primaryDocument", [])
        
        target_idx = -1
        
        # Find first matching form
        for i, form in enumerate(forms):
            if form == form_type:
                target_idx = i
                break
                
        if target_idx == -1:
            print(f"[EDGAR] No {form_type} found for {ticker}")
            return None
            
        acc_num = accession_nums[target_idx]
        primary_doc = primary_docs[target_idx]
        
        # Construct Document URL
        # Accession number needs dashes removed for the folder path
        # Format: https://www.sec.gov/Archives/edgar/data/{cik}/{acc_num_no_dash}/{primary_doc}
        acc_clean = acc_num.replace("-", "")
        # Note: CIK in path must be integer (no leading zeros) usually, but try string first
        cik_int = int(cik)
        
        doc_url = f"https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc_clean}/{primary_doc}"
        print(f"[EDGAR] Fetching document: {doc_url}")
        
        # Add delay to be nice to SEC
        time.sleep(0.15)
        
        doc_content = _make_request(doc_url, host="www.sec.gov")
        
        if doc_content:
            # Basic cleanup: scrape text from HTML
            text = _clean_html(doc_content.decode('utf-8', errors='ignore'))
            return text
            
    except Exception as e:
        print(f"[EDGAR] Error processing filing: {e}")
        
    return None

def _clean_html(html_content):
    """
    Simple HTML stripper. 
    In production, use BeautifulSoup. This is a Zero-Dependency regex fallback.
    """
    # Remove script and style tags
    cleaned = re.sub(r'<(script|style).*?</\1>', '', html_content, flags=re.DOTALL)
    # Remove comments
    cleaned = re.sub(r'<!--.*?-->', '', cleaned, flags=re.DOTALL)
    # Remove tags
    cleaned = re.sub(r'<[^>]+>', ' ', cleaned)
    
    # Simple Entity Decoding (Zero-Dependency)
    import html
    cleaned = html.unescape(cleaned)

    # Fix whitespace
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    return cleaned

if __name__ == "__main__":
    # Test
    print("Testing EDGAR Fetcher...")
    text = get_latest_filing_text("AAPL", "10-Q")
    if text:
        print(f"Success! Retrieved {len(text)} characters.")
        print("Sample (first 500 chars):")
        print(text[:500])
    else:
        print("Failed.")
