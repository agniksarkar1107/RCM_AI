"""
Utility modules for the Risk Control Matrix Analyzer.

This package contains modules for:
- Document processing: Extract structured data from different document formats
- Database integration: Store and query data in ChromaDB
- Gemini integration: Analyze data using Google's Gemini AI
"""

from utils.document_processor import process_document
from utils.db import initialize_chroma, store_in_chroma, query_chroma
from utils.gemini import initialize_gemini, analyze_risk_with_gemini 