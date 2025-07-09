#!/usr/bin/env python3
"""
Run script for the Risk Control Matrix Analyzer.

This script sets up the environment and starts the Streamlit application.
"""

import os
import sys
import subprocess
import webbrowser
from dotenv import load_dotenv

def main():
    """Set up environment and start the Streamlit application."""
    # Load environment variables from .env file
    load_dotenv()
    
    # Check if Gemini API key is set
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("Error: GEMINI_API_KEY environment variable not found.")
        print("Please set it in a .env file or in your environment variables.")
        print("You can obtain a Gemini API key from: https://ai.google.dev/")
        return 1
    
    # Check if dependencies are installed
    try:
        import streamlit
        import pandas
        import google.generativeai
        import chromadb
    except ImportError as e:
        print(f"Error: Missing dependencies - {e}")
        print("Please install the required dependencies with: pip install -r requirements.txt")
        return 1
    
    # Start Streamlit application
    print("Starting Risk Control Matrix Analyzer...")
    
    # Use subprocess to start Streamlit in a separate process
    streamlit_process = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", "app.py"],
        env=os.environ.copy()
    )
    
    # Open browser automatically
    webbrowser.open("http://localhost:8501")
    
    try:
        # Wait for the process to terminate
        streamlit_process.wait()
    except KeyboardInterrupt:
        print("\nStopping application...")
        streamlit_process.terminate()
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 