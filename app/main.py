import streamlit as st
import sys
from pathlib import Path

# Add parent directory to Python path
sys.path.append(str(Path(__file__).parent.parent))

from core import automate, matching

def main():
    st.title("FIA Automation Tools")
    # Add your Streamlit interface code here

if __name__ == "__main__":
    main()