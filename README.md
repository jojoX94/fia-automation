# FIA Automation Tool

This project automates FIA-related tasks using Selenium and provides a Streamlit interface for easy interaction.

## Project Structure

```
fia-automation/
├─ app/               ← Interface Streamlit
│   └─ main.py
├─ core/              ← Logique Selenium + traitement Excel
│   ├─ automate.py
│   └─ matching.py
├─ io/
│   ├─ excel_in/      ← Input Excel files
│   ├─ reports/       ← Generated reports
│   └─ logs/          ← Application logs
├─ .env               ← Environment variables
├─ requirements.txt   ← Project dependencies
└─ README.md
```

## Setup

1. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Unix/macOS
   # or
   .\venv\Scripts\activate  # On Windows
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure environment variables in `.env` file

## Usage

Run the Streamlit application:
```bash
streamlit run app/main.py
```