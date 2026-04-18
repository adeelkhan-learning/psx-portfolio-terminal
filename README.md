PSX Portfolio Terminal: Automated AI Extractor & Interactive Dashboard
Overview
A comprehensive financial data pipeline and visualization terminal designed for the Pakistan Stock Exchange (PSX). This project automates the extraction of unstructured broker data (PDFs and emails) using Vision AI, processes the data into a mathematically verified ledger, and visualizes the portfolio through a live, interactive Streamlit dashboard.

P.S. The pdf input file used for this project came from AKD and CDC through email. 

System Architecture
1. Data Extraction & Transformation (extractor.py)
Vision AI Processing: Utilizes PyMuPDF and the Groq Llama 3 Vision API to accurately parse complex, misaligned tabular data from broker trade confirmations and dividend warrants.

Deterministic Cleaning: Applies Python regex to sanitize broker extensions (e.g., -READY) and strictly calculates cost-basis, taxes, and commissions to ensure mathematical accuracy.

Automated Ledger: Maintains a centralized Excel database (PSX_Portfolio_Tracker.xlsx), automatically categorizing Trades, Dividends, and Fund Transfers while preventing duplicate entries.

2. Live Portfolio Dashboard (app.py)
Top Ticker: Shows the current stock price and average price of stock I bought and calculates if its in positive or negative trend.

Advanced Analytics: Tracks realized and unrealized profit/loss, visualizes asset allocation, and provides a running cash-flow balance of deposits and withdrawals.

Dividend Intelligence: Analyzes historical payment intervals to predict upcoming dividend dates and tracks total yield per asset.

Transaction Mapping: Plots historical buy, sell, and dividend dates directly onto interactive stock price charts using Plotly.

Repository Structure
Plaintext
psx-portfolio-terminal/
├── extractor.py              # AI-powered PDF/text extraction script
├── app.py                    # Streamlit interactive dashboard
├── generate_sample_data.py   # Utility to create sample Excel data for testing
├── requirements.txt          # Python dependencies
├── .env.example              # Environment variables template
├── .gitignore                # Git ignore rules for security
└── README.md                 # Project documentation
Setup and Installation
1. Clone the Repository

Bash
git clone https://github.com/yourusername/psx-portfolio-terminal.git
cd psx-portfolio-terminal
2. Configure the Virtual Environment (Recommended)

Bash
python -m venv venv

# On Windows:
venv\Scripts\activate
# On Mac/Linux:
source venv/bin/activate
3. Install Dependencies

Bash
pip install -r requirements.txt
4. Configure Environment Variables
Rename the .env.example file to .env. Open the file and insert your Groq API key:

Plaintext
GROQ_API_KEY="your_api_key_here"
5. Directory Initialization
Ensure the following directories exist in your root folder to store your raw broker files prior to extraction:

Trade_Confirmations/

Dividends/

Funds_Transfers/

Usage Guide
Step 1: Run the Extraction Pipeline
Place your unformatted broker PDFs or .txt email confirmations into their respective folders. Execute the extractor script to parse the files and update your local database:

Bash
python extractor.py
Note: The script will automatically skip files it has already processed in previous runs.

Step 2: Launch the Dashboard
Once your database is populated, launch the Streamlit terminal to view your live portfolio:

Bash
streamlit run app.py
The application will automatically open in your default web browser at http://localhost:8501.

(Optional) If you wish to test the dashboard UI without processing your own broker files, run python generate_sample_data.py to create a mock PSX_Portfolio_Tracker.xlsx file.

Security Notice
This application processes highly sensitive personal financial data. Ensure that your .env file containing your API keys and your PSX_Portfolio_Tracker.xlsx database file are strictly excluded from version control via the included .gitignore file. Never commit these files to a public repository.

License
This project is licensed under the MIT License. See the LICENSE file for details.

## ⚖️ Legal Disclaimer

This project is open-source and intended solely for **educational and personal use**. 

* **Data Ownership:** All market data, stock prices, and related financial information accessed via this script are the exclusive property of the Pakistan Stock Exchange (PSX). 
* **No Commercial Use:** This tool is not intended for commercial use, data redistribution, or monetization. Do not use this codebase to bypass PSX Data Services Vending programs.
* **Rate Limiting & Server Respect:** The dashboard architecture deliberately incorporates server-side caching to respect PSX server resources and prevent high-frequency scraping. Users modifying this code are strongly advised to maintain these ethical scraping limits.
* **Liability:** The author of this repository is not responsible for any financial losses, trading errors, or legal repercussions resulting from the use of this software. By cloning or running this code, you accept full responsibility for your actions and your compliance with PSX Terms of Use.
