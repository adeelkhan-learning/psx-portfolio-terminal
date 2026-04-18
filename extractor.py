#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Mar 31 20:39:22 2026

@author: ADeel
"""

import base64
import fitz  # This is PyMuPDF
import os
import glob
import json
import re
import time
import pdfplumber
import pandas as pd
from datetime import datetime
from groq import Groq
from dotenv import load_dotenv

# --- SETUP & AUTHENTICATION ---
script_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(script_dir, '.env')
load_dotenv(dotenv_path=env_path)

api_key = os.environ.get("GROQ_API_KEY")
if not api_key:
    raise ValueError("CRITICAL ERROR: GROQ_API_KEY is not loaded.")

client = Groq(api_key=api_key)
EXCEL_FILE = "PSX_Portfolio_Tracker.xlsx"

# --- 1. EXCEL DATABASE CONFIGURATION ---
def setup_excel_database():
    """Creates the Excel file and necessary sheets if they don't exist."""
    if not os.path.exists(EXCEL_FILE):
        print(f"Creating new database: {EXCEL_FILE}")
        trades_df = pd.DataFrame(columns=[
            "Trade Date", "Settlement Date", "Transaction Type", "Ticker", 
            "Quantity", "Price", "Commission", "Taxes and Fees", "Net Total Value"
        ])
        div_df = pd.DataFrame(columns=[
            "Payment Date", "Company Name", "Ticker", "No. of Securities", 
            "Rate Per Security", "Gross Dividend", "Zakat Deducted", "Tax Deducted", "Net Amount Paid"
        ])
        funds_df = pd.DataFrame(columns=[
            "Date", "Transfer Type", "Amount Deposit", "Amount Hold Against Charges / Dues", "Amount Transferred To Exposure"
        ])
        log_df = pd.DataFrame(columns=["Processed Filename", "Date Processed"])
        
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            trades_df.to_excel(writer, sheet_name='Trades', index=False)
            div_df.to_excel(writer, sheet_name='Dividends', index=False)
            funds_df.to_excel(writer, sheet_name='Funds', index=False)
            log_df.to_excel(writer, sheet_name='Processed_Files', index=False)
    else:
        xls = pd.ExcelFile(EXCEL_FILE)
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            if 'Dividends' not in xls.sheet_names:
                div_df = pd.DataFrame(columns=[
                    "Payment Date", "Company Name", "Ticker", "No. of Securities", 
                    "Rate Per Security", "Gross Dividend", "Zakat Deducted", "Tax Deducted", "Net Amount Paid"
                ])
                div_df.to_excel(writer, sheet_name='Dividends', index=False)
            if 'Funds' not in xls.sheet_names:
                funds_df = pd.DataFrame(columns=[
                    "Date", "Transfer Type", "Amount Deposit", "Amount Hold Against Charges / Dues", "Amount Transferred To Exposure"
                ])
                funds_df.to_excel(writer, sheet_name='Funds', index=False)

def is_file_processed(filename):
    try:
        log_df = pd.read_excel(EXCEL_FILE, sheet_name='Processed_Files')
        
        if log_df.empty or 'Processed Filename' not in log_df.columns:
            return False
        
        # 1. Drop any empty rows
        # 2. Convert everything to a strict string
        # 3. Strip all invisible spaces from the Excel data
        processed_list = [str(name).strip() for name in log_df['Processed Filename'].dropna().tolist()]
        
        # Strip invisible spaces from our current file and check if it's in the list
        return str(filename).strip() in processed_list
        
    except Exception as e:
        print(f"Warning: Could not read log sheet properly ({e})")
        return False

def log_processed_file(filename):
    log_df = pd.read_excel(EXCEL_FILE, sheet_name='Processed_Files')
    new_log = pd.DataFrame({"Processed Filename": [filename], "Date Processed": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]})
    log_df = pd.concat([log_df, new_log], ignore_index=True)
    
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        log_df.to_excel(writer, sheet_name='Processed_Files', index=False)

def save_trades_to_excel(trades_data):
    if isinstance(trades_data, dict):
        trades_data = [trades_data]
        
    trades_df = pd.read_excel(EXCEL_FILE, sheet_name='Trades')
    new_trades_df = pd.DataFrame(trades_data)
    
    if trades_df.empty or trades_df.isna().all().all():
        trades_df = new_trades_df
    else:
        trades_df = trades_df.dropna(how='all', axis=1)
        trades_df = pd.concat([trades_df, new_trades_df], ignore_index=True)
    
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        trades_df.to_excel(writer, sheet_name='Trades', index=False)

def save_dividends_to_excel(div_data):
    if isinstance(div_data, dict):
        div_data = [div_data]
        
    div_df = pd.read_excel(EXCEL_FILE, sheet_name='Dividends')
    new_div_df = pd.DataFrame(div_data)
    
    if div_df.empty or div_df.isna().all().all():
        div_df = new_div_df
    else:
        div_df = div_df.dropna(how='all', axis=1)
        div_df = pd.concat([div_df, new_div_df], ignore_index=True)
    
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        div_df.to_excel(writer, sheet_name='Dividends', index=False)

# --- 2. EXTRACTION LOGIC ---

def extract_images_from_pdf(pdf_path):
    """Converts ALL pages of the PDF into high-res images for the Vision AI."""
    base64_images = []
    try:
        doc = fitz.open(pdf_path)
        zoom = fitz.Matrix(2, 2) 
        
        for page in doc:
            pix = page.get_pixmap(matrix=zoom)
            img_bytes = pix.tobytes("jpeg")
            base64_image = base64.b64encode(img_bytes).decode('utf-8')
            base64_images.append(base64_image)
            
        return base64_images
    except Exception as e:
        print(f"Error converting PDF to images: {e}")
        return []

def parse_trade_data_with_vision(pdf_path):
    """Sends the actual images of the document to Groq's Llama 4 Scout Vision model."""
    base64_images = extract_images_from_pdf(pdf_path)
    if not base64_images:
        return None

    prompt = """
    You are a precise financial data extraction tool. 
    Carefully look at these images of an AKD Securities Trade Confirmation. It may span multiple pages.
    
    CRITICAL RULES:
    1. "Ticker": Perform STRICT literal character-by-character transcription of the symbol column. DO NOT auto-correct spelling, and DO NOT inject extra vowels. If the image says 'CNERGY', do NOT output 'CENERGY'. If it says 'MZNPETF', do NOT output 'MZNPEETF'. Copy the letters EXACTLY as printed. Include the broker extensions (e.g., "-READY") if present.
    2. "Printed Taxes": Look visually at the table grid. Extract the exact number written in the Taxes/Levies column for EACH specific row. Do not calculate it. Do not let numbers bleed together.
    3. "Printed Net Value": Extract the final number on that specific row. DO NOT grab the Grand Total at the bottom of the page.
    
    You MUST return EXACTLY this JSON format and nothing else. Do not include markdown tags.
    {
      "trades": [
        {
          "Trade Date": "YYYY-MM-DD",
          "Settlement Date": "YYYY-MM-DD",
          "Transaction Type": "BUY",  // MUST be "BUY", "SELL", or "IPO". If the document says "PURCHASE", output "BUY".
          "Ticker": "ILP",
          "Quantity": 500,
          "Price": 75.40,
          "Commission": 56.55,
          "Printed Taxes": 0.75,
          "Printed Net Value": 37757.3
        }
      ]
    }
    """
    
    content_array = [{"type": "text", "text": prompt}]
    for b64_img in base64_images:
        content_array.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/jpeg;base64,{b64_img}"
            }
        })
    
    try:
        chat_completion = client.chat.completions.create(
            messages=[
                {
                    "role": "user",
                    "content": content_array
                }
            ],
            model="meta-llama/llama-4-scout-17b-16e-instruct", 
            temperature=0, 
        )
        
        response_content = chat_completion.choices[0].message.content.strip()
        
        match = re.search(r'\{.*\}', response_content, re.DOTALL)
        if match:
            response_content = match.group(0)
            
        parsed_data = json.loads(response_content)
        trades_list = parsed_data.get("trades", [])
        
        # --- PYTHON SANITY CHECK & MATH ---
        for trade in trades_list:
            
            # 1. --- THE TRUE REGEX FIX (No Dictionary Needed) ---
            raw_ticker = str(trade.get("Ticker", "")).upper().strip()
            
            # This safely removes extensions ONLY if preceded by a dash or space. 
            # "OGDC-OCT" becomes "OGDC". "OCTOPUS" remains "OCTOPUS".
            clean_ticker = re.sub(r'[\-\s]+(READY|FUTURE|SPOT|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b.*', '', raw_ticker).strip()
            
            # 2. Fix the AI's visual reading errors
            if clean_ticker == "MZNPEETF":
                clean_ticker = "MZNPETF"
            elif clean_ticker == "CENERGY":
                clean_ticker = "CNERGY"
            
            trade["Ticker"] = clean_ticker
            # -------------------------------------------------
            
            base_value = trade.get("Price", 0.0) * trade.get("Quantity", 0)
            commission = trade.get("Commission", 0.0)
            printed_net = trade.get("Printed Net Value", 0.0)
            printed_tax = trade.get("Printed Taxes", 0.0)
            
            max_logical_tax = base_value * 0.05 
            
            tx_type = str(trade.get("Transaction Type", "")).upper().strip()
            
            if tx_type == "PURCHASE":
                tx_type = "BUY"
                trade["Transaction Type"] = "BUY"
            
            if tx_type == "IPO":
                trade["Taxes and Fees"] = 0.0
                trade["Net Total Value"] = round((trade.get("Price", 0.0) * trade.get("Quantity", 0)), 2)
                
            elif tx_type == "BUY":
                calc_tax = printed_net - base_value - commission
                if printed_net > 0 and (0 <= calc_tax <= max_logical_tax):
                    trade["Taxes and Fees"] = round(calc_tax, 2)
                    trade["Net Total Value"] = round(printed_net, 2)
                else:
                    trade["Taxes and Fees"] = round(printed_tax, 2)
                    trade["Net Total Value"] = round(base_value + commission + printed_tax, 2)
                    
            elif tx_type == "SELL":
                calc_tax = base_value - printed_net - commission
                if printed_net > 0 and (0 <= calc_tax <= max_logical_tax):
                    trade["Taxes and Fees"] = round(calc_tax, 2)
                    trade["Net Total Value"] = round(printed_net, 2)
                else:
                    trade["Taxes and Fees"] = round(printed_tax, 2)
                    trade["Net Total Value"] = round(base_value - commission - printed_tax, 2)
            
            trade.pop("Printed Net Value", None) 
            trade.pop("Printed Taxes", None)
            
        return trades_list
    except Exception as e:
        print(f"An error occurred during Vision extraction: {e}")
        return None

def extract_text_from_pdf(pdf_path):
    extracted_data = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    extracted_data += "--- STRUCTURED TABLE DATA ---\n"
                    for table in tables:
                        for row in table:
                            clean_row = [str(cell).replace('\n', ' ').strip() if cell is not None else "" for cell in row]
                            if any(clean_row):
                                extracted_data += " | ".join(clean_row) + "\n"
                
                extracted_data += "\n--- RAW TEXT FOR DATES ---\n"
                extracted_data += page.extract_text() + "\n"
                
        return extracted_data
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return None

def parse_trade_data_with_groq(raw_text):
    prompt = f"""
    You are a precise financial data extraction tool specializing in the Pakistan Stock Exchange (PSX). 
    Analyze the following text. The tabular data has been pre-processed so that columns are separated by the "|" symbol to prevent number bleeding.
    
    CRITICAL RULES:
    1. "Ticker": Perform STRICT literal character-by-character transcription of the symbol column. DO NOT auto-correct spelling, and DO NOT inject extra vowels. Copy the letters EXACTLY as printed. Include the broker extensions (e.g., "-READY") if present.
    2. "Printed Taxes": Look at the "STRUCTURED TABLE DATA" section. Find the specific row for the trade. The numbers are separated by "|". Look for the Tax/Levy column for that specific row and extract the exact float. DO NOT calculate it.
    3. "Printed Net Value": Extract the absolute LAST number on the row for this specific trade. DO NOT extract the "Grand Total" from the bottom of the page.
    4. NUMBERS & FORMAT: Output all numbers as floats WITHOUT commas (e.g., 1250.0).
    
    You MUST return a JSON object containing a single key "trades", which is a list of dictionaries.
    Each dictionary represents one transaction and MUST include exactly these keys:
    - "Trade Date" (Format: YYYY-MM-DD. Look in the RAW TEXT section).
    - "Settlement Date" (Format: YYYY-MM-DD. Use Trade Date if not explicitly mentioned).
    - "Transaction Type" (Must be exactly "BUY", "SELL", or "IPO")
    - "Ticker" (Follow Rule 1)
    - "Quantity" (Integer)
    - "Price" (Float)
    - "Commission" (Float. Default 0.0)
    - "Printed Taxes" (Float. Default 0.0)
    - "Printed Net Value" (Float. The final total for the row. Default 0.0)
    - "IPO Amount Paid" (Float. Default 0.0)
    - "IPO Amount Refunded" (Float. Default 0.0)
    
    Raw Text:
    {raw_text}
    """
    try:
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.1-8b-instant", 
            temperature=0, 
            response_format={"type": "json_object"} 
        )
        response_content = chat_completion.choices[0].message.content.strip()
        parsed_data = json.loads(response_content)
        trades_list = parsed_data.get("trades", [])
        
        # --- PYTHON SANITY CHECK & MATH ---
        for trade in trades_list:
            
            # --- THE TRUE REGEX FIX (No Dictionary Needed) ---
            raw_ticker = str(trade.get("Ticker", "")).upper().strip()
            
            # This safely removes extensions ONLY if preceded by a dash or space. 
            # "OGDC-OCT" becomes "OGDC". "OCTOPUS" remains "OCTOPUS".
            clean_ticker = re.sub(r'[\-\s]+(READY|FUTURE|SPOT|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b.*', '', raw_ticker).strip()
            
            # 2. Fix the AI's visual reading errors
            if clean_ticker == "MZNPEETF":
                clean_ticker = "MZNPETF"
            elif clean_ticker == "CENERGY":
                clean_ticker = "CNERGY"
            
            trade["Ticker"] = clean_ticker
            # -------------------------------------------------
            
            base_value = trade.get("Price", 0.0) * trade.get("Quantity", 0)
            commission = trade.get("Commission", 0.0)
            printed_net = trade.get("Printed Net Value", 0.0)
            printed_tax = trade.get("Printed Taxes", 0.0)
            
            max_logical_tax = base_value * 0.05 
            
            if trade.get("Transaction Type") == "IPO":
                qty = trade.get("Quantity", 0)
                paid = trade.get("IPO Amount Paid", 0.0)
                refunded = trade.get("IPO Amount Refunded", 0.0)
                if qty > 0:
                    trade["Price"] = round((paid - refunded) / qty, 4)
                trade["Commission"] = 0.0
                trade["Taxes and Fees"] = 0.0
                trade["Net Total Value"] = round((trade.get("Price", 0.0) * qty), 2)
                
            elif trade.get("Transaction Type") == "BUY":
                calc_tax = printed_net - base_value - commission
                if printed_net > 0 and (0 <= calc_tax <= max_logical_tax):
                    trade["Taxes and Fees"] = round(calc_tax, 2)
                    trade["Net Total Value"] = round(printed_net, 2)
                else:
                    trade["Taxes and Fees"] = round(printed_tax, 2)
                    trade["Net Total Value"] = round(base_value + commission + printed_tax, 2)
                    
            elif trade.get("Transaction Type") == "SELL":
                calc_tax = base_value - printed_net - commission
                if printed_net > 0 and (0 <= calc_tax <= max_logical_tax):
                    trade["Taxes and Fees"] = round(calc_tax, 2)
                    trade["Net Total Value"] = round(printed_net, 2)
                else:
                    trade["Taxes and Fees"] = round(printed_tax, 2)
                    trade["Net Total Value"] = round(base_value - commission - printed_tax, 2)
            
            trade.pop("IPO Amount Paid", None)
            trade.pop("IPO Amount Refunded", None)
            trade.pop("Printed Net Value", None) 
            trade.pop("Printed Taxes", None)
            
        return trades_list
    except Exception as e:
        print(f"An error occurred during trade extraction: {e}")
        return None
    
def parse_dividend_data_with_groq(raw_text):
    prompt = f"""
    You are a precise financial data extraction tool specializing in the Pakistan Stock Exchange (PSX). 
    Analyze the following text from a Dividend Warrant PDF.
    
    CRITICAL RULES:
    1. "Company Name": Extract the exact full company name printed on the document (e.g., "Engro Corporation Limited").
    2. "Payment Date": The date in the document is written in DD-MM-YYYY format (Day first, then Month). You MUST convert this and output it EXACTLY as YYYY-MM-DD. (e.g., If the document says "05-12-2023", it means December 5th, so you must output "2023-12-05").
    3. NUMBERS & FORMAT: Output all numbers as standard floats/integers WITHOUT commas (e.g., 1250.0). Look strictly at the tabular data for the financial values.
    
    You MUST return a JSON object containing a single key "dividends", which is a list of dictionaries.
    Each dictionary represents one dividend payout and MUST include exactly these keys:
    - "Payment Date" (Follow Rule 2).
    - "Company Name" (Follow Rule 1).
    - "Ticker" (Always output an empty string "").
    - "No. of Securities" (Integer).
    - "Rate Per Security" (Float).
    - "Gross Dividend" (Float. Look for 'Amount of Dividend').
    - "Zakat Deducted" (Float. Default 0.0 if empty).
    - "Tax Deducted" (Float. Default 0.0 if empty).
    - "Net Amount Paid" (Float. Look for 'Amount Paid').
    
    Raw Text:
    {raw_text}
    """
    try:
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.1-8b-instant", 
            temperature=0, 
            response_format={"type": "json_object"} 
        )
        response_content = chat_completion.choices[0].message.content.strip()
        parsed_data = json.loads(response_content)
        div_list = parsed_data.get("dividends", [])
        
        for div in div_list:
            qty = div.get("No. of Securities", 0)
            rate = div.get("Rate Per Security", 0.0)
            zakat = div.get("Zakat Deducted", 0.0)
            tax = div.get("Tax Deducted", 0.0)
            
            calculated_gross = round(qty * rate, 2)
            if rate == 0.0 and div.get("Gross Dividend", 0.0) > 0 and qty > 0:
                div["Rate Per Security"] = round(div["Gross Dividend"] / qty, 4)
                calculated_gross = div.get("Gross Dividend", 0.0)
            
            div["Gross Dividend"] = calculated_gross
            div["Net Amount Paid"] = round(calculated_gross - zakat - tax, 2)
            
        return div_list
    except Exception as e:
        print(f"An error occurred during dividend extraction: {e}")
        return None

def save_funds_to_excel(funds_data):
    if isinstance(funds_data, dict):
        funds_data = [funds_data]
        
    funds_df = pd.read_excel(EXCEL_FILE, sheet_name='Funds')
    new_funds_df = pd.DataFrame(funds_data)
    
    if funds_df.empty or funds_df.isna().all().all():
        funds_df = new_funds_df
    else:
        funds_df = funds_df.dropna(how='all', axis=1)
        funds_df = pd.concat([funds_df, new_funds_df], ignore_index=True)
    
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        funds_df.to_excel(writer, sheet_name='Funds', index=False)

def parse_funds_data_with_groq(raw_text):
    prompt = f"""
    You are a precise data extraction tool. Analyze this email, which will either be a standard fund transfer OR an IPO subscription confirmation.
    
    CRITICAL RULES based on email type:
    - Type A (Standard Deposit): Set "Transfer Type" to "Direct Deposit". Extract "Amount Deposit", "Amount Hold Against Charges / Dues", and "Amount Transferred To Exposure" exactly as they appear.
    - Type B (IPO Subscription): Set "Transfer Type" to "IPO". Look for "Amount Payable". Set both "Amount Deposit" and "Amount Transferred To Exposure" equal to this "Amount Payable" value. Set "Amount Hold Against Charges / Dues" to 0.0.
    
    Also, extract the "Date" of the transaction (Format: YYYY-MM-DD). If the exact date isn't visible in the table, try to find a date in the email text headers.
    
    You MUST return EXACTLY this JSON format and nothing else:
    {{
      "funds": [
        {{
          "Date": "2023-10-25",
          "Transfer Type": "Direct Deposit",
          "Amount Deposit": 50000.0,
          "Amount Hold Against Charges / Dues": 0.0,
          "Amount Transferred To Exposure": 50000.0
        }}
      ]
    }}
    
    Raw Text:
    {raw_text}
    """
    try:
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.1-8b-instant", 
            temperature=0, 
            response_format={"type": "json_object"} 
        )
        response_content = chat_completion.choices[0].message.content.strip()
        parsed_data = json.loads(response_content)
        funds_list = parsed_data.get("funds", [])
        
        # --- NEW: PYTHON SANITIZATION LOOP ---
        # This forces the AI's output into a clean number format before saving
        for fund in funds_list:
            for key in ["Amount Deposit", "Amount Hold Against Charges / Dues", "Amount Transferred To Exposure"]:
                if key in fund:
                    try:
                        # Convert to string, strip commas, then convert back to strict float
                        clean_val = str(fund[key]).replace(',', '').strip()
                        fund[key] = float(clean_val)
                    except ValueError:
                        fund[key] = 0.0 # Failsafe
                        
        return funds_list
        
    except Exception as e:
        print(f"An error occurred during funds extraction: {e}")
        return None

def sort_all_sheets_by_date():
    """Reads the entire workbook, sorts the main sheets by their respective dates, and saves it back."""
    print("Sorting all ledger sheets by date...")
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        sheets_dict = pd.read_excel(xls, sheet_name=None)
        
        # Sort Trades
        if 'Trades' in sheets_dict and not sheets_dict['Trades'].empty:
            sheets_dict['Trades']['Trade Date'] = pd.to_datetime(sheets_dict['Trades']['Trade Date'], errors='coerce')
            sheets_dict['Trades'] = sheets_dict['Trades'].sort_values(by="Trade Date", ascending=True)
            sheets_dict['Trades']['Trade Date'] = sheets_dict['Trades']['Trade Date'].dt.strftime('%Y-%m-%d').fillna('')
            
        # Sort Dividends
        if 'Dividends' in sheets_dict and not sheets_dict['Dividends'].empty:
            sheets_dict['Dividends']['Payment Date'] = pd.to_datetime(sheets_dict['Dividends']['Payment Date'], errors='coerce')
            sheets_dict['Dividends'] = sheets_dict['Dividends'].sort_values(by="Payment Date", ascending=True)
            sheets_dict['Dividends']['Payment Date'] = sheets_dict['Dividends']['Payment Date'].dt.strftime('%Y-%m-%d').fillna('')
            
        # Sort Funds
        if 'Funds' in sheets_dict and not sheets_dict['Funds'].empty:
            sheets_dict['Funds']['Date'] = pd.to_datetime(sheets_dict['Funds']['Date'], errors='coerce')
            sheets_dict['Funds'] = sheets_dict['Funds'].sort_values(by="Date", ascending=True)
            sheets_dict['Funds']['Date'] = sheets_dict['Funds']['Date'].dt.strftime('%Y-%m-%d').fillna('')
        
        # Write everything back cleanly
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            for sheet_name, df in sheets_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        print("  ✅ Sorting complete.")
        
    except Exception as e:
        print(f"  ❌ Error sorting sheets: {e}")

# --- 3. MAIN AUTOMATION LOOPS ---
def run_automation_pipeline(folder_path):
    pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
    txt_files = glob.glob(os.path.join(folder_path, "*.txt"))
    all_files = pdf_files + txt_files
    
    if not all_files:
        print(f"No PDFs or TXT files found in '{folder_path}'.")
        return

    print(f"Found {len(all_files)} files in Trades folder. Checking for new files...\n")
    new_trades_count = 0

    for file_path in all_files:
        filename = os.path.basename(file_path)
        
        if is_file_processed(filename):
            print(f"⏭️  Skipping {filename} (Already processed)")
            continue
            
        structured_data = None 
        print(f"📄 Processing trade file: {filename}...")
        
        if filename.lower().endswith('.pdf'):
            structured_data = parse_trade_data_with_vision(file_path)
            
        elif filename.lower().endswith('.txt'):
            with open(file_path, 'r', encoding='utf-8') as f:
                text = f.read()
            structured_data = parse_trade_data_with_groq(text) 
        
        if structured_data:
            save_trades_to_excel(structured_data)
            log_processed_file(filename)
            new_trades_count += len(structured_data)
            print(f"  ✅ Success! Added {len(structured_data)} transactions to Excel.\n")
        else:
            print(f"  ❌ Failed to parse data from {filename}.\n")
            
        time.sleep(3)  

    print(f"--- Trade Pipeline Complete ---")
    print(f"Added {new_trades_count} new transaction records to {EXCEL_FILE}.\n")

def run_dividend_pipeline(folder_path):
    pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
    
    if not pdf_files:
        print(f"No PDFs found in '{folder_path}'.")
        return

    print(f"Found {len(pdf_files)} files in Dividends folder. Checking for new files...\n")
    new_div_count = 0

    for file_path in pdf_files:
        filename = os.path.basename(file_path)
        
        if is_file_processed(filename):
            print(f"⏭️  Skipping {filename} (Already processed)")
            continue
            
        print(f"📄 Processing dividend file: {filename}...")
        text = extract_text_from_pdf(file_path)
                
        if not text:
            print(f"  ❌ Could not read text from {filename}.\n")
            continue
            
        structured_data = parse_dividend_data_with_groq(text)
        
        if structured_data:
            save_dividends_to_excel(structured_data)
            log_processed_file(filename)
            new_div_count += len(structured_data)
            print(f"  ✅ Success! Added {len(structured_data)} dividend records to Excel.\n")
        else:
            print(f"  ❌ Failed to parse data from {filename}.\n")
            
        time.sleep(3) 

    print(f"--- Dividend Pipeline Complete ---")
    print(f"Added {new_div_count} new dividend records to {EXCEL_FILE}.")

def run_funds_pipeline(folder_path):
    txt_files = glob.glob(os.path.join(folder_path, "*.txt"))
    
    if not txt_files:
        print(f"No TXT files found in '{folder_path}'.")
        return

    print(f"Found {len(txt_files)} files in Funds folder. Checking for new files...\n")
    new_funds_count = 0

    for file_path in txt_files:
        filename = os.path.basename(file_path)
        
        if is_file_processed(filename):
            print(f"⏭️  Skipping {filename} (Already processed)")
            continue
            
        print(f"📄 Processing funds file: {filename}...")
        
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
            
        structured_data = parse_funds_data_with_groq(text)
        
        if structured_data:
            save_funds_to_excel(structured_data)
            log_processed_file(filename)
            new_funds_count += len(structured_data)
            print(f"  ✅ Success! Added {len(structured_data)} fund records to Excel.\n")
        else:
            print(f"  ❌ Failed to parse data from {filename}.\n")
            
        time.sleep(3) 

    print(f"--- Funds Pipeline Complete ---")
    print(f"Added {new_funds_count} new deposit records to {EXCEL_FILE}.")


if __name__ == "__main__":
    trade_folder = "Trade_Confirmations"
    div_folder = "Dividends"
    funds_folder = "Funds_Transfers"  # <--- NEW FOLDER
    
    for folder in [trade_folder, div_folder, funds_folder]:
        if not os.path.exists(folder):
            os.makedirs(folder)
            print(f"Created folder '{folder}'.")

    setup_excel_database()
    
    print("========================================")
    print("  STARTING TRADE CONFIRMATION PIPELINE  ")
    print("========================================")
    run_automation_pipeline(trade_folder)
    
    print("========================================")
    print("      STARTING DIVIDEND PIPELINE        ")
    print("========================================")
    run_dividend_pipeline(div_folder)

    print("========================================")
    print("        STARTING FUNDS PIPELINE         ")
    print("========================================")
    run_funds_pipeline(funds_folder) # <--- RUN NEW PIPELINE
    
    print("\n========================================")
    print("          FINALIZING DATABASE           ")
    print("========================================")
    sort_all_sheets_by_date()