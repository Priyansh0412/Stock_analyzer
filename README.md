# Accurate NSE Stock Analyzer

A **Python-based project** to fetch, analyze, and report live stock data from multiple sources such as **Yahoo Finance** and **NSE Official**. The tool calculates price ranges over **52-week, 3-month, and 1-month periods** (high/low), fetches the **current price**, and generates a clean **Excel report** for further analysis.

---

## Project Brief

The Accurate NSE Stock Analyzer allows users to:

- Fetch **52-week high/low** value of a stock (Period 1)  
- Fetch **3-month high/low** value of a stock (Period 2)  
- Fetch **1-month high/low** value of a stock (Period 3)  
- Fetch **Current Price**  
- Calculate **Current Price with respect to above three periods**

**Stocks in the sample list:**

- Idea  
- Adani Industries  
- Reliance Industries  
- Bajaj Auto  

This project is suitable for **investors, traders**, and anyone interested in analyzing **NSE stock trends** quickly.

---

## Features

- Fetch live stock data using **yfinance**  
- Calculate:  
  - Current Price  
  - 52-Week High / Low  
  - 3-Month High / Low  
  - 1-Month High / Low  
  - Position percentages (vs 52W, 3M, 1M, and combined "Current vs All")  
- Generate **Excel reports** with:  
  - Borders  
  - Number formatting  
  - Auto-adjusted column widths  
- Display stock analysis summary in the console

---

- Usage & Output

The console will display a summary for each stock, including:

Current Price

52-Week, 3-Month, and 1-Month Ranges

Position percentages

Data source

An Excel report will be generated automatically in the project folder:

Stock_Analysis_Report_<YYYYMMDD_HHMMSS>.xlsx

Output Excel Columns

Symbol

Price (₹)

52W Low / High

3M % / 1M %

Current vs All (%)

Data Source

Example Code Snippet
# Sample usage
from stock_analyzer import AccurateNSEStockAnalyzer

analyzer = AccurateNSEStockAnalyzer()
stock_dict = {
    'IDEA': 'Vodafone Idea Limited',
    'ADANIPORTS': 'Adani Ports and SEZ',
    'RELIANCE': 'Reliance Industries',
    'BAJAJ-AUTO': 'Bajaj Auto Limited'
}

df = analyzer.analyze_stocks(stock_dict)
analyzer.create_excel_with_speedometer(df)

##Installations

1. Create a virtual environment**
python -m venv .venv

2. **Run a virtual environment**
.venv\Scripts\activate

3. **Install required packages**
pip install -r requirements.txt

4.Running the Project
python stock_analyzer.py


