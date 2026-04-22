## ACC102--WRDS-SEC-EDGAR-Financial-Product
## WRDS & SEC EDGAR Financial Product
 
## 1. Problem & User
   *This interactive data analysis tool is designed for financial analysts, finance students, and investors. It solves the critical problem of fragmented financial data by integrating quantitative market data (WRDS), fundamental corporate data (SEC), and macroeconomic indicators (World Bank) into a single, seamless interactive platform.Users can use this tool for interactive visual WRDS data export, calculation of 20+ core financial indicators, and SEC EDGAR integration to crawl and parse 10‑K, 10‑Q, 8‑K filings and reconstruct US GAAP financial statements. It supports an interactive DCF model with real‑time OCF data and waterfall visualization, plus market benchmark analysis to compare target companies against major indices (SPY for S&P 500, QQQ for Nasdaq 100).


## 2. Data Sources
   
   *WRDS CRSP Database: Provides daily stock prices, returns, and trading volumes (`crsp.dsf`, `crsp.msenames`). Accessed dynamically via PostgreSQL.  
   *SEC EDGAR API: Enables live retrieval of corporate filings (10-K, 10-Q, 8-K), company facts (US GAAP metrics including Operating Cash Flow, Net Income, Revenue), and qualitative text insights (MD&A, Risk Factors).  
   *World Bank API: Supplies macroeconomic indicators (GDP Growth, CPI, Unemployment) for broader market context.  
   *Access Date: Live API and database retrieval at runtime.  

## 3. Methodology
   *Data Acquisition: Utilizes `sqlalchemy` for secure WRDS database queries and `requests` for RESTful API calls to SEC EDGAR and the World Bank.  
   *Data Processing: Employs `pandas` and `numpy` for missing value handling, forward-filling methodologies, and corporate action-adjusted price calculations.  
   *Financial Analytics: Computes advanced quantitative metrics, including Annualized Return, Volatility, VaR (Historical & Parametric), CVaR, Maximum Drawdown, Sharpe Ratio, and DuPont analysis.  
   *Valuation Modeling: Implements a dynamic Discounted Cash Flow (DCF) model, computing the Present Value of future free cash flows and Terminal Value to derive an asset's Intrinsic Value.  
   *Automated Reporting: Streamlines the creation of polished `.xlsx` financial models (`openpyxl`) and `.docx` investment reports (`python-docx`).  
   *Visualization: Uses `plotly` to render interactive line charts, bar charts, correlation heatmaps, and DCF waterfall charts.  

## 4. Key Findings
   *Holistic Risk-Adjusted Assessment: The platform successfully demonstrates that evaluating asset performance requires a synthesis of cumulative returns, rolling volatility, and maximum drawdown, rather than looking at isolated metrics.  
   *Qualitative Edge: Automated text-mining of SEC filings (MD&A and CEO quotes) provides critical qualitative context to purely quantitative metrics, offering a complete picture of corporate health.  
   *Valuation Disparities: The interactive DCF module frequently reveals actionable disparities between a stock's current trading price and its intrinsic value based on fundamental operating cash flows.  
   Market Comparison Analysis: The benchmarking module enables direct performance comparison between target companies and broad market indices (e.g., SPY for S&P 500, QQQ for Nasdaq 100). By combining interactive cumulative return charts and risk‑return scatter plots, it accurately calculates Alpha, Beta, and Tracking Error, allowing users to objectively identify whether a company outperforms or underperforms the overall market.

## 5. How to Run (Local Deployment)

   *Step 1: Download and Save the Project  
1. Download the project files (or clone the repository).  
2. Extract and save the project folder to a specific, easy-to-find location on your computer (For example, save it directly to the desktop. It is recommended to save files to the desktop for easy access.).

  *Step 2: Open Terminal and Navigate to the Folder
You must point your terminal to the exact location where you saved the folder.（For example, if you save the project folder on your computer desktop, the exact terminal path to the project is the desktop.）  

   *For macOS (Apple):
   1. Open the **Terminal** application (Press `Command + Space`, type `Terminal`, and press Enter).  
   2. Use the `cd` (change directory) command to navigate to your folder. Please refer to the instruction template:  
   cd/Users/"Username"/"The location of the project folder"/"Folder name"  
 
   For example, if your computer username is "wuduu", the project folder is located on the desktop, and the folder name is "Financial_App", type:
     
   cd /Users/wuduu/Desktop/Financial_App
  

   *For Windows Users:
   1. Open the **Command Prompt** (Press `Windows Key + R`, type `cmd`, and press Enter).  
   2. Use the `cd` command to navigate to your folder. For example:
   
   cd C:/Users/”YourUsername“/"The location of the project folder"/"Folder name"  

   For example, if your computer username is "Amy", the project folder is located on the desktop, and the folder name is "Financial_App", type:
   If files saved to the desktop are stored by default on the C drive, type:  
   cd C:\Users/Amy/Desktop/Financial_App  
   Otherwise,   
   
   cd (the storage location of desktop files, eg. C or D or F):\Users/Amy/Desktop/Financial_App   


   *Step 3: Install Required Dependencies** 
   Once your terminal is operating inside the correct project folder, install all necessary Python libraries by executing the following command:  

   pip install streamlit pandas numpy sqlalchemy plotly openpyxl python-docx wbdata sec-edgar-downloader kaleido psycopg2-binary  
   
   *Step 4: Run the Application
   After the installation process finishes completely, launch the application by running:  
   
   streamlit run app.py  
   
   This will start the local server and automatically open the interactive web interface in your default browser.  

   *Note: A valid WRDS (Wharton Research Data Services) account is required to log in via the app's sidebar to connect to the PostgreSQL database for market data retrieval.（After being redirected to the website, you need to log in to your WRDS account.). After entering your WRDS username and password, use the DUO application linked to your account. A login request may be sent to the DUO app on your bound mobile device. If a request appears, simply approve it to complete the login successfully.

## 6. Product Link / Demo
   App Link: [https://github.com/ZixuanLu24/ACC102--WRDS-SEC-EDGAR-Financial-Product.git]

   Demo Video: []
 
## 7. Limitations & Next Steps
   *The application depends heavily on the availability and uptime of the WRDS PostgreSQL database and SEC EDGAR API.  
   *Current SEC text extraction relies on basic regex and heuristic rules, which may miss nuanced qualitative statements or non-standard filing formats.  
   *Text embedded in images or complex nested HTML tables in SEC filings cannot be parsed, leading to occasional “Data not found” errors.  
   *The current f‑string SQL query format introduces potential SQL injection risks in enterprise environments.  
   *SEC EDGAR API rate limits may cause timeouts during bulk ticker requests.  

   Next Steps: 
   *Enhanced SEC Parsing  
   Integrate OCR (Tesseract / cloud vision APIs) to extract text from images in filings.  
Adopt advanced libraries like unstructured.io or sec-parser for better table and document layout analysis.
Use LLM-powered summarization to improve accuracy and readability of extracted content.  
  
   *Security Upgrade  
   Replace f‑string queries with SQLAlchemy parameterized statements to prevent SQL injection.  
   
   *API Stability  
Implement exponential backoff retry (via tenacity) and request throttling to handle SEC EDGAR rate limits.  
   
   *Advanced Analytics  
   Integrate financial NLP models (e.g., FinBERT) for precise sentiment analysis on SEC filings.  
   Add a portfolio optimization module supporting Efficient Frontier and Markowitz allocation.  
