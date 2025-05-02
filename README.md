# Risk-Analyzer

Risk Analyzer is a data-driven decision-making tool designed to help Small and Medium-sized Businesses (SMBs) evaluate whether to use Original Equipment Manufacturer (OEM) services or develop in-house manufacturing capabilities. The tool analyzes multiple risk factors including cost, time, quality, scalability, and market risks to provide actionable recommendations.

Key Features: 
1. Comprehensive Risk Analysis - Evaluates 8 key risk factors with weighted scoring
2. Cost Comparison - Calculates total costs for in-house manufacturing vs OEM
3. Financial Risk Assessment - Incorporates beta calculations from market data
4. Automated Recommendations - Provides clear guidance based on analysis results
5. User-friendly Interface: Simple data input through Google Sheets sidebar

Technical Implementation:
1. Built using Google Apps Script (JavaScript).
2. Integrates with Google Sheets for data input and analysis.
3. Uses Google Finance API for market data.
4. Modular design with separate HTML forms for different input types.

How to Use
1. Set up the script (RiskAnalyzer.gs) in Google Sheets by pasting the code into the Apps Script editor.
2. Input your data using the custom menu options.
3. Cost attributes for in-house manufacturing, OEM costs and Company financial data for risk assessment.
6. Run the analysis to get recommendations.
