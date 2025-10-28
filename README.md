# Historical CS Analysis Tool

## Overview

This tool provides comprehensive analysis of Customer Success (CS) cases with Integration Applications (IA) focus. It generates detailed reports with customer extraction, resolution analysis, and pattern identification.

## 📋 How to Get Started

### Step 1: Export Data from JIRA

Use this JQL query to fetch CS cases with Integration Apps:

```jql
project = CS AND 
(created >= "2024-12-01" AND created < "2025-01-01") AND 
"integration apps[dropdown]" IN (
  "Amazon - Acumatica IA", "Amazon - MS Dynamics Business Central 365", 
  "Amazon - NetSuite IA", "Amazon - NetSuite IA (SS)", 
  "Amazon - SAP Business ByDesign", "Amazon MCF - NetSuite IA", 
  "BigCommerce - NetSuite IA", "BigCommerce - SAP Business ByDesign IA", 
  "BigCommerce - MS Dynamics 365 Business Central", 
  "Cash Application Manager (IO)", "eBay - NetSuite IA (IO)", 
  "eBay - NetSuite IA (SS)", "Magento 1 - NetSuite IA (SS)", 
  "Magento 2 - NetSuite IA", " Magento 2 - SAP Business ByDesign IA", 
  "OpenAir - Salesforce IA", "Orderful - NetSuite IA", 
  "Salesforce - MS Dynamics Business Central 365", 
  "Salesforce - NetSuite IA (IO)", 
  "Salesforce - NetSuite IA (SS - v1 & v2)", 
  "Salesforce - SAP Business ByDesign", 
  "Shopify - MS Dynamics Business Central 365", "Shopify - NetSuite IA", 
  " Shopify - SAP Business ByDesign", "Square - NetSuite IA", 
  "Vendor Payment Manager", "Walmart - NetSuite IA", "Zendesk - NetSuite IA", 
  "Payout to Reconciliation IA", "Reconciliation Process"
)
```

**Note**: Change the dates (`created >= "2024-12-01" AND created < "2025-01-01"`) to match your desired time period.

Example queries for different time periods:
- **December 2024**: `created >= "2024-12-01" AND created < "2025-01-01"`
- **Full Year 2024**: `created >= "2024-01-01" AND created < "2025-01-01"`
- **November 2024**: `created >= "2024-11-01" AND created < "2024-12-01"`

### Step 2: Download and Save CSV

1. Click **Export** in JIRA
2. Select **CSV (all fields)**
3. Save the CSV file to this repository
4. Name it appropriately (e.g., `2024.csv`, `December-2024.csv`, `Holiday-2024.csv`)

### Step 3: Generate Master Report

Run the master report generator:

```bash
python3 generate_master_report.py --file <file_name>.csv --output <Output_Report>.xlsx
```

#### Examples:

```bash
# Generate report for 2024 data
python3 generate_master_report.py --file 2024.csv --output 2024_Master_Report.xlsx

# Generate report for December 2024 data
python3 generate_master_report.py --file December-2024.csv --output December_2024_Report.xlsx

# Generate report for holiday season data
python3 generate_master_report.py --file Holiday.csv --output Holiday_Season_2024_Report.xlsx
```

## 📊 What the Master Report Contains

Each master report includes **16 beautifully formatted Excel sheets**:

### Executive Summary
- Overall Summary with dynamic filename
- Total cases analyzed
- Priority breakdown
- Resolution distribution

### Integration Analysis  
- Integration Overview
- Count by Integration
- Resolution by Integration App

### Temporal Analysis
- Count by Month

### Customer Analysis (Enhanced)
- Customer Analysis Enhanced (with extracted customer info)
- Customer Analysis Original

### Technical Details
- Frequent Flow Issues
- Error Categories
- Error Distribution
- Recurring Errors

### Pattern Analysis
- Pattern Analysis
- Top Integrations

### Code Fix Analysis
- Code Fix with Links

### Complete Data
- **All Cases Summary** - High-level summary of all cases
- **Complete Case Details** - Full detailed information including:
  - Full Resolution Comments (not truncated)
  - Issue identification
  - How the issue was fixed
  - All related details

## ✨ Key Features

- ✓ **Full Resolution Comments** - Complete text without truncation (up to full length)
- ✓ **Customer Extraction** - Automatically extracts customer names from descriptions
- ✓ **Dynamic Filename** - Overall Summary shows actual source file name
- ✓ **Professional Formatting** - Dark blue headers, alternating row colors
- ✓ **Auto-filters** - Enabled on all sheets for easy filtering
- ✓ **Freeze Panes** - Row 1 and column A frozen for navigation
- ✓ **Complete Case Details** - Every case with full context and resolution information

## 📈 Report Statistics

The tool automatically provides:

- **Customer Extraction Success Rate** - Shows percentage of cases where customer info was extracted
- **Unique Customers Identified** - Count of different customers in the data
- **Total Cases Analyzed** - Complete case count
- **Pattern Distribution** - Error categories and their frequencies
- **Integration Breakdown** - Cases by integration application
- **Resolution Methods** - How issues were addressed

## 🛠️ Available Scripts

### Core Scripts (6 Essential Files)

1. **generate_master_report.py** - Main orchestrator
   - Generates comprehensive master reports
   - Calls other analysis scripts
   - Combines results into single Excel file

2. **analyze_combined_report.py** - Comprehensive analysis
   - Overall Summary sheet
   - All Cases Summary
   - Pattern Analysis
   - Top Integrations

3. **deep_dive_detailed_analysis.py** - Deep dive analysis
   - Complete Case Details (with full Resolution Comments)
   - Integration Overview
   - Error Categories
   - Frequent Flow Issues

4. **create_dynamic_dashboards.py** - Management dashboards
   - Standalone dashboard generator
   - Uses CSV directly

5. **holiday_resolution_analysis.py** - Detailed holiday analysis
   - Analyzes Resolution Comments
   - Provides specific recommendations
   - Preventive actions

6. **simplified_holiday_analysis.py** - Simplified holiday analysis
   - Individual case analysis
   - Holiday-specific insights

## 📝 Sample Usage

### Example 1: Generate Annual Report

```bash
python3 generate_master_report.py --file 2024.csv --output 2024_Master_Report.xlsx
```

**Result**: 16-sheet comprehensive report for 2024 data

### Example 2: Analyze Holiday Season Issues

```bash
python3 generate_master_report.py --file Holiday.csv --output Holiday_Season_2024_Master_Report.xlsx
```

**Result**: Detailed analysis of holiday season cases with specific recommendations

### Example 3: Monthly Analysis

```bash
python3 generate_master_report.py --file December-2024.csv --output December_2024_Master_Report.xlsx
```

**Result**: Month-specific breakdown and analysis

## 🎯 What You'll Get

Each report provides:

1. **Complete Case Details** with full Resolution Comments
2. **Customer Extraction** from descriptions (90%+ success rate)
3. **Issue Analysis** - What the issue was
4. **Resolution Information** - How it was fixed
5. **Pattern Recognition** - Recurring issues and patterns
6. **Integration-Specific Insights** - Analysis by integration app
7. **Temporal Trends** - Monthly breakdown
8. **Error Categories** - Classification of issues
9. **Code Fix Tracking** - Cases requiring code changes
10. **Preventive Actions** - Recommendations to prevent future issues

## 📊 Output Example

When you run the tool, you'll see:

```
✅ MASTER REPORT GENERATED SUCCESSFULLY!

📁 Output file: 2024_Master_Report.xlsx

📋 Master report contains 16 beautifully formatted sheets:
  📊 EXECUTIVE SUMMARY
  🔧 INTEGRATION ANALYSIS
  📅 TEMPORAL ANALYSIS
  👥 CUSTOMER ANALYSIS (ENHANCED)
  🔍 TECHNICAL DETAILS
  📈 PATTERN ANALYSIS
  ✅ CODE FIX ANALYSIS
  📝 COMPLETE DATA

📊 CUSTOMER EXTRACTION SUMMARY:
   • Total cases analyzed: 469
   • Customer info extracted: 444 cases (94.7%)
   • Unique customers identified: 318
```

## 🔧 Requirements

- Python 3.7+
- pandas
- openpyxl
- See `requirements.txt` for complete list