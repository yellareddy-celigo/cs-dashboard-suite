# CS Dashboard Suite

## 🎯 **Overview**

A comprehensive suite of tools for analyzing Customer Success (CS) data and creating professional dashboards. This repository provides multiple approaches for data analysis, from Google Sheets integration to direct JIRA MCP server access, with Excel export capabilities and specialized holiday season analysis.

## 🚀 **Key Features**

### **Multiple Data Sources**
- ✅ **Google Sheets Integration** - Read from existing Google Sheets
- ✅ **JIRA MCP Server** - Direct access to JIRA data via MCP
- ✅ **Manual Data Upload** - Support for CSV and manual data entry

### **Dashboard Creation**
- 📊 **Clear Data Dashboards** - Professional Excel dashboards
- 🎨 **Visual Elements** - ASCII charts and color coding
- 📈 **Multiple Analysis Views** - Executive, detailed, and raw data views
- 🎄 **Holiday Season Analysis** - Specialized holiday period analysis

### **Export Formats**
- 📁 **Excel Files** - Native Excel with formatting and charts
- 📋 **Google Sheets** - Direct integration with Google Sheets
- 📊 **CSV Export** - Raw data export capabilities

## 📁 **Repository Structure**

### **Core Analysis Tools**
```
├── jira_mcp_excel_analyzer_fixed.py    # Main JIRA MCP analyzer
├── create_advanced_dashboard.py        # Advanced dashboard creator
├── create_simple_dashboard.py          # Simple dashboard creator
├── run_real_jira_analysis.py           # Real JIRA data analysis
└── run_jira_analysis_fixed.py          # Sample data analysis
```

### **Google Sheets Integration**
```
├── jira_sheets_automation.py           # Google Sheets automation
├── update_google_sheets.py             # Direct JIRA to Sheets
├── scheduler.py                         # Automated scheduling
└── setup_automation.sh                 # Setup script
```

### **Documentation**
```
├── README.md                           # This file
├── README_JIRA_MCP_EXCEL.md           # JIRA MCP + Excel guide
├── README_MANUAL_UPLOAD.md            # Manual upload guide
└── SETUP.md                           # Setup guide
```

### **Configuration & Setup**
```
├── requirements.txt                    # Python dependencies
├── config.properties.template         # Configuration template
├── service_account_key.json.template  # Google Sheets auth template
└── .gitignore                         # Git ignore rules
```

## 🛠️ **Installation & Setup**

### **Prerequisites**
```bash
# Install Python dependencies
pip3 install -r requirements.txt
```

### **Quick Setup**
1. Clone the repository
2. Install dependencies: `pip3 install -r requirements.txt`
3. Copy templates: `cp config.properties.template config.properties`
4. Configure your Google Sheets or JIRA MCP access
5. Run analysis: `python3 create_simple_dashboard.py`

See [SETUP.md](SETUP.md) for detailed setup instructions.

## �� **Usage Examples**

### **1. Create Clear Dashboard from Google Sheets**
```bash
# Simple dashboard
python3 create_simple_dashboard.py

# Advanced dashboard with visual elements
python3 create_advanced_dashboard.py
```

### **2. Analyze JIRA Data with MCP**
```bash
# Real JIRA data analysis
python3 run_real_jira_analysis.py

# Sample data analysis
python3 run_jira_analysis_fixed.py
```

### **3. Google Sheets Automation**
```bash
# Manual upload mode
python3 jira_sheets_automation.py

# Direct JIRA integration
python3 update_google_sheets.py

# Scheduled execution
python3 scheduler.py
```

## 🎨 **Dashboard Features**

### **Clear Data Dashboards**
- **Executive Dashboard** - Key metrics and KPIs
- **Detailed Analysis** - Comprehensive breakdown
- **Raw Data View** - All issues with color coding
- **Visual Charts** - ASCII charts and visualizations
- **Monthly Trends** - Time-based analysis

### **Holiday Season Analysis**
- **Black Friday Week** (Nov 20-27)
- **Cyber Monday** (Nov 27-Dec 1)
- **Holiday Shopping** (Dec 1-24)
- **Christmas Week** (Dec 24-Jan 1)
- **New Year Recovery** (Jan 1-15)

### **Visual Elements**
- 🟨 **ASCII Charts** - Visual bar charts
- 🎨 **Color Coding** - CS (blue) vs PRE (yellow)
- 📊 **Professional Formatting** - Headers, borders, styling
- 📈 **Trend Analysis** - Monthly and yearly patterns

## 📊 **Sample Output**

### **Dashboard Metrics**
```
🎉 SUCCESS! Clear dashboard created!
📊 Analyzed 118 total issues
   • CS Issues: 81 (68.6%)
   • PRE Issues: 37 (31.4%)
   • Resolution Rate: 30.0%
   • Recent Issues (30d): 10
```

### **Excel File Structure**
- **🎯 Executive Dashboard** - Key metrics overview
- **📊 Detailed Analysis** - CS vs PRE comparison
- **📋 Raw Data** - All issues with details
- **📈 Charts & Visualizations** - Visual insights

## 🔧 **Customization**

### **Adding New Analysis Metrics**
```python
# In the analyzer class
def generate_analysis(self):
    analysis = {}
    # Add your custom metrics here
    analysis['custom_metric'] = your_calculation()
    return analysis
```

### **Modifying Dashboard Layout**
```python
# In the dashboard creator
def create_excel_dashboard(self, analysis):
    # Customize sheet names, headers, and layout
    ws_custom = wb.create_sheet("Custom Analysis")
    # Add your custom data and formatting
```

## 🚀 **Integration Examples**

### **JIRA MCP Integration**
```python
from jira_mcp_excel_analyzer_fixed import JiraMCPExcelAnalyzer

# Initialize analyzer
analyzer = JiraMCPExcelAnalyzer()

# Process JIRA data
df = analyzer.process_jira_issues(jira_issues)

# Generate analysis
analysis = analyzer.generate_holiday_analysis(df)

# Export to Excel
filename = analyzer.export_to_excel(df, analysis)
```

### **Google Sheets Integration**
```python
import gspread
from google.oauth2.service_account import Credentials

# Connect to Google Sheets
scopes = ['https://www.googleapis.com/auth/spreadsheets']
credentials = Credentials.from_service_account_file('service_account_key.json', scopes=scopes)
gc = gspread.authorize(credentials)

# Access your sheet
sheet = gc.open_by_key('YOUR_SHEET_ID')
```

## 📈 **Use Cases**

### **Executive Reporting**
- Monthly issue distribution analysis
- Team workload distribution
- Resolution rate tracking
- Holiday season impact analysis

### **QA Planning**
- Peak period identification
- Resource allocation planning
- Trend analysis for capacity planning
- Historical pattern recognition

### **Team Management**
- Assignee workload analysis
- Priority distribution tracking
- Status progression monitoring
- Performance metrics dashboard

## 🔄 **Workflow Examples**

### **Daily Analysis Workflow**
1. Run `create_simple_dashboard.py` for quick overview
2. Review Executive Dashboard for key metrics
3. Check Detailed Analysis for trends
4. Export specific data for further analysis

### **Weekly Reporting Workflow**
1. Run `create_advanced_dashboard.py` for comprehensive analysis
2. Review visual charts and trends
3. Generate reports for stakeholders
4. Update Google Sheets with new data

### **Monthly Deep Dive Workflow**
1. Run holiday season analysis
2. Generate comprehensive reports
3. Analyze long-term trends
4. Plan resource allocation

## 🛡️ **Security & Best Practices**

### **Data Protection**
- Service account keys are in `.gitignore`
- Sensitive data is not committed to repository
- Use environment variables for production

### **Error Handling**
- Comprehensive logging throughout
- Graceful error handling
- Data validation and sanitization

### **Performance**
- Efficient data processing
- Memory optimization for large datasets
- Caching for repeated operations

## 📞 **Support & Contributing**

### **Getting Help**
- Check the documentation in each script
- Review the README files for specific features
- Look at the sample data and output examples

### **Contributing**
- Fork the repository
- Create feature branches
- Submit pull requests
- Follow the existing code style

### **Issues & Bugs**
- Report issues in the GitHub repository
- Include sample data and error messages
- Provide steps to reproduce

## 🎯 **Future Enhancements**

### **Planned Features**
- [ ] Real-time dashboard updates
- [ ] Additional chart types and visualizations
- [ ] Automated report generation
- [ ] Integration with more data sources
- [ ] Machine learning for pattern recognition

### **Potential Integrations**
- [ ] Slack notifications
- [ ] Email reporting
- [ ] Web dashboard interface
- [ ] API endpoints for external access

## 📄 **License**

This project is open source and available under the MIT License.

---

**Ready to analyze your CS data? Choose your preferred method and start generating insights!** 🚀
