# JIRA MCP + Excel Analysis Tool

## 🎯 **Overview**

This is a **JIRA MCP + Excel Analysis Tool** that combines the power of JIRA MCP server integration with comprehensive Excel export capabilities. It's designed to analyze Customer Success cases with a special focus on **holiday season pattern recognition** as described in the original Historical CS Analysis project.

## 🚀 **Key Features**

### **Direct JIRA Integration**
- ✅ **No API tokens needed** - Uses JIRA MCP server
- ✅ **Real-time data access** - Fetches live JIRA data
- ✅ **Multiple query support** - Customer Success, PRE, custom queries
- ✅ **Automatic pagination** - Handles large datasets

### **Holiday Season Analysis**
- 🎄 **Black Friday Week** (Nov 20-27)
- 🎄 **Cyber Monday** (Nov 27-Dec 1)  
- 🎄 **Holiday Shopping** (Dec 1-24)
- 🎄 **Christmas Week** (Dec 24-Jan 1)
- 🎄 **New Year Recovery** (Jan 1-15)

### **Excel Export with Multiple Sheets**
- 📊 **Raw Data** - All JIRA issues with full details
- 🎄 **Holiday Season Analysis** - Focus on holiday periods
- 📈 **General Analysis** - Status, priority, assignee distributions
- 📋 **Executive Dashboard** - Key metrics and insights

### **Advanced Analytics**
- 📊 **Statistical Analysis** - Resolution rates, trends, patterns
- 🎯 **Holiday Season Impact** - Peak period analysis
- ⏱️ **Resolution Time Analysis** - Performance metrics
- 📈 **Trend Analysis** - Monthly, quarterly, yearly patterns

## 🛠️ **Installation & Setup**

### **Prerequisites**
```bash
# Install required packages
pip3 install pandas openpyxl xlsxwriter numpy python-dateutil

# Or install from requirements.txt
pip3 install -r requirements.txt
```

### **No Configuration Required**
Unlike the original project, this tool requires **no configuration files** or API tokens. It uses the JIRA MCP server directly.

## 📋 **Usage**

### **1. Basic Analysis**
```bash
# Run with sample data (for testing)
python3 run_jira_analysis_fixed.py

# Run with real JIRA data
python3 run_real_jira_analysis.py
```

### **2. Custom JIRA Queries**
The tool can be easily modified to use different JIRA queries:

```python
# Customer Success cases
jql = "project = 'Customer Success' AND type = Case ORDER BY created DESC"

# PRE cases  
jql = "project = PRE AND type = Bug ORDER BY created DESC"

# Holiday season specific
jql = "project = 'Customer Success' AND created >= '2024-11-01' AND created <= '2025-01-15'"
```

### **3. MCP Integration**
To use with the JIRA MCP server:

```python
from jira_mcp_excel_analyzer_fixed import JiraMCPExcelAnalyzer

# Initialize analyzer
analyzer = JiraMCPExcelAnalyzer()

# Process JIRA data (from MCP)
df = analyzer.process_jira_issues(jira_issues)

# Generate analysis
analysis = analyzer.generate_holiday_analysis(df)

# Export to Excel
filename = analyzer.export_to_excel(df, analysis)
```

## 📊 **Output Structure**

### **Excel File Contents**

#### **1. Raw Data Sheet**
- All JIRA issues with complete details
- Holiday season classification
- Calculated fields (resolution time, quarters, etc.)
- Color-coded holiday season rows

#### **2. Holiday Season Analysis Sheet**
- Holiday period breakdown
- Holiday vs off-season comparison
- Resolution time analysis by season
- Key performance indicators

#### **3. General Analysis Sheet**
- Status distribution
- Priority distribution  
- Top assignees
- Component analysis

#### **4. Executive Dashboard Sheet**
- Key metrics summary
- Holiday season impact
- Recent activity trends
- Quick navigation links

## 🎯 **Holiday Season Analysis Features**

### **Automatic Classification**
The tool automatically classifies issues into holiday periods:
- **Black Friday Week** - November 20-27
- **Cyber Monday** - November 27 - December 1
- **Holiday Shopping** - December 1-24
- **Christmas Week** - December 24 - January 1
- **New Year Recovery** - January 1-15
- **Off-Season** - All other periods

### **Holiday-Specific Metrics**
- Holiday season issue count and percentage
- Resolution time comparison (holiday vs off-season)
- Peak period impact analysis
- Seasonal trend identification

## 📈 **Sample Output**

```
🎉 SUCCESS! Real JIRA analysis completed!
📁 Excel file: jira_holiday_analysis_20250919_104942.xlsx
📊 Analyzed 10 real issues
🎄 Holiday season issues: 0
📈 Resolution rate: 30.0%
🔍 Open issues: 7
✅ Closed issues: 3
```

## 🔧 **Customization**

### **Adding New Holiday Periods**
```python
# In jira_mcp_excel_analyzer_fixed.py
self.holiday_periods = {
    'Black Friday Week': {'start': 'Nov 20', 'end': 'Nov 27'},
    'Cyber Monday': {'start': 'Nov 27', 'end': 'Dec 1'},
    # Add your custom periods here
    'Custom Period': {'start': 'Dec 15', 'end': 'Dec 25'}
}
```

### **Modifying Analysis Metrics**
```python
# Add custom metrics in generate_holiday_analysis()
analysis['custom_metric'] = your_calculation(df)
```

## 🚀 **Integration with MCP Server**

### **Using with JIRA MCP Server**
```python
# Fetch data from JIRA MCP
jira_issues = mcp_atlassian_searchJiraIssuesUsingJql(
    cloudId="your-cloud-id",
    jql="project = 'Customer Success' AND type = Case",
    maxResults=100
)

# Process with analyzer
analyzer = JiraMCPExcelAnalyzer()
df = analyzer.process_jira_issues(jira_issues['issues'])
analysis = analyzer.generate_holiday_analysis(df)
filename = analyzer.export_to_excel(df, analysis)
```

## 📋 **File Structure**

```
historicalCSAnalysisIA/
├── jira_mcp_excel_analyzer_fixed.py    # Main analyzer class
├── run_jira_analysis_fixed.py          # Sample data runner
├── run_real_jira_analysis.py           # Real data runner
├── jira_mcp_integration.py             # MCP integration example
├── requirements.txt                     # Dependencies
└── README_JIRA_MCP_EXCEL.md           # This file
```

## 🎯 **Benefits Over Original Project**

### **Advantages**
- ✅ **No configuration needed** - Works out of the box
- ✅ **Direct JIRA access** - No API token setup
- ✅ **Excel export** - Native Excel files with formatting
- ✅ **Real-time data** - Always current information
- ✅ **Holiday season focus** - Specialized analysis
- ✅ **Multiple sheets** - Organized output structure

### **Use Cases**
- 📊 **Executive reporting** - Holiday season impact analysis
- 🎯 **QA planning** - Identify peak period patterns
- 📈 **Performance tracking** - Resolution time analysis
- 🔍 **Trend analysis** - Historical pattern recognition
- 📋 **Team management** - Workload distribution analysis

## 🚀 **Next Steps**

1. **Run the analysis** with your JIRA data
2. **Customize the queries** for your specific needs
3. **Integrate with MCP** for automated reporting
4. **Schedule regular runs** for ongoing analysis
5. **Share insights** with your team

## 📞 **Support**

This tool is based on the original Historical CS Analysis project and extends it with JIRA MCP integration and Excel export capabilities. It maintains all the holiday season analysis features while providing a more streamlined user experience.

---

**Ready to analyze your JIRA data with holiday season focus? Run the tool and start generating insights!** 🎉
