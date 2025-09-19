# CS Dashboard Suite - Setup Guide

## 🚀 **Quick Start**

### **1. Clone the Repository**
```bash
git clone https://github.com/yellareddy-celigo/cs-dashboard-suite.git
cd cs-dashboard-suite
```

### **2. Install Dependencies**
```bash
pip3 install -r requirements.txt
```

### **3. Configure Authentication**

#### **For Google Sheets Integration:**
1. Copy the template: `cp service_account_key.json.template service_account_key.json`
2. Create a Google Cloud Project and enable Google Sheets API
3. Create a service account and download the key
4. Replace the content in `service_account_key.json` with your actual service account key
5. Share your Google Sheet with the service account email

#### **For JIRA MCP Integration:**
- No additional setup required
- Uses MCP server for direct JIRA access

### **4. Configure Settings**
```bash
cp config.properties.template config.properties
# Edit config.properties with your values
```

### **5. Run Analysis**
```bash
# Simple dashboard from Google Sheets
python3 create_simple_dashboard.py

# Advanced dashboard with visual elements
python3 create_advanced_dashboard.py

# JIRA MCP analysis
python3 run_real_jira_analysis.py
```

## 📋 **Configuration Details**

### **Google Sheets Setup**
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing
3. Enable Google Sheets API
4. Create credentials > Service Account
5. Download the JSON key file
6. Replace `service_account_key.json` with your key
7. Share your Google Sheet with the service account email

### **JIRA MCP Setup**
- No configuration needed
- Works with MCP server integration
- Direct access to JIRA data

## 🔧 **Troubleshooting**

### **Common Issues**
- **Google Sheets Access Denied**: Make sure the service account email has access to your sheet
- **JIRA MCP Not Working**: Ensure MCP server is properly configured
- **Missing Dependencies**: Run `pip3 install -r requirements.txt`

### **File Permissions**
- Make sure `service_account_key.json` has proper permissions
- Check that `config.properties` is readable

## 📊 **Sample Data**
- The repository includes sample data for testing
- Use `run_jira_analysis_fixed.py` for sample data analysis
- Generated Excel files will be in the same directory

## 🎯 **Next Steps**
1. Run the analysis scripts
2. Review the generated Excel dashboards
3. Customize the analysis for your needs
4. Set up automated scheduling if needed
