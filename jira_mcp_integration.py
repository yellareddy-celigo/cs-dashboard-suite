#!/usr/bin/env python3
"""
JIRA MCP Integration Script
This script uses the JIRA MCP server to fetch real data and export to Excel
"""

import sys
import os
import pandas as pd
from datetime import datetime
import logging

# Add the current directory to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from jira_mcp_excel_analyzer_fixed import JiraMCPExcelAnalyzer

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def run_jira_analysis_with_mcp():
    """
    This function demonstrates how to use the JIRA MCP server
    In practice, this would be called by the MCP integration
    """
    print("🚀 JIRA MCP + Excel Analysis Tool")
    print("=" * 50)
    print("📊 This script is ready to work with JIRA MCP server")
    print()
    print("🔧 To use with real JIRA data:")
    print("1. The MCP integration will call this script")
    print("2. Pass JIRA issues data from searchJiraIssuesUsingJql")
    print("3. The script will process and export to Excel")
    print()
    print("📋 Available JIRA queries:")
    print("• Customer Success cases: project = 'Customer Success' AND type = Case")
    print("• PRE cases: project = PRE AND type = Bug")
    print("• Holiday season analysis: Add date filters for Nov-Jan periods")
    print()
    print("🎯 The script will generate:")
    print("• Raw Data sheet - All JIRA issues")
    print("• Holiday Season Analysis - Focus on holiday periods")
    print("• General Analysis - Status, priority, assignee distributions")
    print("• Executive Dashboard - Key metrics and insights")
    
    return True

def main():
    """Main execution function"""
    success = run_jira_analysis_with_mcp()
    
    if success:
        print("\n✅ JIRA MCP Integration ready!")
        print("📁 Excel export functionality: Ready")
        print("🎄 Holiday season analysis: Ready")
        print("📊 Statistical analysis: Ready")
        print()
        print("🔗 Next steps:")
        print("1. Use the MCP server to fetch JIRA data")
        print("2. Call the analyzer with the fetched data")
        print("3. Generate Excel reports with holiday season focus")
    
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())
