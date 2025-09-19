#!/usr/bin/env python3
"""
Main script to run JIRA analysis using MCP server
This script fetches data from JIRA using MCP and exports to Excel
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

def fetch_jira_data_via_mcp():
    """
    This function will be called by the MCP integration
    For now, it returns sample data structure
    """
    # This is where the MCP integration will provide the actual JIRA data
    # For demonstration, we'll show the structure expected
    
    print("📊 JIRA MCP Integration Ready")
    print("This script expects to receive JIRA data from the MCP server")
    print("The data should be in the format returned by searchJiraIssuesUsingJql")
    
    return None

def run_analysis_with_sample_data():
    """
    Run analysis with sample data for demonstration
    """
    print("🔧 Running with sample data for demonstration...")
    
    # Sample JIRA data structure (based on what we saw from MCP)
    sample_issues = [
        {
            "key": "CS-9679",
            "fields": {
                "summary": "The auto-billing flow is incorrectly applying tax towards Shipping Cost",
                "status": {"name": "Open"},
                "priority": {"name": "P4"},
                "issuetype": {"name": "Case"},
                "created": "2025-09-19T08:32:28.125+0530",
                "updated": "2025-09-19T08:32:28.125+0530",
                "resolutiondate": None,
                "assignee": {"displayName": "John Doe"},
                "reporter": {"displayName": "Jane Smith"},
                "project": {"name": "Customer Success"},
                "labels": ["integration", "billing"],
                "components": [{"name": "NetSuite"}],
                "resolution": None
            }
        },
        {
            "key": "CS-9678", 
            "fields": {
                "summary": "Amazon - NetSuite Integration App to support Item Groups",
                "status": {"name": "Open"},
                "priority": {"name": "P4"},
                "issuetype": {"name": "Case"},
                "created": "2025-09-19T02:14:27.748+0530",
                "updated": "2025-09-19T02:14:27.748+0530",
                "resolutiondate": None,
                "assignee": {"displayName": "Unassigned"},
                "reporter": {"displayName": "Bob Wilson"},
                "project": {"name": "Customer Success"},
                "labels": ["amazon", "netsuite"],
                "components": [{"name": "Amazon Integration"}],
                "resolution": None
            }
        }
    ]
    
    # Initialize analyzer
    analyzer = JiraMCPExcelAnalyzer()
    
    # Process the data
    df = analyzer.process_jira_issues(sample_issues)
    
    # Generate analysis
    analysis = analyzer.generate_holiday_analysis(df)
    
    # Export to Excel
    filename = analyzer.export_to_excel(df, analysis)
    
    if filename:
        print(f"\n🎉 SUCCESS! Analysis completed!")
        print(f"📁 Excel file: {filename}")
        print(f"📊 Analyzed {len(df)} issues")
        print(f"🎄 Holiday season issues: {analysis['holiday_season_issues']}")
        return True
    else:
        print("\n❌ Analysis failed")
        return False

def main():
    """Main execution function"""
    print("🚀 JIRA MCP + Excel Analysis Tool")
    print("=" * 50)
    
    # For now, run with sample data
    # In production, this would be called by the MCP integration
    success = run_analysis_with_sample_data()
    
    if success:
        print("\n✅ Analysis completed successfully!")
        print("📋 Next steps:")
        print("1. Open the generated Excel file")
        print("2. Review the Holiday Season Analysis sheet")
        print("3. Check the Executive Dashboard for key metrics")
    else:
        print("\n❌ Analysis failed. Check the logs above for details.")
    
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())
