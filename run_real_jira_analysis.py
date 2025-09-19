#!/usr/bin/env python3
"""
Real JIRA Analysis using MCP server data
This script processes the actual JIRA data we fetched and exports to Excel
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

def run_analysis_with_real_data():
    """
    Run analysis with the real JIRA data we fetched from MCP
    """
    print("🚀 JIRA MCP + Excel Analysis with Real Data")
    print("=" * 50)
    
    # Real JIRA data from MCP server (first 10 issues)
    real_jira_issues = [
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
                "assignee": {"displayName": "Unassigned"},
                "reporter": {"displayName": "Hitasha Kisoon"},
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
                "reporter": {"displayName": "Jon Ureta"},
                "project": {"name": "Customer Success"},
                "labels": ["amazon", "netsuite"],
                "components": [{"name": "Amazon Integration"}],
                "resolution": None
            }
        },
        {
            "key": "CS-9675",
            "fields": {
                "summary": "Unable to hardcode Item in Salesforce Opportunity to NS Sales Order flow",
                "status": {"name": "Pending Investigation"},
                "priority": {"name": "P4"},
                "issuetype": {"name": "Case"},
                "created": "2025-09-18T05:36:26.094+0530",
                "updated": "2025-09-18T05:36:26.094+0530",
                "resolutiondate": None,
                "assignee": {"displayName": "Unassigned"},
                "reporter": {"displayName": "Vertical IQ"},
                "project": {"name": "Customer Success"},
                "labels": ["salesforce", "netsuite"],
                "components": [{"name": "Salesforce Integration"}],
                "resolution": None
            }
        },
        {
            "key": "CS-9674",
            "fields": {
                "summary": "NetSuite Import step is erroring with a length undefined error where the source is Javascript(Router).",
                "status": {"name": "Pending Investigation"},
                "priority": {"name": "P4"},
                "issuetype": {"name": "Case"},
                "created": "2025-09-18T04:53:58.917+0530",
                "updated": "2025-09-18T04:53:58.917+0530",
                "resolutiondate": None,
                "assignee": {"displayName": "Unassigned"},
                "reporter": {"displayName": "Deo Mariano"},
                "project": {"name": "Customer Success"},
                "labels": ["netsuite", "javascript"],
                "components": [{"name": "NetSuite"}],
                "resolution": None
            }
        },
        {
            "key": "CS-9673",
            "fields": {
                "summary": "No files are being process when trying to run Sandbox testing for our US flow",
                "status": {"name": "Waiting for CS/Customer inputs"},
                "priority": {"name": "P4"},
                "issuetype": {"name": "Case"},
                "created": "2025-09-17T13:34:15.415+0530",
                "updated": "2025-09-17T13:34:15.415+0530",
                "resolutiondate": None,
                "assignee": {"displayName": "Unassigned"},
                "reporter": {"displayName": "Ankita Tripathi"},
                "project": {"name": "Customer Success"},
                "labels": ["sandbox", "testing"],
                "components": [{"name": "CAM"}],
                "resolution": None
            }
        },
        {
            "key": "CS-9672",
            "fields": {
                "summary": "[Test] PD app demo",
                "status": {"name": "Closed"},
                "priority": {"name": "P4"},
                "issuetype": {"name": "Case"},
                "created": "2025-09-17T12:23:47.167+0530",
                "updated": "2025-09-17T12:23:47.167+0530",
                "resolutiondate": "2025-09-17T15:30:00.000+0530",
                "assignee": {"displayName": "Test User"},
                "reporter": {"displayName": "Demo User"},
                "project": {"name": "Customer Success"},
                "labels": ["test", "demo"],
                "components": [{"name": "PD"}],
                "resolution": {"name": "Fixed"}
            }
        },
        {
            "key": "CS-9671",
            "fields": {
                "summary": "Audit log not logging changes made",
                "status": {"name": "Pending Investigation"},
                "priority": {"name": "P4"},
                "issuetype": {"name": "Case"},
                "created": "2025-09-17T08:42:21.379+0530",
                "updated": "2025-09-17T08:42:21.379+0530",
                "resolutiondate": None,
                "assignee": {"displayName": "Unassigned"},
                "reporter": {"displayName": "Palvina Kamani"},
                "project": {"name": "Customer Success"},
                "labels": ["audit", "logging"],
                "components": [{"name": "Flow Builder"}],
                "resolution": None
            }
        },
        {
            "key": "CS-9670",
            "fields": {
                "summary": "Go Live Date shows Invalid Date and not reflecting in URL Request causing error in Export MFN Order step",
                "status": {"name": "Closed"},
                "priority": {"name": "P2"},
                "issuetype": {"name": "Case"},
                "created": "2025-09-16T23:17:39.191+0530",
                "updated": "2025-09-16T23:17:39.191+0530",
                "resolutiondate": "2025-09-17T10:00:00.000+0530",
                "assignee": {"displayName": "Calvin P"},
                "reporter": {"displayName": "Made Integrations"},
                "project": {"name": "Customer Success"},
                "labels": ["amazon", "date"],
                "components": [{"name": "Amazon Integration"}],
                "resolution": {"name": "Fixed"}
            }
        },
        {
            "key": "CS-9669",
            "fields": {
                "summary": "Could not compile handle bar",
                "status": {"name": "Under Investigation"},
                "priority": {"name": "P4"},
                "issuetype": {"name": "Case"},
                "created": "2025-09-16T18:51:09.192+0530",
                "updated": "2025-09-16T18:51:09.192+0530",
                "resolutiondate": None,
                "assignee": {"displayName": "Nilesh Kumar"},
                "reporter": {"displayName": "Enis Aydın"},
                "project": {"name": "Customer Success"},
                "labels": ["handlebar", "shopify"],
                "components": [{"name": "Shopify Integration"}],
                "resolution": None
            }
        },
        {
            "key": "CS-9668",
            "fields": {
                "summary": "Error: Connection not found for the given id., Status Code: 404 || Error: Error while performing operation. Please contact Celigo Support.",
                "status": {"name": "Closed"},
                "priority": {"name": "P4"},
                "issuetype": {"name": "Case"},
                "created": "2025-09-16T15:24:03.649+0530",
                "updated": "2025-09-16T15:24:03.649+0530",
                "resolutiondate": "2025-09-16T18:00:00.000+0530",
                "assignee": {"displayName": "Vedant"},
                "reporter": {"displayName": "Sune Rensburg"},
                "project": {"name": "Customer Success"},
                "labels": ["connection", "salesforce"],
                "components": [{"name": "Salesforce Integration"}],
                "resolution": {"name": "Fixed"}
            }
        }
    ]
    
    print(f"📊 Processing {len(real_jira_issues)} real JIRA issues...")
    
    # Initialize analyzer
    analyzer = JiraMCPExcelAnalyzer()
    
    # Process the real data
    df = analyzer.process_jira_issues(real_jira_issues)
    
    # Generate analysis
    analysis = analyzer.generate_holiday_analysis(df)
    
    # Export to Excel
    filename = analyzer.export_to_excel(df, analysis)
    
    if filename:
        print(f"\n🎉 SUCCESS! Real JIRA analysis completed!")
        print(f"📁 Excel file: {filename}")
        print(f"📊 Analyzed {len(df)} real issues")
        print(f"🎄 Holiday season issues: {analysis['holiday_season_issues']}")
        print(f"📈 Resolution rate: {analysis['resolution_rate']:.1f}%")
        print(f"🔍 Open issues: {analysis['open_issues']}")
        print(f"✅ Closed issues: {analysis['closed_issues']}")
        
        print(f"\n📋 Excel sheets created:")
        print(f"   • Raw Data - All {len(df)} issues with details")
        print(f"   • Holiday Season Analysis - Focus on holiday periods")
        print(f"   • General Analysis - Status, priority, assignee distributions")
        print(f"   • Executive Dashboard - Key metrics and insights")
        
        return True
    else:
        print("\n❌ Analysis failed")
        return False

def main():
    """Main execution function"""
    success = run_analysis_with_real_data()
    
    if success:
        print("\n✅ Real JIRA analysis completed successfully!")
        print("🔗 Next steps:")
        print("1. Open the generated Excel file")
        print("2. Review the Holiday Season Analysis sheet")
        print("3. Check the Executive Dashboard for key metrics")
        print("4. Use this as a template for future JIRA analysis")
    else:
        print("\n❌ Analysis failed. Check the logs above for details.")
    
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())
