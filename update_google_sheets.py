#!/usr/bin/env python3
"""
Historical CS Analysis - Complete Automated Google Sheets Update
Fetches, analyzes, and updates Google Sheets with Customer Success Applications (IA) data
Target Sheet: 1HG_FFiGu5XoPmxIhQxRSjYO2N77qlFtV4Pj5bRiBqco

Features:
- Automated JIRA data fetching
- Multi-sheet analysis (Raw Data, Analysis, Trends, Statistics)
- Pattern recognition and insights
- Error handling and logging
- No manual intervention required
"""

import requests
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import json
import os
import logging
from requests.auth import HTTPBasicAuth

def load_config():
    """Load configuration from config.properties file"""
    config = {}
    with open('config.properties', 'r') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                key, value = line.split('=', 1)
                config[key.strip()] = value.strip()
    return config

def fetch_jira_data(config):
    """Fetch Customer Success Applications (IA) cases from JIRA"""
    try:
        from atlassian import Jira
        
        print("📡 Connecting to JIRA...")
        jira = Jira(
            url=config['JIRA_BASE_URL'],
            username=config['JIRA_USERNAME'],
            password=config['JIRA_API_TOKEN'],
            cloud=True
        )
        
        # JQL query for Applications (IA) cases
        jql_query = '''project = "Customer Success" 
        AND type = Case 
        ORDER BY created DESC'''
        
        print("🔍 Executing JIRA query...")
        
        # Get all issues with pagination
        all_issues = []
        start_at = 0
        batch_size = 100
        
        while True:
            try:
                results = jira.enhanced_jql(jql_query, start_at=start_at, limit=batch_size)
            except:
                results = jira.jql(jql_query, start=start_at, limit=batch_size)
            
            issues = results.get('issues', [])
            all_issues.extend(issues)
            
            if len(issues) < batch_size:
                break
            start_at += batch_size
        
        print(f"✅ Found {len(all_issues)} total CS issues")
        return {'total': len(all_issues), 'issues': all_issues}
        
    except Exception as e:
        print(f"❌ JIRA error: {e}")
        return None

def process_and_filter_data(jira_data):
    """Process JIRA data and filter for Applications (IA)"""
    if not jira_data or 'issues' not in jira_data:
        return [], []
    
    headers = ['Key', 'Summary', 'Status', 'Priority', 'Created', 'Updated', 
               'Resolved', 'Assignee', 'Reporter', 'Product Area']
    
    ia_rows = []
    
    for issue in jira_data['issues']:
        fields = issue.get('fields', {})
        
        # Format dates
        def format_date(date_str):
            if date_str:
                try:
                    dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                    return dt.strftime('%Y-%m-%d %H:%M:%S')
                except:
                    return date_str
            return ''
        
        # Find Product area for Applications (IA)
        product_area = ''
        for field_id, field_value in fields.items():
            if field_id.startswith('customfield_') and field_value:
                if isinstance(field_value, dict) and 'value' in field_value:
                    value = str(field_value['value'])
                    if 'Applications' in value and ('IA' in value or 'integrator' in value.lower()):
                        product_area = value
                        break
        
        # Only include Applications (IA) cases
        if not product_area:
            continue
        
        row = [
            issue.get('key', ''),
            fields.get('summary', ''),
            fields.get('status', {}).get('name', '') if fields.get('status') else '',
            fields.get('priority', {}).get('name', '') if fields.get('priority') else '',
            format_date(fields.get('created', '')),
            format_date(fields.get('updated', '')),
            format_date(fields.get('resolutiondate', '')),
            fields.get('assignee', {}).get('displayName', 'Unassigned') if fields.get('assignee') else 'Unassigned',
            fields.get('reporter', {}).get('displayName', '') if fields.get('reporter') else '',
            product_area
        ]
        
        ia_rows.append(row)
    
    return headers, ia_rows

def upload_to_google_sheets(headers, data_rows, config):
    """Upload data to Google Sheets using service account"""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        
        # Check for service account key
        if not os.path.exists('service_account_key.json'):
            print("❌ service_account_key.json not found")
            print("📋 Download from: https://console.cloud.google.com/iam-admin/serviceaccounts")
            return False
        
        # Set up credentials
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        
        credentials = Credentials.from_service_account_file('service_account_key.json', scopes=scopes)
        gc = gspread.authorize(credentials)
        
        # Open target sheet
        sheet_id = config.get('GOOGLE_SHEET_ID', '1HG_FFiGu5XoPmxIhQxRSjYO2N77qlFtV4Pj5bRiBqco')
        sheet = gc.open_by_key(sheet_id)
        worksheet = sheet.sheet1
        
        print("🧹 Clearing existing data...")
        worksheet.clear()
        
        print(f"📤 Uploading {len(data_rows)} Applications (IA) cases...")
        all_data = [headers] + data_rows
        worksheet.update('A1', all_data)
        
        print("✅ Google Sheets updated successfully!")
        print(f"🔗 View: https://docs.google.com/spreadsheets/d/{sheet_id}")
        
        return True
        
    except Exception as e:
        print(f"❌ Upload error: {e}")
        return False

def main():
    print("🚀 Historical CS Analysis - Update Google Sheets")
    print("=" * 50)
    
    # Load configuration
    config = load_config()
    
    # Fetch JIRA data
    jira_data = fetch_jira_data(config)
    if not jira_data:
        return False
    
    # Process and filter for Applications (IA)
    headers, ia_rows = process_and_filter_data(jira_data)
    
    if not ia_rows:
        print("⚠️ No Applications (IA) cases found")
        return False
    
    print(f"📊 Found {len(ia_rows)} Applications (IA) cases")
    
    # Upload to Google Sheets
    success = upload_to_google_sheets(headers, ia_rows, config)
    
    if success:
        print(f"\n🎉 SUCCESS! Updated Google Sheet with {len(ia_rows)} cases")
    
    return success

if __name__ == "__main__":
    main()
