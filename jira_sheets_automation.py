#!/usr/bin/env python3
"""
JIRA to Google Sheets Complete Automation
Automatically fetches, analyzes, and updates Google Sheets with no manual intervention
"""

import os
import sys
import json
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from atlassian import Jira
import gspread
from google.oauth2.service_account import Credentials
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class JiraGoogleSheetsAutomation:
    def __init__(self, config_file='config.properties'):
        """Initialize with configuration"""
        self.config = self.load_config(config_file)
        self.sheet_id = self.config.get('GOOGLE_SHEET_ID', '1HG_FFiGu5XoPmxIhQxRSjYO2N77qlFtV4Pj5bRiBqco')
        self.jira = None
        self.sheets_client = None
        self.data = None
        
    def load_config(self, config_file):
        """Load configuration from properties file"""
        config = {}
        try:
            with open(config_file, 'r') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and '=' in line:
                        key, value = line.split('=', 1)
                        config[key.strip()] = value.strip()
            logger.info("✅ Configuration loaded successfully")
            return config
        except Exception as e:
            logger.error(f"❌ Failed to load config: {e}")
            sys.exit(1)
    
    def connect_jira(self):
        """Connect to JIRA using Atlassian API"""
        try:
            self.jira = Jira(
                url=self.config['JIRA_BASE_URL'],
                username=self.config['JIRA_USERNAME'],
                password=self.config['JIRA_API_TOKEN'],
                cloud=True
            )
            logger.info("✅ Connected to JIRA successfully")
            return True
        except Exception as e:
            logger.error(f"❌ JIRA connection failed: {e}")
            return False
    
    def connect_google_sheets(self):
        """Connect to Google Sheets using service account"""
        try:
            if not os.path.exists('service_account_key.json'):
                logger.error("❌ service_account_key.json not found!")
                logger.info("📋 Download from: https://console.cloud.google.com")
                return False
            
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            credentials = Credentials.from_service_account_file('service_account_key.json', scopes=scopes)
            self.sheets_client = gspread.authorize(credentials)
            logger.info("✅ Connected to Google Sheets successfully")
            return True
        except Exception as e:
            logger.error(f"❌ Google Sheets connection failed: {e}")
            return False
    
    def fetch_manual_data_from_sheets(self):
        """Read manually uploaded JIRA data from Google Sheets instead of fetching from JIRA API"""
        try:
            logger.info("📊 Reading manually uploaded JIRA data from Google Sheets...")
            
            # Check if there's already data in the Raw Data sheet
            try:
                raw_sheet = self.spreadsheet.worksheet('Raw Data')
                logger.info("✅ Found existing 'Raw Data' sheet")
            except:
                logger.info("❌ No 'Raw Data' sheet found - please manually upload your JIRA data first")
                return False
            
            # Get all data from the Raw Data sheet
            try:
                all_data = raw_sheet.get_all_records()
                logger.info(f"📋 Found {len(all_data)} records in manually uploaded data")
                
                if len(all_data) == 0:
                    logger.info("⚠️ No data found in Raw Data sheet - please upload your JIRA Applications (IA) data first")
                    return False
                
                # Convert the sheet data to our expected format for analysis
                processed_data = []
                for row in all_data:
                    # Convert sheet row to our analysis format
                    processed_row = {
                        'JIRA ID': row.get('JIRA ID', row.get('Key', '')),
                        'JIRA Text/Summary': row.get('JIRA Text/Summary', row.get('Summary', '')),
                        'Description': row.get('Description', ''),
                        'Status': row.get('Status', ''),
                        'Priority': row.get('Priority', ''),
                        'Created Date': row.get('Created Date', row.get('Created', '')),
                        'Updated Date': row.get('Updated Date', row.get('Updated', '')),
                        'Resolved Date': row.get('Resolved Date', row.get('Resolved', '')),
                        'Assignee': row.get('Assignee', ''),
                        'Reporter': row.get('Reporter', ''),
                        'Product Area': row.get('Product Area', ''),
                        'Components': row.get('Components', ''),
                        'Labels': row.get('Labels', ''),
                    }
                    processed_data.append(processed_row)
                
                # Store the processed data for analysis
                self.manual_data = processed_data
                
                # Show some statistics about the uploaded data
                status_stats = {}
                product_area_stats = {}
                
                for row in processed_data:
                    status = row.get('Status', 'Unknown')
                    if status not in status_stats:
                        status_stats[status] = 0
                    status_stats[status] += 1
                    
                    product_area = row.get('Product Area', 'Unknown')
                    if product_area not in product_area_stats:
                        product_area_stats[product_area] = 0
                    product_area_stats[product_area] += 1
                
                logger.info(f"📊 Status distribution in uploaded data:")
                for status, count in sorted(status_stats.items(), key=lambda x: x[1], reverse=True):
                    logger.info(f"   • {status}: {count} cases")
                
                logger.info(f"📊 Product Area distribution in uploaded data:")
                for area, count in sorted(product_area_stats.items(), key=lambda x: x[1], reverse=True):
                    logger.info(f"   • {area}: {count} cases")
                
                logger.info(f"✅ Successfully loaded {len(processed_data)} manually uploaded Applications (IA) cases")
                return True
                
            except Exception as e:
                logger.error(f"❌ Error reading data from sheet: {e}")
                return False
                
        except Exception as e:
            logger.error(f"❌ Failed to read manual data from sheets: {e}")
            return False
    
    def process_data(self):
        """Process and analyze manually uploaded JIRA data"""
        if not hasattr(self, 'manual_data') or not self.manual_data:
            logger.error("❌ No manually uploaded data to process")
            return None
        
        # Convert manual data to pandas DataFrame for analysis
        df = pd.DataFrame(self.manual_data)
        
        # Add additional analysis columns if dates are present
        for idx, row in df.iterrows():
            created_date = row.get('Created Date', '')
            if created_date:
                try:
                    from datetime import datetime
                    if isinstance(created_date, str):
                        # Try to parse common date formats
                        for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']:
                            try:
                                parsed_date = datetime.strptime(created_date.split('T')[0], fmt)
                                df.loc[idx, 'Year Created'] = parsed_date.year
                                df.loc[idx, 'Month Created'] = parsed_date.strftime('%B')
                                df.loc[idx, 'Quarter Created'] = f"Q{((parsed_date.month - 1) // 3) + 1}"
                                df.loc[idx, 'Week Number'] = parsed_date.isocalendar()[1]
                                df.loc[idx, 'Day of Week Created'] = parsed_date.strftime('%A')
                                break
                            except:
                                continue
                except:
                    pass
        
        logger.info(f"✅ Processed {len(df)} manually uploaded records with {len(df.columns)} columns")
        return df
    
    def generate_analysis(self, df):
        """Generate analysis and insights"""
        analysis = {}
        
        # Basic statistics
        analysis['total_cases'] = len(df)
        analysis['open_cases'] = len(df[~df['Status'].str.contains('Closed|Resolved', case=False, na=False)])
        analysis['closed_cases'] = len(df[df['Status'].str.contains('Closed|Resolved', case=False, na=False)])
        analysis['resolution_rate'] = (analysis['closed_cases'] / analysis['total_cases'] * 100) if analysis['total_cases'] > 0 else 0
        
        # Status distribution
        analysis['status_distribution'] = df['Status'].value_counts().to_dict()
        
        # Priority distribution
        analysis['priority_distribution'] = df['Priority'].value_counts().to_dict()
        
        # Time-based analysis
        if 'Year' in df.columns:
            analysis['yearly_trend'] = df.groupby('Year').size().to_dict()
        
        if 'Month' in df.columns:
            analysis['monthly_pattern'] = df['Month'].value_counts().to_dict()
        
        if 'Day of Week' in df.columns:
            analysis['day_pattern'] = df['Day of Week'].value_counts().to_dict()
        
        # Top assignees
        analysis['top_assignees'] = df['Assignee'].value_counts().head(10).to_dict()
        
        # Recent activity (last 7 days)
        try:
            df['Created_dt'] = pd.to_datetime(df['Created'], errors='coerce')
            seven_days_ago = datetime.now() - timedelta(days=7)
            recent_cases = df[df['Created_dt'] > seven_days_ago]
            analysis['recent_cases_7d'] = len(recent_cases)
        except:
            analysis['recent_cases_7d'] = 0
        
        logger.info("✅ Analysis completed")
        return analysis
    
    def update_google_sheets(self, df, analysis):
        """Update Google Sheets with data and analysis"""
        try:
            spreadsheet = self.sheets_client.open_by_key(self.sheet_id)
            
            # Create/update worksheets
            worksheet_names = [ws.title for ws in spreadsheet.worksheets()]
            
            # 1. Raw Data Sheet
            if 'Raw Data' not in worksheet_names:
                raw_sheet = spreadsheet.add_worksheet(title='Raw Data', rows=1000, cols=20)
            else:
                raw_sheet = spreadsheet.worksheet('Raw Data')
            
            raw_sheet.clear()
            headers = df.columns.tolist()
            # Convert all data to strings to avoid JSON serialization issues
            df_str = df.astype(str).fillna('')
            data_rows = df_str.values.tolist()
            all_data = [headers] + data_rows
            raw_sheet.update(values=all_data, range_name='A1')
            logger.info(f"✅ Updated 'Raw Data' sheet with {len(data_rows)} rows")
            
            # 2. Analysis Sheet
            if 'Analysis' not in worksheet_names:
                analysis_sheet = spreadsheet.add_worksheet(title='Analysis', rows=100, cols=10)
            else:
                analysis_sheet = spreadsheet.worksheet('Analysis')
            
            analysis_sheet.clear()
            analysis_rows = [
                ['Metric', 'Value', 'Last Updated'],
                ['Total Cases', analysis['total_cases'], datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                ['Open Cases', analysis['open_cases'], ''],
                ['Closed Cases', analysis['closed_cases'], ''],
                ['Resolution Rate', f"{analysis['resolution_rate']:.1f}%", ''],
                ['Recent Cases (7 days)', analysis['recent_cases_7d'], ''],
                ['', '', ''],
                ['Status Distribution', '', '']
            ]
            
            for status, count in analysis['status_distribution'].items():
                analysis_rows.append([status, count, ''])
            
            analysis_rows.append(['', '', ''])
            analysis_rows.append(['Priority Distribution', '', ''])
            
            for priority, count in analysis['priority_distribution'].items():
                analysis_rows.append([priority, count, ''])
            
            analysis_sheet.update('A1', analysis_rows)
            logger.info("✅ Updated 'Analysis' sheet")
            
            # 3. Dashboard Sheet (first sheet)
            dashboard = spreadsheet.sheet1
            dashboard.clear()
            dashboard_data = [
                ['Historical CS Analysis - Applications (IA)', '', ''],
                ['Last Updated:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'), ''],
                ['', '', ''],
                ['Key Metrics', '', ''],
                ['Total Cases:', analysis['total_cases'], ''],
                ['Open:', analysis['open_cases'], f"{(analysis['open_cases']/analysis['total_cases']*100):.1f}%"],
                ['Closed:', analysis['closed_cases'], f"{(analysis['closed_cases']/analysis['total_cases']*100):.1f}%"],
                ['Resolution Rate:', f"{analysis['resolution_rate']:.1f}%", ''],
                ['', '', ''],
                ['Recent Activity', '', ''],
                ['Last 7 days:', analysis['recent_cases_7d'], 'new cases'],
                ['', '', ''],
                ['View Details:', 'See Raw Data sheet', ''],
                ['View Analysis:', 'See Analysis sheet', '']
            ]
            
            dashboard.update('A1', dashboard_data)
            logger.info("✅ Updated 'Dashboard' sheet")
            
            logger.info(f"✅ Google Sheets fully updated!")
            logger.info(f"🔗 View at: https://docs.google.com/spreadsheets/d/{self.sheet_id}")
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Failed to update Google Sheets: {e}")
            return False
    
    def run(self):
        """Execute the complete automation workflow"""
        logger.info("🚀 Starting Google Sheets Analysis of Manually Uploaded JIRA Data")
        logger.info("=" * 50)
        
        # Connect to Google Sheets (no need for JIRA connection for manual data)
        if not self.connect_google_sheets():
            return False
        
        # Read manually uploaded data from sheets
        if not self.fetch_manual_data_from_sheets():
            return False
        
        df = self.process_data()
        if df is None:
            return False
        
        # Generate analysis
        analysis = self.generate_analysis(df)
        
        # Update Google Sheets
        if not self.update_google_sheets(df, analysis):
            return False
        
        logger.info("✅ Automation completed successfully!")
        return True

def main():
    """Main execution function"""
    automation = JiraGoogleSheetsAutomation()
    success = automation.run()
    
    if success:
        print("\n🎉 SUCCESS! Your Google Sheets analysis has been completed!")
        print("   • Analysis sheet - Statistics and distributions from your uploaded data")
        print("   • Dashboard sheet - Key metrics and insights")
        print(f"\n🔗 View your sheets: https://docs.google.com/spreadsheets/d/1HG_FFiGu5XoPmxIhQxRSjYO2N77qlFtV4Pj5bRiBqco")
    else:
        print("\n❌ Analysis failed. Check the logs above for details.")
        print("\n📋 Common issues:")
        print("   1. Missing service_account_key.json")
        print("   2. Sheet not shared with service account email")
        print("   3. No data found in 'Raw Data' sheet - please upload your JIRA data manually first")
    
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())
