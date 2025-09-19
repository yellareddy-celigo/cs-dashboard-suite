#!/usr/bin/env python3
"""
Scheduler for Automated JIRA to Google Sheets Updates
Runs the automation at specified intervals
"""

import schedule
import time
import subprocess
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def run_automation():
    """Run the JIRA to Google Sheets automation"""
    logger.info("🔄 Starting scheduled automation run...")
    try:
        result = subprocess.run(['python3', 'jira_sheets_automation.py'], 
                              capture_output=True, text=True)
        
        if result.returncode == 0:
            logger.info("✅ Automation completed successfully")
        else:
            logger.error(f"❌ Automation failed: {result.stderr}")
    
    except Exception as e:
        logger.error(f"❌ Error running automation: {e}")

def main():
    logger.info("🕐 Scheduler Started - JIRA to Google Sheets Automation")
    logger.info("=" * 50)
    
    # Schedule options (uncomment the one you want to use):
    
    # Option 1: Run every hour
    schedule.every().hour.do(run_automation)
    
    # Option 2: Run every day at specific time
    # schedule.every().day.at("09:00").do(run_automation)
    
    # Option 3: Run every 30 minutes
    # schedule.every(30).minutes.do(run_automation)
    
    # Option 4: Run every Monday at 9 AM
    # schedule.every().monday.at("09:00").do(run_automation)
    
    logger.info("📅 Schedule configured: Running every hour")
    logger.info("🔗 Target Sheet: https://docs.google.com/spreadsheets/d/1HG_FFiGu5XoPmxIhQxRSjYO2N77qlFtV4Pj5bRiBqco")
    logger.info("Press Ctrl+C to stop the scheduler")
    
    # Run once immediately
    run_automation()
    
    # Keep running the schedule
    while True:
        schedule.run_pending()
        time.sleep(60)  # Check every minute

if __name__ == "__main__":
    main()
