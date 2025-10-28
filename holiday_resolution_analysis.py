#!/usr/bin/env python3
"""
HOLIDAY CS CASES RESOLUTION COMMENTS ANALYSIS
Read Resolution Comments for each CS case and provide specific recommendations

Usage:
    python3 holiday_resolution_analysis.py
"""

import pandas as pd
import numpy as np
import re
import argparse
import sys
from datetime import datetime
from collections import Counter

VERSION = "1.0.0"
LAST_UPDATED = "2025-01-09"

def analyze_holiday_resolution_comments(csv_file, output_file=None):
    """Analyze Resolution Comments for each holiday CS case and provide recommendations"""
    
    print("="*100)
    print(f"HOLIDAY CS CASES RESOLUTION COMMENTS ANALYSIS")
    print(f"Version: {VERSION} | Last Updated: {LAST_UPDATED}")
    print("="*100)
    
    # Load CSV data
    df = pd.read_csv(csv_file)
    print(f"\n✅ Loaded {len(df)} holiday cases from {csv_file}")
    
    # Check if Resolution Comments column exists
    resolution_comments_col = 'Custom field (Resolution Comments)'
    if resolution_comments_col not in df.columns:
        print(f"❌ Column '{resolution_comments_col}' not found in CSV")
        return None
    
    # Process each case
    analyzed_cases = []
    recommendations = []
    
    for idx, case in df.iterrows():
        case_key = case.get('Issue key', '')
        summary = case.get('Summary', '')
        resolution = case.get('Resolution', '')
        status = case.get('Status', '')
        priority = case.get('Priority', '')
        integration = case.get('Custom field (Integration Apps)', '')
        case_type = case.get('Custom field (Case Type)', '')
        created = case.get('Created', '')
        description = case.get('Description', '')
        
        # Get Resolution Comments
        resolution_comments = case.get(resolution_comments_col, '')
        
        # Analyze this specific case
        case_analysis = analyze_case_resolution(case_key, summary, description, resolution_comments, integration, case_type, priority)
        
        # Extract issue and fix from resolution comments
        extracted_issue, extracted_fix = extract_issue_and_fix_from_comments(resolution_comments)
        
        # Create case record
        case_record = {
            'Case Key': case_key,
            'Summary': summary,
            'Priority': priority,
            'Status': status,
            'Resolution': resolution,
            'Case Type': case_type,
            'Integration': integration,
            'Created': created,
            'Description': description[:300] + '...' if len(str(description)) > 300 else description,
            'Resolution Comments': resolution_comments,
            'Issue Identified': extracted_issue if extracted_issue else case_analysis['issue_identified'],
            'How It Was Fixed': extracted_fix if extracted_fix else case_analysis['resolution_method'],
            'Root Cause': case_analysis['root_cause'],
            'Resolution Method': case_analysis['resolution_method'],
            'Customer Impact': case_analysis['customer_impact'],
            'Recurrence Risk': case_analysis['recurrence_risk'],
            'Specific Recommendations': case_analysis['recommendations'],
            'Preventive Actions': case_analysis['preventive_actions'],
            'Holiday Season Risk': case_analysis['holiday_risk']
        }
        
        analyzed_cases.append(case_record)
        
        # Collect recommendations
        if case_analysis['recommendations']:
            recommendations.extend(case_analysis['recommendations'].split('; '))
    
    # Create DataFrame
    analyzed_df = pd.DataFrame(analyzed_cases)
    
    # Generate summary recommendations
    summary_recommendations = generate_summary_recommendations(analyzed_df, recommendations)
    
    # Generate summary statistics
    total_cases = len(analyzed_df)
    cases_with_comments = len(analyzed_df[analyzed_df['Resolution Comments'].notna()])
    high_risk_cases = len(analyzed_df[analyzed_df['Recurrence Risk'] == 'High'])
    high_holiday_risk = len(analyzed_df[analyzed_df['Holiday Season Risk'] == 'High'])
    
    print(f"\n📊 ANALYSIS SUMMARY:")
    print(f"  Total Cases Analyzed: {total_cases}")
    print(f"  Cases with Resolution Comments: {cases_with_comments}")
    print(f"  High Recurrence Risk Cases: {high_risk_cases}")
    print(f"  High Holiday Season Risk Cases: {high_holiday_risk}")
    print(f"  Total Recommendations Generated: {len(set(recommendations))}")
    
    # Create output file
    if output_file is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f'Holiday_Resolution_Analysis_{timestamp}.xlsx'
    
    print(f"\n📝 Creating resolution analysis: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Main analysis sheet
        analyzed_df.to_excel(writer, sheet_name='Resolution Comments Analysis', index=False)
        
        # Summary recommendations
        summary_df = pd.DataFrame(summary_recommendations)
        summary_df.to_excel(writer, sheet_name='Summary Recommendations', index=False)
        
        # High risk cases
        high_risk_cases = analyzed_df[analyzed_df['Recurrence Risk'] == 'High']
        high_risk_cases.to_excel(writer, sheet_name='High Risk Cases', index=False)
        
        # High holiday risk cases
        high_holiday_risk_cases = analyzed_df[analyzed_df['Holiday Season Risk'] == 'High']
        high_holiday_risk_cases.to_excel(writer, sheet_name='High Holiday Risk Cases', index=False)
        
        # Root cause analysis
        root_cause_analysis = analyzed_df['Root Cause'].value_counts().reset_index()
        root_cause_analysis.columns = ['Root Cause', 'Count']
        root_cause_analysis['Percentage'] = (root_cause_analysis['Count'] / total_cases * 100).round(1).astype(str) + '%'
        root_cause_analysis.to_excel(writer, sheet_name='Root Cause Analysis', index=False)
        
        # Integration analysis
        integration_analysis = analyzed_df['Integration'].value_counts().reset_index()
        integration_analysis.columns = ['Integration', 'Count']
        integration_analysis['Percentage'] = (integration_analysis['Count'] / total_cases * 100).round(1).astype(str) + '%'
        integration_analysis.to_excel(writer, sheet_name='Integration Analysis', index=False)
        
        # Resolution method analysis
        resolution_analysis = analyzed_df['Resolution Method'].value_counts().reset_index()
        resolution_analysis.columns = ['Resolution Method', 'Count']
        resolution_analysis['Percentage'] = (resolution_analysis['Count'] / total_cases * 100).round(1).astype(str) + '%'
        resolution_analysis.to_excel(writer, sheet_name='Resolution Method Analysis', index=False)
    
    print(f"\n✅ Holiday resolution analysis complete!")
    print(f"📁 Output file: {output_file}")
    print(f"📋 Contains 7 sheets:")
    print(f"  1. Resolution Comments Analysis - All {total_cases} cases with detailed analysis")
    print(f"  2. Summary Recommendations - {len(summary_recommendations)} key recommendations")
    print(f"  3. High Risk Cases - {high_risk_cases} high recurrence risk cases")
    print(f"  4. High Holiday Risk Cases - {high_holiday_risk} high holiday risk cases")
    print(f"  5. Root Cause Analysis - Root cause breakdown")
    print(f"  6. Integration Analysis - Integration breakdown")
    print(f"  7. Resolution Method Analysis - Resolution method breakdown")
    
    return output_file

def analyze_case_resolution(case_key, summary, description, resolution_comments, integration, case_type, priority):
    """Analyze resolution comments for a specific case"""
    
    # Combine all text for analysis
    combined_text = f"{summary} {description} {resolution_comments}".lower()
    
    # Identify specific issue - Extract from resolution comments first as it's most reliable
    issue_identified = identify_specific_issue(summary, description, resolution_comments)
    
    # Extract exact issue and fix from resolution comments
    extracted_issue, extracted_fix = extract_issue_and_fix_from_comments(resolution_comments)
    
    # Use extracted details if available
    if extracted_issue:
        issue_identified = extracted_issue
    if extracted_fix:
        fix_applied = extracted_fix
    else:
        fix_applied = determine_resolution_method(resolution_comments)
    
    # Determine root cause
    root_cause = determine_root_cause(combined_text, resolution_comments)
    
    # Determine resolution method
    resolution_method = fix_applied if 'fix_applied' in locals() else determine_resolution_method(resolution_comments)
    
    # Assess customer impact
    customer_impact = assess_customer_impact(combined_text, resolution_comments)
    
    # Assess recurrence risk
    recurrence_risk = assess_recurrence_risk(combined_text, resolution_comments, root_cause)
    
    # Generate specific recommendations
    recommendations = generate_specific_recommendations(case_key, issue_identified, root_cause, integration, resolution_method, resolution_comments)
    
    # Generate preventive actions
    preventive_actions = generate_preventive_actions(root_cause, integration, resolution_method)
    
    # Assess holiday season risk
    holiday_risk = assess_holiday_risk(customer_impact, recurrence_risk, root_cause)
    
    return {
        'issue_identified': issue_identified,
        'root_cause': root_cause,
        'resolution_method': resolution_method,
        'customer_impact': customer_impact,
        'recurrence_risk': recurrence_risk,
        'recommendations': recommendations,
        'preventive_actions': preventive_actions,
        'holiday_risk': holiday_risk
    }

def extract_issue_and_fix_from_comments(resolution_comments):
    """Extract the actual issue and fix directly from resolution comments"""
    
    if pd.isna(resolution_comments) or str(resolution_comments).strip() in ['', 'nan']:
        return None, None
    
    comments_text = str(resolution_comments)
    comments_lower = comments_text.lower()
    
    # Extract issue patterns
    issue = None
    
    # Pattern 1: "Issue: ..." or "Problem: ..."
    issue_match = re.search(r'(?:issue|problem):\s*([^\.]+(?:\.[^\.]+)*)', comments_text, re.IGNORECASE)
    if issue_match:
        issue = issue_match.group(1).strip()
    
    # Pattern 2: "Customer reported ..."
    elif re.search(r'customer (?:reported|was|had)', comments_lower):
        issue_match = re.search(r'customer (?:reported|was|had)\s+([^\.]+)', comments_text, re.IGNORECASE)
        if issue_match:
            issue = issue_match.group(1).strip()
    
    # Pattern 3: "Error occurred ..." or "Error: ..."
    elif re.search(r'error', comments_lower):
        error_match = re.search(r'error[:\s]+([^\.]+)', comments_text, re.IGNORECASE)
        if error_match:
            issue = f"Error: {error_match.group(1).strip()}"
    
    # Pattern 4: "Not working" or "Failed" patterns
    elif re.search(r'(?:not working|failed|failing|not able)', comments_lower):
        # Extract the full sentence with the failure
        failure_match = re.search(r'([^\.]*(?:not working|failed|failing|not able)[^\.]*)', comments_text, re.IGNORECASE)
        if failure_match:
            issue = failure_match.group(1).strip()
    
    # Extract fix patterns
    fix = None
    
    # Pattern 1: "Fixed by ..." or "Resolved by ..."
    fix_match = re.search(r'(?:fixed|resolved)\s+by\s+([^\.]+(?:\.[^\.]+)*)', comments_text, re.IGNORECASE)
    if fix_match:
        fix = fix_match.group(1).strip()
    
    # Pattern 2: "Solution: ..."
    elif re.search(r'solution:', comments_lower):
        solution_match = re.search(r'solution:\s*([^\.]+(?:\.[^\.]+)*)', comments_text, re.IGNORECASE)
        if solution_match:
            fix = solution_match.group(1).strip()
    
    # Pattern 3: "Action taken: ..."
    elif re.search(r'action taken:', comments_lower):
        action_match = re.search(r'action taken:\s*([^\.]+(?:\.[^\.]+)*)', comments_text, re.IGNORECASE)
        if action_match:
            fix = action_match.group(1).strip()
    
    # Pattern 4: "Changed ..." or "Updated ..." or "Modified ..."
    elif re.search(r'(?:changed|updated|modified|configured|reconfigured)', comments_lower):
        action_match = re.search(r'(?:changed|updated|modified|configured|reconfigured)\s+([^\.]+)', comments_text, re.IGNORECASE)
        if action_match:
            fix = f"Changed: {action_match.group(1).strip()}"
    
    # Pattern 5: "Customer advised ..." or "Customer informed ..."
    elif re.search(r'customer (?:advised|informed|told|guided)', comments_lower):
        advice_match = re.search(r'customer (?:advised|informed|told|guided)\s+to\s+([^\.]+)', comments_text, re.IGNORECASE)
        if advice_match:
            fix = f"Customer advised: {advice_match.group(1).strip()}"
    
    # Pattern 6: Workaround mentioned
    elif re.search(r'workaround|temporary fix|interim', comments_lower):
        fix = "Workaround applied (see resolution comments for details)"
    
    # If no specific pattern found, extract first 2-3 sentences as context
    if not issue:
        sentences = re.split(r'[.!?]+', comments_text)
        if len(sentences) > 0:
            issue = sentences[0].strip()
    
    if not fix and len(sentences) > 1:
        fix_parts = [s.strip() for s in sentences[1:3] if s.strip()]
        if fix_parts:
            fix = '. '.join(fix_parts)
    
    return issue, fix

def identify_specific_issue(summary, description, resolution_comments):
    """Identify the specific issue from the case details"""
    
    # Start with summary as primary identifier
    issue = summary
    
    # Look for specific patterns in resolution comments
    if pd.notna(resolution_comments):
        comments_text = str(resolution_comments).lower()
        
        # Extract specific error messages
        if 'error:' in comments_text:
            error_match = re.search(r'error:\s*([^\.]+)', comments_text)
            if error_match:
                issue = f"{summary} - {error_match.group(1).strip()}"
        
        # Extract specific problems
        elif 'issue:' in comments_text:
            issue_match = re.search(r'issue:\s*([^\.]+)', comments_text)
            if issue_match:
                issue = f"{summary} - {issue_match.group(1).strip()}"
        
        # Extract specific failures
        elif 'failed' in comments_text or 'failure' in comments_text:
            failure_match = re.search(r'(failed|failure)[^\.]*', comments_text)
            if failure_match:
                issue = f"{summary} - {failure_match.group(0).strip()}"
    
    return issue

def determine_root_cause(combined_text, resolution_comments):
    """Determine the root cause of the issue"""
    
    # Holiday-specific causes
    if any(word in combined_text for word in ['holiday', 'peak', 'high volume', 'increased load', 'seasonal']):
        return 'Holiday Season Volume'
    
    # Configuration issues
    elif any(word in combined_text for word in ['configuration', 'setup', 'config', 'not configured', 'misconfigured']):
        return 'Configuration Error'
    
    # API-related causes
    elif any(word in combined_text for word in ['api', 'rate limit', 'quota', 'endpoint', 'request failed']):
        return 'API Limitations'
    
    # Authentication issues
    elif any(word in combined_text for word in ['authentication', 'auth', 'token', 'credential', 'unauthorized', '401', '403']):
        return 'Authentication Failure'
    
    # Data mapping issues
    elif any(word in combined_text for word in ['mapping', 'field', 'invalid field', 'missing field', 'field mapping']):
        return 'Data Mapping Issue'
    
    # Sync issues
    elif any(word in combined_text for word in ['sync', 'synchronization', 'not syncing', 'sync error', 'sync failed']):
        return 'Data Synchronization Problem'
    
    # Performance issues
    elif any(word in combined_text for word in ['performance', 'slow', 'timeout', 'delay', 'lag', 'bottleneck']):
        return 'Performance Issue'
    
    # Data validation issues
    elif any(word in combined_text for word in ['validation', 'invalid', 'required', 'format', 'data format']):
        return 'Data Validation Error'
    
    # Duplicate data issues
    elif any(word in combined_text for word in ['duplicate', 'duplication', 'duplicated', 'already exists']):
        return 'Duplicate Data Issue'
    
    # Connection issues
    elif any(word in combined_text for word in ['connection', 'connectivity', 'network', 'disconnect', 'connection failed']):
        return 'Connection Problem'
    
    # Code/script issues
    elif any(word in combined_text for word in ['script', 'code', 'bug', 'error', 'exception', 'crash']):
        return 'Code/Script Error'
    
    # External system issues
    elif any(word in combined_text for word in ['external', 'third party', 'vendor', 'partner', 'system down']):
        return 'External System Issue'
    
    # Check resolution comments for more specific causes
    if pd.notna(resolution_comments):
        comments_text = str(resolution_comments).lower()
        
        if 'customer' in comments_text and ('advised' in comments_text or 'informed' in comments_text):
            return 'Customer Education Needed'
        
        elif 'workaround' in comments_text or 'temporary' in comments_text:
            return 'Requires Workaround'
        
        elif 'escalated' in comments_text or 'dev team' in comments_text:
            return 'Engineering Issue'
    
    return 'Unknown/Other'

def determine_resolution_method(resolution_comments):
    """Determine how the issue was resolved"""
    
    if pd.isna(resolution_comments) or str(resolution_comments).strip() in ['', 'nan']:
        return 'No Resolution Comments'
    
    comments_text = str(resolution_comments).lower()
    
    # Code fixes
    if any(word in comments_text for word in ['fixed', 'resolved', 'implemented', 'deployed', 'code fix', 'bug fix']):
        return 'Code Fix'
    
    # Workarounds
    elif any(word in comments_text for word in ['workaround', 'work-around', 'temporary', 'interim', 'manual']):
        return 'Workaround Applied'
    
    # Customer guidance
    elif any(word in comments_text for word in ['customer advised', 'customer informed', 'customer told', 'guided', 'instructed']):
        return 'Customer Guidance'
    
    # Configuration changes
    elif any(word in comments_text for word in ['configuration', 'setup', 'reconfigured', 'reauthorized', 'settings']):
        return 'Configuration Change'
    
    # Escalation
    elif any(word in comments_text for word in ['escalated', 'escalation', 'dev team', 'engineering', 'product team']):
        return 'Escalated to Engineering'
    
    # Data fixes
    elif any(word in comments_text for word in ['data', 'record', 'deleted', 'updated', 'corrected']):
        return 'Data Fix'
    
    # External resolution
    elif any(word in comments_text for word in ['external', 'vendor', 'partner', 'third party']):
        return 'External Resolution'
    
    # No action needed
    elif any(word in comments_text for word in ['no action', 'not needed', 'by design', 'expected behavior']):
        return 'No Action Required'
    
    return 'Other/Unknown'

def assess_customer_impact(combined_text, resolution_comments):
    """Assess the impact on the customer"""
    
    # High impact indicators
    if any(word in combined_text for word in ['critical', 'urgent', 'blocking', 'stopped', 'down', 'broken', 'not working']):
        return 'High'
    
    # Medium impact indicators
    elif any(word in combined_text for word in ['important', 'affecting', 'impacting', 'delayed', 'slow', 'issue']):
        return 'Medium'
    
    # Check resolution comments for impact indicators
    if pd.notna(resolution_comments):
        comments_text = str(resolution_comments).lower()
        
        if any(word in comments_text for word in ['customer', 'user', 'client']) and any(word in comments_text for word in ['blocked', 'stopped', 'cannot', 'unable']):
            return 'High'
        
        elif any(word in comments_text for word in ['customer', 'user', 'client']) and any(word in comments_text for word in ['delayed', 'slow', 'issue']):
            return 'Medium'
    
    return 'Low'

def assess_recurrence_risk(combined_text, resolution_comments, root_cause):
    """Assess the risk of this issue recurring"""
    
    # High recurrence risk indicators
    if any(word in combined_text for word in ['recurring', 'repeated', 'happening again', 'same issue', 'similar problem']):
        return 'High'
    
    # Workaround indicates high recurrence risk
    if pd.notna(resolution_comments):
        comments_text = str(resolution_comments).lower()
        if any(word in comments_text for word in ['workaround', 'temporary', 'interim', 'manual']):
            return 'High'
    
    # Root cause based assessment
    if root_cause in ['Configuration Error', 'Data Mapping Issue', 'Authentication Failure']:
        return 'High'
    
    elif root_cause in ['API Limitations', 'Performance Issue', 'Data Validation Error']:
        return 'Medium'
    
    elif root_cause in ['Code/Script Error', 'External System Issue']:
        return 'Low'
    
    return 'Medium'

def generate_specific_recommendations(case_key, issue_identified, root_cause, integration, resolution_method, resolution_comments):
    """Generate specific recommendations for this case"""
    
    recommendations = []
    
    # Root cause specific recommendations
    if root_cause == 'Configuration Error':
        recommendations.extend([
            'Implement configuration validation tools',
            'Add configuration testing and preview',
            'Create configuration templates and best practices',
            'Implement configuration health checks'
        ])
    
    elif root_cause == 'Data Mapping Issue':
        recommendations.extend([
            'Implement field mapping validation',
            'Add mapping preview and testing',
            'Create mapping templates',
            'Implement duplicate detection'
        ])
    
    elif root_cause == 'Authentication Failure':
        recommendations.extend([
            'Implement automated token refresh',
            'Add credential validation',
            'Create authentication monitoring',
            'Implement token expiration warnings'
        ])
    
    elif root_cause == 'API Limitations':
        recommendations.extend([
            'Implement API rate limit monitoring',
            'Add API health checks',
            'Create API quota management',
            'Implement API retry logic'
        ])
    
    elif root_cause == 'Data Synchronization Problem':
        recommendations.extend([
            'Implement sync monitoring',
            'Add automated retry mechanisms',
            'Create sync health checks',
            'Implement sync queue management'
        ])
    
    # Resolution method specific recommendations
    if resolution_method == 'Workaround Applied':
        recommendations.extend([
            'Plan permanent fix to replace workaround',
            'Document workaround limitations',
            'Create monitoring for workaround effectiveness',
            'Implement automated permanent fix deployment'
        ])
    
    elif resolution_method == 'Customer Guidance':
        recommendations.extend([
            'Create self-service documentation',
            'Implement guided setup wizards',
            'Add validation and error prevention',
            'Create customer education materials'
        ])
    
    # Integration specific recommendations
    if integration:
        if 'NetSuite' in integration:
            recommendations.extend(['Add NetSuite-specific monitoring', 'Implement NetSuite health checks'])
        if 'Amazon' in integration:
            recommendations.extend(['Add Amazon API monitoring', 'Implement Amazon quota management'])
        if 'Salesforce' in integration:
            recommendations.extend(['Add Salesforce monitoring', 'Implement Salesforce health checks'])
        if 'Shopify' in integration:
            recommendations.extend(['Add Shopify API monitoring', 'Implement Shopify rate limit management'])
    
    # Holiday specific recommendations
    recommendations.extend([
        'Implement holiday season monitoring',
        'Add holiday season capacity planning',
        'Create holiday season runbooks',
        'Implement holiday season alerting'
    ])
    
    # Remove duplicates and limit
    unique_recommendations = list(dict.fromkeys(recommendations))
    return '; '.join(unique_recommendations[:8])

def generate_preventive_actions(root_cause, integration, resolution_method):
    """Generate preventive actions for this case"""
    
    actions = []
    
    # Root cause specific actions
    if root_cause == 'Configuration Error':
        actions.extend([
            'Add configuration validation before deployment',
            'Create configuration templates for common setups',
            'Implement configuration health checks',
            'Add configuration testing environment'
        ])
    
    elif root_cause == 'Data Mapping Issue':
        actions.extend([
            'Implement field mapping validation tools',
            'Add mapping preview and testing',
            'Create mapping templates',
            'Add duplicate detection and prevention'
        ])
    
    elif root_cause == 'Authentication Failure':
        actions.extend([
            'Implement automated token refresh',
            'Add credential validation and health checks',
            'Create authentication monitoring dashboard',
            'Add token expiration warnings and alerts'
        ])
    
    elif root_cause == 'API Limitations':
        actions.extend([
            'Implement API rate limit monitoring',
            'Add API health checks and circuit breakers',
            'Create API quota management system',
            'Add API retry logic with exponential backoff'
        ])
    
    elif root_cause == 'Data Synchronization Problem':
        actions.extend([
            'Implement sync monitoring dashboard',
            'Add automated retry mechanisms',
            'Create sync health checks',
            'Implement sync queue management'
        ])
    
    # Integration specific actions
    if integration:
        if 'NetSuite' in integration:
            actions.append('Add NetSuite-specific monitoring and health checks')
        if 'Amazon' in integration:
            actions.append('Add Amazon API monitoring and quota management')
        if 'Salesforce' in integration:
            actions.append('Add Salesforce monitoring and health checks')
        if 'Shopify' in integration:
            actions.append('Add Shopify API monitoring and rate limit management')
    
    # Holiday specific actions
    actions.extend([
        'Implement holiday season monitoring dashboard',
        'Add holiday season capacity planning',
        'Create holiday season runbooks and procedures',
        'Implement holiday season alerting system'
    ])
    
    # Remove duplicates and limit
    unique_actions = list(dict.fromkeys(actions))
    return '; '.join(unique_actions[:6])

def assess_holiday_risk(customer_impact, recurrence_risk, root_cause):
    """Assess the risk during holiday season"""
    
    # High holiday risk
    if customer_impact == 'High' and recurrence_risk == 'High':
        return 'High'
    
    # Medium holiday risk
    elif customer_impact == 'High' or recurrence_risk == 'High':
        return 'Medium'
    
    # Check for holiday-specific root causes
    if root_cause in ['Holiday Season Volume', 'Performance Issue', 'API Limitations']:
        return 'High'
    
    return 'Low'

def generate_summary_recommendations(analyzed_df, recommendations):
    """Generate summary recommendations based on all cases"""
    
    # Count recommendation frequency
    rec_counter = Counter(recommendations)
    
    # Get top recommendations
    top_recommendations = rec_counter.most_common(20)
    
    summary_recs = []
    for rec, count in top_recommendations:
        summary_recs.append({
            'Recommendation': rec,
            'Frequency': count,
            'Priority': 'High' if count > 10 else 'Medium' if count > 5 else 'Low',
            'Impact': 'High' if 'monitoring' in rec.lower() or 'validation' in rec.lower() else 'Medium',
            'Effort': 'High' if 'implement' in rec.lower() or 'create' in rec.lower() else 'Low'
        })
    
    return summary_recs

# ============================================================================
# COMMAND LINE INTERFACE
# ============================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Holiday CS Cases Resolution Comments Analysis',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python3 holiday_resolution_analysis.py --file Holiday.csv
  python3 holiday_resolution_analysis.py --file /path/to/holiday.csv --output Analysis.xlsx
  
This tool provides:
  - Analysis of resolution comments for each holiday CS case
  - Specific recommendations for each case
  - Root cause analysis and preventive actions
  - Holiday season risk assessment
        ''',
        add_help=True
    )
    
    parser.add_argument('--file', '-f', default='~/Downloads/Holiday.csv',
                       help='Path to the Holiday CSV file')
    parser.add_argument('--output', '-o',
                       help='Output Excel file name (optional)')
    
    args = parser.parse_args()
    
    try:
        output_file = analyze_holiday_resolution_comments(args.file, args.output)
        if output_file:
            print(f"\n🎯 Holiday resolution analysis complete: {output_file}")
    except Exception as e:
        print(f"\n❌ Error during analysis: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

