#!/usr/bin/env python3
"""
SIMPLIFIED HOLIDAY CASE ANALYSIS
Analyzes each case resolution comment individually - Only Individual Case Analysis sheet

Usage:
    python3 simplified_holiday_analysis.py
"""

import pandas as pd
import numpy as np
import re
import argparse
from datetime import datetime

VERSION = "1.0.0"
LAST_UPDATED = "2025-01-09"

def analyze_individual_cases_only(csv_file, output_file=None):
    """Analyze each case resolution comment individually - simplified version"""
    
    print("="*100)
    print(f"SIMPLIFIED HOLIDAY CASE ANALYSIS")
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
    
    # Process each case individually
    detailed_cases = []
    
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
        
        # Detailed analysis of this specific case
        case_analysis = analyze_individual_case(case_key, summary, description, resolution_comments, integration, case_type, priority)
        
        # Create detailed case record
        case_record = {
            'Case Key': case_key,
            'Summary': summary,
            'Priority': priority,
            'Status': status,
            'Resolution': resolution,
            'Case Type': case_type,
            'Integration': integration,
            'Created': created,
            'Description': description[:500] + '...' if len(str(description)) > 500 else description,
            'Resolution Comments': resolution_comments,
            'Issue Identified': case_analysis['issue_identified'],
            'Root Cause': case_analysis['root_cause'],
            'Resolution Method': case_analysis['resolution_method'],
            'Technical Details': case_analysis['technical_details'],
            'Customer Impact': case_analysis['customer_impact'],
            'Recurrence Risk': case_analysis['recurrence_risk'],
            'Preventive Actions': case_analysis['preventive_actions'],
            'Holiday Season Impact': case_analysis['holiday_impact'],
            'Urgency Level': case_analysis['urgency_level'],
            'Will This Happen Again?': case_analysis['will_happen_again'],
            'How to Prevent': case_analysis['how_to_prevent']
        }
        
        detailed_cases.append(case_record)
    
    # Create DataFrame
    detailed_df = pd.DataFrame(detailed_cases)
    
    # Generate summary statistics
    total_cases = len(detailed_df)
    high_recurrence_cases = len(detailed_df[detailed_df['Recurrence Risk'] == 'High'])
    high_impact_cases = len(detailed_df[detailed_df['Holiday Season Impact'] == 'High'])
    
    print(f"\n📊 ANALYSIS SUMMARY:")
    print(f"  Total Cases Analyzed: {total_cases}")
    print(f"  High Recurrence Risk Cases: {high_recurrence_cases}")
    print(f"  High Holiday Impact Cases: {high_impact_cases}")
    
    # Create output file
    if output_file is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f'Individual_Case_Analysis_{timestamp}.xlsx'
    
    print(f"\n📝 Creating individual case analysis: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Only Individual Case Analysis sheet
        detailed_df.to_excel(writer, sheet_name='Individual Case Analysis', index=False)
    
    print(f"\n✅ Individual case analysis complete!")
    print(f"📁 Output file: {output_file}")
    print(f"📋 Contains 1 sheet:")
    print(f"  Individual Case Analysis - All {total_cases} cases with detailed analysis")
    
    return output_file

def analyze_individual_case(case_key, summary, description, resolution_comments, integration, case_type, priority):
    """Analyze an individual case in detail"""
    
    # Combine all text for analysis
    combined_text = f"{summary} {description} {resolution_comments}".lower()
    
    # Identify specific issue
    issue_identified = identify_specific_issue(summary, description, resolution_comments)
    
    # Determine root cause
    root_cause = determine_root_cause(combined_text, resolution_comments)
    
    # Determine resolution method
    resolution_method = determine_resolution_method(resolution_comments)
    
    # Extract technical details
    technical_details = extract_technical_details(resolution_comments)
    
    # Assess customer impact
    customer_impact = assess_customer_impact(combined_text, resolution_comments)
    
    # Assess recurrence risk
    recurrence_risk = assess_recurrence_risk(combined_text, resolution_comments, root_cause)
    
    # Generate preventive actions
    preventive_actions = generate_case_specific_preventive_actions(issue_identified, root_cause, integration, resolution_method)
    
    # Assess holiday season impact
    holiday_impact = assess_holiday_impact(combined_text, customer_impact, recurrence_risk)
    
    # Determine urgency level
    urgency_level = determine_urgency_level(priority, holiday_impact, recurrence_risk)
    
    # Will this happen again?
    will_happen_again = determine_if_will_happen_again(root_cause, resolution_method, recurrence_risk)
    
    # How to prevent
    how_to_prevent = generate_specific_prevention_steps(root_cause, integration, resolution_method)
    
    return {
        'issue_identified': issue_identified,
        'root_cause': root_cause,
        'resolution_method': resolution_method,
        'technical_details': technical_details,
        'customer_impact': customer_impact,
        'recurrence_risk': recurrence_risk,
        'preventive_actions': preventive_actions,
        'holiday_impact': holiday_impact,
        'urgency_level': urgency_level,
        'will_happen_again': will_happen_again,
        'how_to_prevent': how_to_prevent
    }

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

def extract_technical_details(resolution_comments):
    """Extract technical details from resolution comments"""
    
    if pd.isna(resolution_comments) or str(resolution_comments).strip() in ['', 'nan']:
        return 'No technical details available'
    
    comments_text = str(resolution_comments)
    
    # Extract specific technical information
    technical_details = []
    
    # API details
    if 'api' in comments_text.lower():
        api_match = re.search(r'api[^\.]*', comments_text.lower())
        if api_match:
            technical_details.append(f"API: {api_match.group(0)[:100]}")
    
    # Error codes
    error_codes = re.findall(r'\b[A-Z]{2,}\d+\b|\b\d{3}\b', comments_text)
    if error_codes:
        technical_details.append(f"Error Codes: {', '.join(error_codes[:5])}")
    
    # URLs or endpoints
    urls = re.findall(r'https?://[^\s]+', comments_text)
    if urls:
        technical_details.append(f"URLs: {', '.join(urls[:3])}")
    
    # File names or paths
    files = re.findall(r'[A-Za-z0-9_\-\.]+\.(?:json|xml|csv|txt|log)', comments_text)
    if files:
        technical_details.append(f"Files: {', '.join(files[:3])}")
    
    # Field names
    fields = re.findall(r'[A-Za-z_][A-Za-z0-9_]*\s*(?:field|mapping)', comments_text.lower())
    if fields:
        technical_details.append(f"Fields: {', '.join(fields[:3])}")
    
    return '; '.join(technical_details) if technical_details else 'General technical issue'

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

def generate_case_specific_preventive_actions(issue_identified, root_cause, integration, resolution_method):
    """Generate specific preventive actions for this case"""
    
    actions = []
    
    # Root cause specific actions
    if root_cause == 'Configuration Error':
        actions.extend([
            'Implement configuration validation tools',
            'Add configuration testing and preview',
            'Create configuration templates and best practices',
            'Implement configuration health checks'
        ])
    
    elif root_cause == 'Data Mapping Issue':
        actions.extend([
            'Implement field mapping validation',
            'Add mapping preview and testing',
            'Create mapping templates',
            'Implement duplicate detection'
        ])
    
    elif root_cause == 'Authentication Failure':
        actions.extend([
            'Implement automated token refresh',
            'Add credential validation',
            'Create authentication monitoring',
            'Implement token expiration warnings'
        ])
    
    elif root_cause == 'API Limitations':
        actions.extend([
            'Implement API rate limit monitoring',
            'Add API health checks',
            'Create API quota management',
            'Implement API retry logic'
        ])
    
    elif root_cause == 'Data Synchronization Problem':
        actions.extend([
            'Implement sync monitoring',
            'Add automated retry mechanisms',
            'Create sync health checks',
            'Implement sync queue management'
        ])
    
    # Resolution method specific actions
    if resolution_method == 'Workaround Applied':
        actions.extend([
            'Plan permanent fix to replace workaround',
            'Document workaround limitations',
            'Create monitoring for workaround effectiveness',
            'Implement automated permanent fix deployment'
        ])
    
    elif resolution_method == 'Customer Guidance':
        actions.extend([
            'Create self-service documentation',
            'Implement guided setup wizards',
            'Add validation and error prevention',
            'Create customer education materials'
        ])
    
    # Integration specific actions
    if integration:
        if 'NetSuite' in integration:
            actions.extend(['Add NetSuite-specific monitoring', 'Implement NetSuite health checks'])
        if 'Amazon' in integration:
            actions.extend(['Add Amazon API monitoring', 'Implement Amazon quota management'])
        if 'Salesforce' in integration:
            actions.extend(['Add Salesforce monitoring', 'Implement Salesforce health checks'])
        if 'Shopify' in integration:
            actions.extend(['Add Shopify API monitoring', 'Implement Shopify rate limit management'])
    
    # Holiday specific actions
    actions.extend([
        'Implement holiday season monitoring',
        'Add holiday season capacity planning',
        'Create holiday season runbooks',
        'Implement holiday season alerting'
    ])
    
    # Remove duplicates and limit
    unique_actions = list(dict.fromkeys(actions))
    return '; '.join(unique_actions[:8])

def assess_holiday_impact(combined_text, customer_impact, recurrence_risk):
    """Assess the impact during holiday season"""
    
    # High holiday impact
    if customer_impact == 'High' and recurrence_risk == 'High':
        return 'High'
    
    # Medium holiday impact
    elif customer_impact == 'High' or recurrence_risk == 'High':
        return 'Medium'
    
    # Check for holiday-specific keywords
    if any(word in combined_text for word in ['holiday', 'peak', 'high volume', 'seasonal', 'busy']):
        return 'High'
    
    return 'Low'

def determine_urgency_level(priority, holiday_impact, recurrence_risk):
    """Determine urgency level for preventive action"""
    
    if priority in ['P1', 'Critical', 'Blocker'] or holiday_impact == 'High':
        return 'High'
    
    elif priority in ['P2', 'High', 'Major'] or recurrence_risk == 'High':
        return 'Medium'
    
    return 'Low'

def determine_if_will_happen_again(root_cause, resolution_method, recurrence_risk):
    """Determine if this issue will happen again"""
    
    if recurrence_risk == 'High':
        return 'YES - High Risk'
    
    elif resolution_method == 'Workaround Applied':
        return 'YES - Workaround Applied'
    
    elif root_cause in ['Configuration Error', 'Data Mapping Issue', 'Authentication Failure']:
        return 'YES - Systemic Issue'
    
    elif root_cause in ['Code/Script Error', 'External System Issue']:
        return 'MAYBE - Depends on Fix'
    
    return 'NO - Likely Fixed'

def generate_specific_prevention_steps(root_cause, integration, resolution_method):
    """Generate specific prevention steps"""
    
    steps = []
    
    if root_cause == 'Configuration Error':
        steps.extend([
            'Add configuration validation before deployment',
            'Create configuration templates for common setups',
            'Implement configuration health checks',
            'Add configuration testing environment'
        ])
    
    elif root_cause == 'Data Mapping Issue':
        steps.extend([
            'Implement field mapping validation tools',
            'Add mapping preview and testing',
            'Create mapping templates',
            'Add duplicate detection and prevention'
        ])
    
    elif root_cause == 'Authentication Failure':
        steps.extend([
            'Implement automated token refresh',
            'Add credential validation and health checks',
            'Create authentication monitoring dashboard',
            'Add token expiration warnings and alerts'
        ])
    
    elif root_cause == 'API Limitations':
        steps.extend([
            'Implement API rate limit monitoring',
            'Add API health checks and circuit breakers',
            'Create API quota management system',
            'Add API retry logic with exponential backoff'
        ])
    
    elif root_cause == 'Data Synchronization Problem':
        steps.extend([
            'Implement sync monitoring dashboard',
            'Add automated retry mechanisms',
            'Create sync health checks',
            'Implement sync queue management'
        ])
    
    # Integration specific steps
    if integration:
        if 'NetSuite' in integration:
            steps.append('Add NetSuite-specific monitoring and health checks')
        if 'Amazon' in integration:
            steps.append('Add Amazon API monitoring and quota management')
        if 'Salesforce' in integration:
            steps.append('Add Salesforce monitoring and health checks')
        if 'Shopify' in integration:
            steps.append('Add Shopify API monitoring and rate limit management')
    
    # Holiday specific steps
    steps.extend([
        'Implement holiday season monitoring dashboard',
        'Add holiday season capacity planning',
        'Create holiday season runbooks and procedures',
        'Implement holiday season alerting system'
    ])
    
    # Remove duplicates and limit
    unique_steps = list(dict.fromkeys(steps))
    return '; '.join(unique_steps[:6])

# ============================================================================
# COMMAND LINE INTERFACE
# ============================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Simplified Holiday Case Analysis - Individual Case Analysis only',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python3 simplified_holiday_analysis.py --file Holiday.csv
  python3 simplified_holiday_analysis.py --file /path/to/holiday.csv --output Analysis.xlsx
  
This tool provides:
  - Individual case analysis with detailed resolution comment examination
  - Root cause analysis for each case
  - Specific preventive measures for each case
  - Holiday season impact assessment
  - Will this happen again assessment
        ''',
        add_help=True
    )
    
    parser.add_argument('--file', '-f', default='~/Downloads/Holiday.csv',
                       help='Path to the Holiday CSV file')
    parser.add_argument('--output', '-o',
                       help='Output Excel file name (optional)')
    
    args = parser.parse_args()
    
    try:
        output_file = analyze_individual_cases_only(args.file, args.output)
        if output_file:
            print(f"\n🎯 Simplified holiday case analysis complete: {output_file}")
    except Exception as e:
        print(f"\n❌ Error during analysis: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

