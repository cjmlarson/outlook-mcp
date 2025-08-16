#!/usr/bin/env python3
"""
Search Outlook items using fast DASL queries.
Requires a specific path - use outlook_list to find folders first.

Search syntax:
- Space = OR: "United ZRH" finds items with either term (intuitive default)
- Ampersand = AND: "United&ZRH" finds items with both terms
- Combined: "ZRH EWR&United" finds items with United AND (ZRH OR EWR)
- Legacy pipe still works: "ZRH|EWR" same as "ZRH EWR"
"""

import win32com.client
import argparse
import sys
import json
import re
from datetime import datetime
from outlook_utils import encode_entry_id

def safe_text(text, max_length=None):
    """Convert text to ASCII-safe string."""
    if text is None:
        return ""
    safe = str(text).encode('ascii', 'ignore').decode('ascii')
    if max_length and len(safe) > max_length:
        return safe[:max_length]
    return safe

def parse_outlook_path(path_str):
    """Parse 'account/folder/subfolder' into (account, folder_path)"""
    parts = path_str.split('/')
    if len(parts) == 1:
        return (parts[0], None)  # Just account
    return (parts[0], '/'.join(parts[1:]))  # Account and folder path

def get_folder(namespace, account_name, folder_path):
    """Get specific folder object from path"""
    # Find account by partial match
    account = None
    for acc in namespace.Folders:
        if account_name.lower() in acc.Name.lower():
            account = acc
            break
    
    if not account:
        return None
    
    if not folder_path:
        return account  # Return account root
    
    # Navigate to specific folder
    folder = account
    for part in folder_path.split('/'):
        try:
            folder = folder.Folders.Item(part)
        except:
            return None
    
    return folder

def parse_search_pattern(pattern):
    """
    Parse search patterns:
    - "term1 term2" -> OR search (space-separated, intuitive default)
    - "term1&term2" -> AND search (ampersand-separated)
    - "term1 term2&term3" -> (term1 OR term2) AND term3
    """
    # Split by & to get AND groups
    and_groups = pattern.split('&')
    
    if len(and_groups) == 1:
        # Single group - check for legacy OR syntax or split by spaces for OR
        if '|' in and_groups[0]:
            return 'OR', and_groups[0].split('|')
        else:
            # Split by spaces for OR (new default behavior)
            terms = and_groups[0].split()
            if len(terms) == 1:
                return 'SIMPLE', terms
            else:
                return 'OR', terms
    
    # Multiple groups - AND them together
    parsed_groups = []
    for group in and_groups:
        group = group.strip()
        if '|' in group:
            # This group contains legacy OR terms
            parsed_groups.append(('OR', group.split('|')))
        else:
            # Split by spaces for OR within this AND group
            terms = group.split()
            if len(terms) == 1:
                parsed_groups.append(('SIMPLE', terms))
            else:
                parsed_groups.append(('OR', terms))
    
    return 'AND_GROUPS', parsed_groups

def build_dasl_filter(pattern, folder, since=None, until=None):
    """Build DASL filter with smart ci_phrasematch/LIKE selection"""
    
    # Check if this is a calendar folder
    is_calendar = False
    try:
        is_calendar = (folder.DefaultItemType == 1)  # 1 = olAppointmentItem
    except:
        pass
    
    # Check if store supports instant search
    try:
        store = folder.Store
        use_ci = store.IsInstantSearchEnabled
    except:
        use_ci = False  # Default to LIKE if can't determine
    
    # Parse the search pattern
    parse_result = parse_search_pattern(pattern)
    mode, data = parse_result
    
    # Build text search filter based on pattern type
    if mode == 'SIMPLE':
        # Single term
        term = data[0]
        if use_ci:
            text_filter = f'"urn:schemas:httpmail:textdescription" ci_phrasematch \'{term}\''
        else:
            text_filter = f'"urn:schemas:httpmail:textdescription" LIKE \'%{term}%\''
    
    elif mode == 'OR':
        # Multiple OR terms
        filters = []
        for term in data:
            if use_ci:
                filters.append(f'"urn:schemas:httpmail:textdescription" ci_phrasematch \'{term}\'')
            else:
                filters.append(f'"urn:schemas:httpmail:textdescription" LIKE \'%{term}%\'')
        text_filter = f'({" OR ".join(filters)})'
    
    elif mode == 'AND_GROUPS':
        # Complex pattern with AND and OR
        and_parts = []
        for group_mode, terms in data:
            if group_mode == 'OR':
                # Build OR clause for this group
                or_filters = []
                for term in terms:
                    if use_ci:
                        or_filters.append(f'"urn:schemas:httpmail:textdescription" ci_phrasematch \'{term}\'')
                    else:
                        or_filters.append(f'"urn:schemas:httpmail:textdescription" LIKE \'%{term}%\'')
                and_parts.append(f'({" OR ".join(or_filters)})')
            else:
                # Single term in this AND group
                term = terms[0]
                if use_ci:
                    and_parts.append(f'"urn:schemas:httpmail:textdescription" ci_phrasematch \'{term}\'')
                else:
                    and_parts.append(f'"urn:schemas:httpmail:textdescription" LIKE \'%{term}%\'')
        text_filter = f'({" AND ".join(and_parts)})'
    
    # Add date filters if specified
    filters = [text_filter]
    
    # Use appropriate date field based on folder type
    date_field = '[Start]' if is_calendar else '[ReceivedTime]'
    
    if since:
        date_str = since.strftime('%m/%d/%Y')
        filters.append(f'{date_field} >= \'{date_str}\'')
    
    if until:
        date_str = until.strftime('%m/%d/%Y')
        filters.append(f'{date_field} <= \'{date_str}\'')
    
    # Combine all filters with AND
    if len(filters) == 1:
        return f'@SQL={filters[0]}'
    else:
        return f'@SQL=({" AND ".join(filters)})'

def search_folder(folder, pattern, args):
    """Search folder using DASL Restrict for massive performance gain"""
    
    # Check if this is a calendar folder
    is_calendar = False
    try:
        is_calendar = (folder.DefaultItemType == 1)  # 1 = olAppointmentItem
    except:
        pass
    
    # For calendar folders, enable recurrence support
    items = folder.Items
    if is_calendar:
        items.IncludeRecurrences = True
    
    # Build the DASL filter
    filter_str = build_dasl_filter(pattern, folder, args.since, args.until)
    
    try:
        # Use Restrict to get ONLY matching items
        # This is the KEY performance improvement - no manual iteration!
        items = items.Restrict(filter_str)
        
        # Sort by date (newest first) - use appropriate field
        date_field = '[Start]' if is_calendar else '[ReceivedTime]'
        items.Sort(date_field, True)
        
        results = []
        
        # Process results based on output mode
        for item in items:
            # Get folder path for display
            try:
                folder_path = f"{folder.Parent.Name}/{folder.Name}"
            except:
                folder_path = folder.Name
            
            # Build result dict with metadata
            # Get appropriate date field based on item type
            if hasattr(item, 'Start'):
                date_value = item.Start.isoformat() if item.Start else None
            elif hasattr(item, 'ReceivedTime'):
                date_value = item.ReceivedTime.isoformat() if item.ReceivedTime else None
            else:
                date_value = None
            
            result = {
                'entry_id': encode_entry_id(item.EntryID),  # Base64 encoded to save tokens
                'subject': safe_text(item.Subject),
                'from': safe_text(getattr(item, 'SenderName', getattr(item, 'Organizer', ''))),
                'date': date_value,
                'path': folder_path
            }
            
            # For content mode, extract match snippets
            if args.output_mode == 'content':
                matches = []
                
                # Get body text (try Body first, fall back to HTMLBody)
                body = str(getattr(item, 'Body', ''))
                if not body and hasattr(item, 'HTMLBody'):
                    # Strip HTML tags for snippet extraction
                    html = str(item.HTMLBody)
                    body = re.sub('<[^<]+?>', '', html)
                
                # Find pattern occurrences for context
                # Extract all search terms from the pattern
                search_terms = []
                for part in pattern.replace('|', ' ').split():
                    if part:
                        search_terms.append(part)
                
                for term in search_terms:
                    # Find matches of this term
                    try:
                        regex = re.compile(re.escape(term), re.IGNORECASE)
                        for match in regex.finditer(body):
                            # Extract context around match
                            start = max(0, match.start() - 50)
                            end = min(len(body), match.end() + 50)
                            context = body[start:end].replace('\r\n', ' ').replace('\n', ' ').strip()
                            
                            matches.append({
                                'term': term,
                                'context': f'...{context}...'
                            })
                            
                            # Limit to 3 snippets per item
                            if len(matches) >= 3:
                                break
                    except:
                        continue
                    
                    if len(matches) >= 3:
                        break
                
                result['matches'] = matches
                result['match_count'] = len(matches)
            
            results.append(result)
        
        return results
        
    except Exception as e:
        print(f"Search error: {e}", file=sys.stderr)
        return []

def display_results(results, output_mode='list', offset=0):
    """Display results as JSON with pagination"""
    
    PAGE_SIZE = 25
    total = len(results)
    
    # Apply pagination
    start_idx = offset
    end_idx = min(start_idx + PAGE_SIZE, total)
    paginated_results = results[start_idx:end_idx]
    
    # Build JSON output
    output = {
        'pagination': {
            'total': total,
            'offset': offset,
            'limit': PAGE_SIZE,
            'has_more': end_idx < total
        },
        'results': paginated_results
    }
    
    # Output as JSON
    print(json.dumps(output, indent=2, default=str))
    
    # Add navigation hints if there are more results
    if output['pagination']['has_more']:
        next_offset = offset + PAGE_SIZE
        print(f"\n# For next page, use: --offset {next_offset}", file=sys.stderr)

def main():
    parser = argparse.ArgumentParser(
        description='Search Outlook using fast DASL queries',
        epilog='Search syntax:\n'
               '  Space = OR: "United ZRH" finds items with either term\n'
               '  Ampersand = AND: "United&ZRH" finds items with both\n'
               '  Combined: "ZRH EWR&United" = (ZRH OR EWR) AND United\n\n'
               'Examples:\n'
               '  outlook_search "ZRH" connor.larson@outlook.com/Archive\n'
               '  outlook_search "ZRH EWR JFK" myaccount/Inbox\n'
               '  outlook_search "flight&ZRH EWR" myaccount/Sent Items\n\n'
               'Tip: Use outlook_list first to find available folders',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    # Required arguments
    parser.add_argument('pattern', 
                       help='Search pattern (use quotes for multiple terms)')
    parser.add_argument('path',
                       help='Path to search (e.g., "account/Inbox")')
    
    # Optional arguments
    parser.add_argument('--output-mode', choices=['list', 'content'],
                       default='list',
                       help='Output mode: list (fast) or content (with snippets)')
    parser.add_argument('--since', type=lambda s: datetime.strptime(s, '%Y-%m-%d'),
                       help='Filter by date (YYYY-MM-DD)')
    parser.add_argument('--until', type=lambda s: datetime.strptime(s, '%Y-%m-%d'),
                       help='Filter by date (YYYY-MM-DD)')
    parser.add_argument('--offset', type=int, default=0,
                       help='Pagination offset (results shown 25 at a time)')
    
    args = parser.parse_args()
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Parse the path
        account_name, folder_path = parse_outlook_path(args.path)
        
        # Validate that a specific folder is provided, not just account
        if folder_path is None:
            print(f"Error: Must specify a folder, not just account: {args.path}", file=sys.stderr)
            print("\nAccount-only paths return no results. Use outlook_list to find folders:", file=sys.stderr)
            print(f"  outlook_list {account_name}       # List folders in account", file=sys.stderr)
            print(f"  Then search specific folder:", file=sys.stderr)
            print(f"  outlook_search \"pattern\" {account_name}/Inbox", file=sys.stderr)
            return 1
        
        # Get the folder
        folder = get_folder(namespace, account_name, folder_path)
        
        if not folder:
            print(f"Error: Folder not found: {args.path}", file=sys.stderr)
            print("\nUse outlook_list to find available folders:", file=sys.stderr)
            print("  outlook_list                    # List accounts", file=sys.stderr)
            print("  outlook_list account_name       # List folders", file=sys.stderr)
            print("  outlook_list account_name/path  # List subfolders", file=sys.stderr)
            return 1
        
        # Perform the search
        results = search_folder(folder, args.pattern, args)
        
        # Display results with pagination
        display_results(results, args.output_mode, args.offset)
        
        return 0
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

if __name__ == "__main__":
    sys.exit(main())