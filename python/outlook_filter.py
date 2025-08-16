#!/usr/bin/env python3
"""
Filter Outlook items by properties (equivalent to Unix find command).
Filters by date, sender, type, read status, etc. without content search.
"""

import win32com.client
import argparse
import sys
import json
from datetime import datetime, timedelta
from outlook_utils import encode_entry_id


def safe_text(text, max_length=None):
    """Convert text to ASCII-safe string, optionally truncating."""
    if text is None:
        return ""
    safe = str(text).encode('ascii', 'ignore').decode('ascii')
    if max_length and len(safe) > max_length:
        return safe[:max_length]
    return safe


def parse_outlook_path(path_str):
    """
    Parse an Outlook path like 'myaccount@domain.com/Inbox'
    Returns (account_pattern, folder_name)
    """
    if not path_str:
        return (None, None)
    
    parts = path_str.rstrip('/').split('/')
    if len(parts) == 1:
        return (parts[0], None)
    elif len(parts) >= 2:
        return (parts[0], parts[1])


def find_account(namespace, pattern):
    """Find an account by pattern matching."""
    if not pattern:
        return None
    
    pattern_lower = pattern.lower()
    
    # Try exact match first
    for i in range(1, namespace.Folders.Count + 1):
        try:
            folder = namespace.Folders.Item(i)
            if folder.Name.lower() == pattern_lower:
                return folder
        except:
            pass
    
    # Try partial match
    for i in range(1, namespace.Folders.Count + 1):
        try:
            folder = namespace.Folders.Item(i)
            if pattern_lower in folder.Name.lower():
                return folder
        except:
            pass
    
    return None


def get_folders_to_filter(namespace, account_pattern, folder_name):
    """Get folders to filter based on path."""
    folders = []
    
    if not account_pattern:
        # No path - filter all accounts
        for i in range(1, namespace.Folders.Count + 1):
            try:
                account = namespace.Folders.Item(i)
                if 'public' not in account.Name.lower():
                    folders.extend(get_email_folders(account))
            except:
                pass
    else:
        # Specific account
        account = find_account(namespace, account_pattern)
        if not account:
            return []
        
        if folder_name:
            # Specific folder
            folder = find_folder(account, folder_name)
            if folder:
                folders.append(folder)
        else:
            # All folders in account
            folders = get_email_folders(account)
    
    return folders


def get_email_folders(account):
    """Get all email-type folders from an account."""
    folders = []
    skip_folders = {'Sync Issues', 'Conflicts', 'Local Failures', 'Server Failures'}
    
    for i in range(1, account.Folders.Count + 1):
        try:
            folder = account.Folders.Item(i)
            if folder.Name not in skip_folders and not folder.Name.startswith('{'):
                # Check if it's an appropriate folder type
                try:
                    if folder.DefaultItemType in [0, 1, 2, 3]:  # Mail, Calendar, Contact, Task
                        folders.append(folder)
                except:
                    pass
        except:
            pass
    
    return folders


def find_folder(account, folder_name):
    """Find a folder within an account."""
    if not folder_name:
        return None
    
    folder_lower = folder_name.lower()
    for i in range(1, account.Folders.Count + 1):
        try:
            folder = account.Folders.Item(i)
            if folder.Name.lower() == folder_lower:
                return folder
        except:
            pass
    
    return None


def format_date(date_obj):
    """Format date for display."""
    if not date_obj:
        return "          "
    
    now = datetime.now()
    if date_obj.date() == now.date():
        # Windows-compatible format (without %-m/%-d)
        return date_obj.strftime("%m/%d %H:%M")
    elif date_obj.year == now.year:
        return date_obj.strftime("%m/%d")
    else:
        return date_obj.strftime("%Y-%m-%d")


def filter_items(folder, args):
    """Filter items in a folder based on criteria."""
    results = []
    
    try:
        items = folder.Items
        
        # Check if this is a calendar folder
        is_calendar = False
        try:
            is_calendar = (folder.DefaultItemType == 1)  # 1 = olAppointmentItem
        except:
            pass
        
        # For calendar folders, enable recurrence support
        if is_calendar:
            items.IncludeRecurrences = True
        
        # Build Restrict filter if possible
        filters = []
        
        # Date filters - use appropriate field based on folder type
        date_field = '[Start]' if is_calendar else '[ReceivedTime]'
        
        if args.since:
            since_date = datetime.strptime(args.since, '%Y-%m-%d')
            filters.append(f'{date_field} >= "{since_date.strftime("%m/%d/%Y")}"')
        
        if args.until:
            until_date = datetime.strptime(args.until, '%Y-%m-%d')
            filters.append(f'{date_field} <= "{until_date.strftime("%m/%d/%Y")}"')
        
        if args.days:
            start_date = datetime.now() - timedelta(days=args.days)
            filters.append(f'{date_field} >= "{start_date.strftime("%m/%d/%Y")}"')
        
        # Apply Restrict filter if we have any
        if filters:
            filter_str = ' AND '.join(filters)
            items = items.Restrict(filter_str)
        
        # Sort by date
        items.Sort(date_field, True)
        
        # Process items
        for item in items:
            try:
                # Type filter
                if args.type:
                    item_type = get_item_type(item)
                    if args.type.lower() not in item_type.lower():
                        continue
                
                # Sender filter
                if args.from_sender:
                    sender = getattr(item, 'SenderName', '')
                    if args.from_sender.lower() not in str(sender).lower():
                        continue
                
                # Unread filter
                if args.unread:
                    if not getattr(item, 'UnRead', False):
                        continue
                
                # Get account and folder info
                parent_folder = item.Parent
                account_name = get_account_name(parent_folder)
                
                # Get appropriate date field based on item type
                if hasattr(item, 'Start'):
                    date_value = getattr(item, 'Start', None)
                else:
                    date_value = getattr(item, 'ReceivedTime', None)
                
                # Build result
                result = {
                    'path': f"{account_name}/{parent_folder.Name}",
                    'entry_id': encode_entry_id(item.EntryID),
                    'from': safe_text(getattr(item, 'SenderName', getattr(item, 'Organizer', '')), 25),
                    'date': date_value,
                    'subject': safe_text(getattr(item, 'Subject', ''), 40)
                }
                results.append(result)
                
                # Limit results
                if len(results) >= args.max_items:
                    break
                    
            except Exception:
                continue
                
    except Exception:
        pass
    
    return results


def get_item_type(item):
    """Determine the type of an Outlook item."""
    try:
        if item.Class == 43:  # Mail
            return "email"
        elif item.Class == 26:  # Appointment
            return "event"
        elif item.Class == 40:  # Contact
            return "contact"
        elif item.Class == 48:  # Task
            return "task"
    except:
        pass
    return "unknown"


def get_account_name(folder):
    """Get the account name for a folder."""
    try:
        current = folder
        while current:
            parent = current.Parent
            if not parent or not hasattr(parent, 'Name'):
                return current.Name
            current = parent
    except:
        pass
    return "Unknown"


def main():
    parser = argparse.ArgumentParser(
        description='Filter Outlook items by properties (equivalent to Unix find command)',
        epilog='Performance tip: Use this to find relevant items without full content search\n\n'
               'Examples:\n'
               '  outlook_filter myaccount --days 7         # Recent items\n'
               '  outlook_filter --from "smith" --type email\n'
               '  outlook_filter myaccount/Archive --since 2024-01-01\n',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument('path', nargs='?', default='',
                        help='Path to filter (account or account/folder)')
    parser.add_argument('--since', type=str,
                        help='Start date (YYYY-MM-DD)')
    parser.add_argument('--until', type=str,
                        help='End date (YYYY-MM-DD)')
    parser.add_argument('--days', type=int,
                        help='Items from last N days')
    parser.add_argument('--from', dest='from_sender', type=str,
                        help='Filter by sender/organizer')
    parser.add_argument('--type', type=str,
                        help='Filter by type: email, event, contact, task')
    parser.add_argument('--unread', action='store_true',
                        help='Only unread items')
    parser.add_argument('--max-items', type=int, default=100,
                        help='Maximum items to return (default: 100)')
    
    args = parser.parse_args()
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Parse path
        account_pattern, folder_name = parse_outlook_path(args.path)
        
        # Get folders to filter
        folders = get_folders_to_filter(namespace, account_pattern, folder_name)
        if not folders:
            print("Error: No folders found to filter")
            return 1
        
        # Collect results from all folders
        all_results = []
        for folder in folders:
            results = filter_items(folder, args)
            all_results.extend(results)
            
            if len(all_results) >= args.max_items:
                all_results = all_results[:args.max_items]
                break
        
        # Display results as JSON
        output = {
            'total': len(all_results),
            'results': []
        }
        
        for item in all_results:
            output['results'].append({
                'path': item['path'],
                'subject': item['subject'],
                'from': item['from'],
                'date': format_date(item['date']) if item['date'] else None,
                'entry_id': item['entry_id']
            })
        
        print(json.dumps(output, indent=2, default=str))
        
        return 0
        
    except Exception as e:
        print(f"Error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())