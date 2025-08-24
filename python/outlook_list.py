#!/usr/bin/env python3
"""
Unified ls-like tool for Outlook.
Lists accounts, folders, or items depending on the path provided.
"""

import win32com.client
import argparse
import sys
import json
from datetime import datetime
import re
from outlook_utils import encode_entry_id


def safe_text(text, max_length=None):
    """Convert text to ASCII-safe string, optionally truncating."""
    if text is None:
        return ""
    safe = str(text).encode('ascii', 'ignore').decode('ascii')
    if max_length and len(safe) > max_length:
        return safe[:max_length]
    return safe


def format_compact_date(dt):
    """Format datetime to compact string like 'Aug 14' or 'Aug 14 15:30'."""
    if dt is None:
        return None
    
    try:
        # Convert string to datetime if needed
        if isinstance(dt, str):
            dt = datetime.fromisoformat(dt.replace('Z', '+00:00'))
        
        # Format based on whether time is meaningful
        now = datetime.now()
        if dt.date() == now.date():
            # Today - show time
            return dt.strftime('%H:%M')
        elif dt.year == now.year:
            # This year - just month and day
            return dt.strftime('%b %d')
        else:
            # Different year - include year
            return dt.strftime('%b %d %Y')
    except:
        return str(dt)[:16] if dt else None


def clean_dict(d):
    """Remove null, empty string, and empty array values from dictionary."""
    cleaned = {}
    for key, value in d.items():
        if value is not None and value != "" and value != []:
            cleaned[key] = value
    return cleaned


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
    elif len(parts) == 2:
        return (parts[0], parts[1])
    else:
        # For deeper paths, just take first two levels for now
        return (parts[0], parts[1])


def find_account(namespace, pattern):
    """
    Find an account by pattern matching.
    Returns (account_index, account_folder) or (None, None)
    """
    if not pattern:
        return (None, None)
    
    pattern_lower = pattern.lower()
    
    # Try exact match first
    for i in range(1, namespace.Folders.Count + 1):
        try:
            folder = namespace.Folders.Item(i)
            if folder.Name.lower() == pattern_lower:
                return (i, folder)
        except:
            pass
    
    # Try partial match
    for i in range(1, namespace.Folders.Count + 1):
        try:
            folder = namespace.Folders.Item(i)
            if pattern_lower in folder.Name.lower():
                return (i, folder)
        except:
            pass
    
    return (None, None)


def find_folder(account, folder_name):
    """
    Find a folder within an account.
    Returns folder object or None.
    """
    if not folder_name:
        return None
    
    # Try exact match
    try:
        return account.Folders[folder_name]
    except:
        pass
    
    # Try case-insensitive match
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
    """Legacy format date function - kept for compatibility."""
    if not date_obj:
        return "          "
    
    now = datetime.now()
    if date_obj.date() == now.date():
        return date_obj.strftime("%H:%M     ")
    elif date_obj.year == now.year:
        return date_obj.strftime("%b %d %H:%M")
    else:
        return date_obj.strftime("%Y-%m-%d")


def list_accounts(namespace, show_all=False, show_empty=False):
    """
    List all Outlook accounts and return as JSON.
    """
    accounts = []
    
    for i in range(1, namespace.Folders.Count + 1):
        try:
            account = namespace.Folders.Item(i)
            account_name = safe_text(account.Name)
            
            # Skip certain accounts unless show_all
            if not show_all:
                skip_patterns = ['Public Folders', 'Online Archive']
                if any(pattern in account_name for pattern in skip_patterns):
                    continue
            
            # Count folders (excluding system folders)
            folder_count = 0
            email_folder_count = 0
            latest_date = None
            
            for j in range(1, account.Folders.Count + 1):
                try:
                    folder = account.Folders.Item(j)
                    folder_name = folder.Name
                    
                    # Skip system folders in count unless show_all
                    if not show_all:
                        system_folders = {'Sync Issues', 'Conflicts', 'Local Failures', 
                                        'Server Failures', 'PersonMetadata', 'ExternalContacts',
                                        'MeContact', 'PeopleCentricConversation Buddies',
                                        'Recipient Cache', 'GAL Contacts', 'Organizational Contacts'}
                        if folder_name in system_folders or folder_name.startswith('{'):
                            continue
                    
                    folder_count += 1
                    
                    # Check if it's an email folder
                    try:
                        if folder.DefaultItemType == 0:  # Mail folder
                            email_folder_count += 1
                            # Try to get latest item date from important folders
                            if folder_name in ['Inbox', 'Sent Items']:
                                if folder.Items.Count > 0:
                                    items = folder.Items
                                    items.Sort('[ReceivedTime]', True)
                                    try:
                                        item = items.Item(1)
                                        item_date = item.ReceivedTime
                                        if not latest_date or item_date > latest_date:
                                            latest_date = item_date
                                    except:
                                        pass
                    except:
                        pass
                except:
                    pass
            
            account_data = clean_dict({
                'name': account_name,
                'folders': folder_count,
                'emails': email_folder_count if email_folder_count > 0 else None,
                'activity': format_compact_date(latest_date),
                'type': 'account'
            })
            accounts.append(account_data)
        except:
            pass
    
    # Return JSON output
    output = {
        'type': 'accounts',
        'count': len(accounts),
        'results': accounts
    }
    print(json.dumps(output, indent=2, default=str))


def list_folders(account, show_all=False, show_empty=False):
    """
    List folders within an account and return as JSON.
    """
    folders = []
    
    # System folders to skip by default
    system_folders = {'Sync Issues', 'Conflicts', 'Local Failures', 
                     'Server Failures', 'PersonMetadata', 'ExternalContacts',
                     'MeContact', 'PeopleCentricConversation Buddies',
                     'Recipient Cache', 'GAL Contacts', 'Organizational Contacts',
                     'Companies', 'Birthdays', 'Journal', 'Notes',
                     'Conversation Action Settings', 'Quick Step Settings',
                     'Yammer Root', 'Social Activity Notifications', 'Files'}
    
    for i in range(1, account.Folders.Count + 1):
        try:
            folder = account.Folders.Item(i)
            folder_name = safe_text(folder.Name)
            
            # Skip system folders unless show_all
            if not show_all:
                if folder_name in system_folders or folder_name.startswith('{'):
                    continue
            
            # Get folder info
            item_count = 0
            folder_type = "Unknown"
            latest_date = None
            
            try:
                item_count = folder.Items.Count
                
                # Get folder type
                if hasattr(folder, 'DefaultItemType'):
                    item_type = folder.DefaultItemType
                    type_map = {
                        0: "Mail",
                        1: "Calendar",
                        2: "Contact",
                        3: "Task",
                        4: "Journal",
                        5: "Note",
                        6: "Post"
                    }
                    folder_type = type_map.get(item_type, f"Type_{item_type}")
                
                # Get latest item date for important folders
                if item_count > 0 and folder_type == "Mail":
                    items = folder.Items
                    items.Sort('[ReceivedTime]', True)
                    try:
                        item = items.Item(1)
                        latest_date = item.ReceivedTime
                    except:
                        pass
                elif item_count > 0 and folder_type == "Calendar":
                    items = folder.Items
                    items.Sort('[Start]', True)
                    try:
                        item = items.Item(1)
                        latest_date = item.Start
                    except:
                        pass
            except:
                pass
            
            # Skip empty folders unless show_empty is True
            if item_count == 0 and not show_empty and not show_all:
                continue
            
            folder_data = clean_dict({
                'name': folder_name,
                'count': item_count,
                'type': folder_type,
                'activity': format_compact_date(latest_date)
            })
            folders.append(folder_data)
        except:
            pass
    
    # Sort folders by activity (most recent first), with empty folders at the end
    def get_activity_timestamp(folder):
        activity = folder.get('activity')
        if not activity:
            return 0
        try:
            # Parse the compact date format back to datetime
            now = datetime.now()
            if ':' in activity:  # Time format like "10:03"
                return now.replace(hour=int(activity.split(':')[0]), minute=int(activity.split(':')[1])).timestamp()
            elif len(activity.split()) == 2:  # Format like "Aug 22"
                month_str, day_str = activity.split()
                months = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                         'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
                month = months.get(month_str, 1)
                day = int(day_str)
                return datetime(now.year, month, day).timestamp()
            elif len(activity.split()) == 3:  # Format like "Aug 22 2024"
                month_str, day_str, year_str = activity.split()
                months = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                         'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
                month = months.get(month_str, 1)
                day = int(day_str)
                year = int(year_str)
                return datetime(year, month, day).timestamp()
            else:
                return 0
        except:
            return 0
    
    folders.sort(key=lambda x: (
        x['count'] == 0,  # Empty folders last
        -get_activity_timestamp(x),  # Most recent activity first
        x['name'].lower()
    ))
    
    # Return JSON output
    output = {
        'type': 'folders',
        'account': safe_text(account.Name),
        'count': len(folders),
        'results': folders
    }
    print(json.dumps(output, indent=2, default=str))


def list_items(folder, count=10, show_all=False):
    """
    List items within a folder and return as JSON.
    """
    folder_type = "Unknown"
    try:
        if hasattr(folder, 'DefaultItemType'):
            item_type = folder.DefaultItemType
            type_map = {0: "Mail", 1: "Calendar", 2: "Contact", 3: "Task"}
            folder_type = type_map.get(item_type, "Unknown")
    except:
        pass
    
    total_items = folder.Items.Count
    
    if total_items == 0:
        output = {
            'type': 'items',
            'folder': safe_text(folder.Name),
            'folder_type': folder_type,
            'total': 0,
            'count': 0,
            'results': []
        }
        print(json.dumps(output, indent=2, default=str))
        return
    
    # Determine how many items to show
    if show_all:
        items_to_show = total_items
    else:
        items_to_show = min(count, total_items)
    
    items = folder.Items
    results = []
    
    # Sort and optimize based on folder type
    if folder_type == "Mail":
        items.Sort('[ReceivedTime]', True)
        
        for i in range(1, items_to_show + 1):
            try:
                item = items.Item(i)
                item_data = clean_dict({
                    'subject': safe_text(item.Subject, 100),
                    'from': safe_text(item.SenderName, 50),
                    'date': format_compact_date(item.ReceivedTime),
                    'unread': True if getattr(item, 'UnRead', False) else None,
                    'entry_id': encode_entry_id(item.EntryID)
                })
                results.append(item_data)
            except:
                pass
                
    elif folder_type == "Calendar":
        items.Sort('[Start]', True)
        items.IncludeRecurrences = False
        
        for i in range(1, items_to_show + 1):
            try:
                item = items.Item(i)
                item_data = clean_dict({
                    'subject': safe_text(item.Subject, 100),
                    'start': format_compact_date(item.Start),
                    'location': safe_text(getattr(item, 'Location', ''), 50) if getattr(item, 'Location', '') else None,
                    'entry_id': encode_entry_id(item.EntryID)
                })
                results.append(item_data)
            except:
                pass
                
    elif folder_type == "Contact":
        items.Sort('[LastName]', False)
        
        for i in range(1, items_to_show + 1):
            try:
                item = items.Item(i)
                item_data = clean_dict({
                    'name': safe_text(item.FullName, 50),
                    'email': safe_text(getattr(item, 'Email1Address', ''), 50) if getattr(item, 'Email1Address', '') else None,
                    'company': safe_text(getattr(item, 'CompanyName', ''), 50) if getattr(item, 'CompanyName', '') else None,
                    'entry_id': encode_entry_id(item.EntryID)
                })
                results.append(item_data)
            except:
                pass
                
    elif folder_type == "Task":
        items.Sort('[DueDate]', False)
        
        for i in range(1, items_to_show + 1):
            try:
                item = items.Item(i)
                status = getattr(item, 'Status', 0)
                status_map = {0: "Not started", 1: "In progress", 2: "Complete"}
                item_data = clean_dict({
                    'subject': safe_text(item.Subject, 100),
                    'due': format_compact_date(getattr(item, 'DueDate', None)),
                    'status': status_map.get(status, "Unknown"),
                    'complete': getattr(item, 'PercentComplete', 0) if getattr(item, 'PercentComplete', 0) > 0 else None,
                    'entry_id': encode_entry_id(item.EntryID)
                })
                results.append(item_data)
            except:
                pass
    else:
        # Generic listing for unknown types
        for i in range(1, items_to_show + 1):
            try:
                item = items.Item(i)
                results.append({
                    'description': safe_text(getattr(item, 'Subject', str(item)), 100),
                    'entry_id': encode_entry_id(getattr(item, 'EntryID', None))
                })
            except:
                pass
    
    # Build output
    output = {
        'type': 'items',
        'folder': safe_text(folder.Name),
        'folder_type': folder_type,
        'total': total_items,
        'count': len(results),
        'has_more': items_to_show < total_items,
        'results': results
    }
    
    if output['has_more']:
        output['remaining'] = total_items - items_to_show
    
    print(json.dumps(output, indent=2, default=str))


def main():
    parser = argparse.ArgumentParser(
        description='List Outlook accounts, folders, or items in ls-like format',
        epilog='Performance tip: Use this tool to identify folders before searching with outlook_search\n\n'
               'Examples:\n'
               '  outlook_list                      # List all accounts\n'
               '  outlook_list myaccount            # List folders in account\n'
               '  outlook_list myaccount/Inbox      # List items in folder\n',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument('path', nargs='?', default='',
                        help='Path to list (account or account/folder)')
    parser.add_argument('-a', '--all', action='store_true',
                        help='Show all items/folders including system folders')
    parser.add_argument('-c', '--count', type=int, default=10,
                        help='Number of items to show (default: 10)')
    parser.add_argument('--show-empty', action='store_true',
                        help='Show empty folders (hidden by default)')
    parser.add_argument('-l', '--long', action='store_true',
                        help='Long format (more details) - not yet implemented')
    
    args = parser.parse_args()
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Parse the path
        account_pattern, folder_name = parse_outlook_path(args.path)
        
        if not account_pattern:
            # No path - list all accounts
            list_accounts(namespace, args.all, args.show_empty)
        elif not folder_name:
            # Account only - list folders
            account_idx, account = find_account(namespace, account_pattern)
            if account:
                list_folders(account, args.all, args.show_empty)
            else:
                print(f"Error: Account '{account_pattern}' not found")
                print("\nAvailable accounts:")
                list_accounts(namespace, args.all, args.show_empty)
                return 1
        else:
            # Account and folder - list items
            account_idx, account = find_account(namespace, account_pattern)
            if not account:
                print(f"Error: Account '{account_pattern}' not found")
                return 1
            
            folder = find_folder(account, folder_name)
            if folder:
                list_items(folder, args.count, args.all)
            else:
                print(f"Error: Folder '{folder_name}' not found in '{account.Name}'")
                print("\nAvailable folders:")
                list_folders(account, args.all, args.show_empty)
                return 1
        
        return 0
        
    except Exception as e:
        print(f"Error: {e}")
        print("\nMake sure Outlook is running and accessible.")
        return 1


if __name__ == "__main__":
    sys.exit(main())