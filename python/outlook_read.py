#!/usr/bin/env python3
"""
Read full Outlook item content by EntryID.
Works with emails, calendar events, contacts, tasks, and other Outlook items.
"""

import win32com.client
import argparse
import sys
import os
import json
import tempfile
from datetime import datetime
from outlook_utils import decode_entry_id


def safe_text(text):
    """Convert text to ASCII-safe string."""
    if text is None:
        return ""
    return str(text).encode('ascii', 'ignore').decode('ascii')


def format_date(date_obj):
    """Format date for display."""
    if not date_obj:
        return None
    try:
        return date_obj.strftime('%Y-%m-%d %H:%M:%S')
    except:
        return str(date_obj)


def read_email_item(item):
    """Read and format email item data."""
    data = {
        'type': 'email',
        'subject': safe_text(item.Subject),
        'sender': safe_text(item.SenderName),
        'sender_email': safe_text(getattr(item, 'SenderEmailAddress', '')),
        'to': safe_text(item.To),
        'cc': safe_text(getattr(item, 'CC', '')),
        'bcc': safe_text(getattr(item, 'BCC', '')),
        'received': format_date(item.ReceivedTime),
        'sent': format_date(getattr(item, 'SentOn', item.ReceivedTime)),
        'body': safe_text(item.Body),
        'html_body': getattr(item, 'HTMLBody', ''),
        'unread': getattr(item, 'UnRead', False),
        'importance': getattr(item, 'Importance', 1),
        'attachments': [],
        'categories': safe_text(getattr(item, 'Categories', '')),
        'conversation_topic': safe_text(getattr(item, 'ConversationTopic', '')),
        'flag_status': getattr(item, 'FlagStatus', 0)
    }
    
    # Get attachment info
    for i, attachment in enumerate(item.Attachments, 1):
        try:
            data['attachments'].append({
                'filename': safe_text(attachment.FileName),
                'size': getattr(attachment, 'Size', 0),
                'type': safe_text(getattr(attachment, 'Type', 'Unknown')),
                'index': i
            })
        except:
            pass
    
    return data


def read_calendar_item(item):
    """Read and format calendar item data."""
    return {
        'type': 'calendar',
        'subject': safe_text(item.Subject),
        'start': format_date(item.Start),
        'end': format_date(item.End),
        'location': safe_text(getattr(item, 'Location', '')),
        'organizer': safe_text(getattr(item, 'Organizer', '')),
        'required_attendees': safe_text(getattr(item, 'RequiredAttendees', '')),
        'optional_attendees': safe_text(getattr(item, 'OptionalAttendees', '')),
        'body': safe_text(item.Body),
        'categories': safe_text(getattr(item, 'Categories', '')),
        'is_recurring': getattr(item, 'IsRecurring', False),
        'all_day_event': getattr(item, 'AllDayEvent', False),
        'busy_status': getattr(item, 'BusyStatus', 0),  # 0=Free, 1=Tentative, 2=Busy, 3=OOF
        'reminder_set': getattr(item, 'ReminderSet', False),
        'reminder_minutes': getattr(item, 'ReminderMinutesBeforeStart', 0),
        'response_status': getattr(item, 'ResponseStatus', 0),  # 0=None, 1=Organized, 2=Tentative, 3=Accepted, 4=Declined
        'importance': getattr(item, 'Importance', 1),
        'attachments': []
    }


def read_contact_item(item):
    """Read and format contact item data."""
    return {
        'type': 'contact',
        'full_name': safe_text(item.FullName),
        'first_name': safe_text(getattr(item, 'FirstName', '')),
        'last_name': safe_text(getattr(item, 'LastName', '')),
        'company': safe_text(getattr(item, 'CompanyName', '')),
        'job_title': safe_text(getattr(item, 'JobTitle', '')),
        'email1': safe_text(getattr(item, 'Email1Address', '')),
        'email2': safe_text(getattr(item, 'Email2Address', '')),
        'email3': safe_text(getattr(item, 'Email3Address', '')),
        'business_phone': safe_text(getattr(item, 'BusinessTelephoneNumber', '')),
        'home_phone': safe_text(getattr(item, 'HomeTelephoneNumber', '')),
        'mobile_phone': safe_text(getattr(item, 'MobileTelephoneNumber', '')),
        'business_address': safe_text(getattr(item, 'BusinessAddress', '')),
        'home_address': safe_text(getattr(item, 'HomeAddress', '')),
        'categories': safe_text(getattr(item, 'Categories', '')),
        'notes': safe_text(getattr(item, 'Body', '')),
        'birthday': format_date(getattr(item, 'Birthday', None)),
        'anniversary': format_date(getattr(item, 'Anniversary', None))
    }


def read_task_item(item):
    """Read and format task item data."""
    status_map = {0: "Not started", 1: "In progress", 2: "Complete", 3: "Waiting", 4: "Deferred"}
    return {
        'type': 'task',
        'subject': safe_text(item.Subject),
        'body': safe_text(item.Body),
        'status': status_map.get(getattr(item, 'Status', 0), "Unknown"),
        'percent_complete': getattr(item, 'PercentComplete', 0),
        'start_date': format_date(getattr(item, 'StartDate', None)),
        'due_date': format_date(getattr(item, 'DueDate', None)),
        'date_completed': format_date(getattr(item, 'DateCompleted', None)),
        'importance': getattr(item, 'Importance', 1),
        'categories': safe_text(getattr(item, 'Categories', '')),
        'reminder_set': getattr(item, 'ReminderSet', False),
        'reminder_time': format_date(getattr(item, 'ReminderTime', None)),
        'owner': safe_text(getattr(item, 'Owner', '')),
        'actual_work': getattr(item, 'ActualWork', 0),
        'total_work': getattr(item, 'TotalWork', 0)
    }


def read_note_item(item):
    """Read and format note item data."""
    return {
        'type': 'note',
        'subject': safe_text(getattr(item, 'Subject', '')),
        'body': safe_text(item.Body),
        'categories': safe_text(getattr(item, 'Categories', '')),
        'created': format_date(getattr(item, 'CreationTime', None)),
        'modified': format_date(getattr(item, 'LastModificationTime', None))
    }


def read_outlook_item(entry_id, save_attachments=False, output_format='text'):
    """
    Read full Outlook item content using its EntryID.
    
    Args:
        entry_id: The EntryID of the item to read (hex or base64)
        save_attachments: Whether to save attachments (for emails/calendar)
        output_format: 'text' or 'json'
    
    Returns:
        Item data dictionary or None if not found
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Decode EntryID if it's base64 encoded
        hex_id = decode_entry_id(entry_id)
        
        # Get item by EntryID
        try:
            item = namespace.GetItemFromID(hex_id)
        except Exception as e:
            # Show truncated version of the hex ID in error messages
            display_id = hex_id[:40] if hex_id else entry_id[:40]
            if output_format == 'json':
                print(json.dumps({'error': f'Could not find item with ID: {display_id}...'}))
            else:
                print(f"Error: Could not find item with ID: {display_id}...")
                print(f"Details: {e}")
            return None
        
        # Determine item type and read data
        item_class = item.Class
        
        # Map Outlook item classes to readers
        if item_class == 43:  # olMail
            item_data = read_email_item(item)
        elif item_class == 26:  # olAppointment
            item_data = read_calendar_item(item)
        elif item_class == 40:  # olContact
            item_data = read_contact_item(item)
        elif item_class == 48:  # olTask
            item_data = read_task_item(item)
        elif item_class == 44:  # olNote
            item_data = read_note_item(item)
        else:
            # Generic handler for unknown types
            item_data = {
                'type': 'unknown',
                'class': item_class,
                'subject': safe_text(getattr(item, 'Subject', '')),
                'body': safe_text(getattr(item, 'Body', '')),
                'categories': safe_text(getattr(item, 'Categories', ''))
            }
        
        # Handle attachments if applicable and requested
        if save_attachments and hasattr(item, 'Attachments') and item.Attachments.Count > 0:
            temp_dir = os.path.join(tempfile.gettempdir(), "outlook_attachments")
            os.makedirs(temp_dir, exist_ok=True)
            
            for att in item_data.get('attachments', []):
                if att['filename']:
                    try:
                        attachment = item.Attachments.Item(att['index'])
                        base_name = att['filename']
                        saved_path = os.path.join(temp_dir, base_name)
                        counter = 1
                        while os.path.exists(saved_path):
                            name_parts = os.path.splitext(base_name)
                            saved_path = os.path.join(temp_dir, f"{name_parts[0]}_{counter}{name_parts[1]}")
                            counter += 1
                        attachment.SaveAsFile(saved_path)
                        att['saved_path'] = saved_path
                    except:
                        pass
        
        # Output based on format
        if output_format == 'json':
            print(json.dumps(item_data, indent=2, default=str))
        else:
            display_item(item_data)
        
        return item_data
        
    except Exception as e:
        if output_format == 'json':
            print(json.dumps({'error': f'Error accessing Outlook: {str(e)}'}))
        else:
            print(f"Error accessing Outlook: {e}")
            print("\nMake sure Outlook is installed and running.")
        return None


def display_item(item_data):
    """Display item in human-readable format."""
    item_type = item_data.get('type', 'unknown')
    
    print("=" * 80)
    
    if item_type == 'email':
        print(f"Type: Email")
        print(f"Subject: {item_data['subject']}")
        print(f"From: {item_data['sender']} <{item_data['sender_email']}>")
        print(f"To: {item_data['to']}")
        if item_data['cc']:
            print(f"CC: {item_data['cc']}")
        print(f"Date: {item_data['received']}")
        
        if item_data['categories']:
            print(f"Categories: {item_data['categories']}")
        
        if item_data['unread']:
            print("Status: UNREAD")
        
        importance_labels = {0: "Low", 1: "Normal", 2: "High"}
        if item_data['importance'] != 1:
            print(f"Importance: {importance_labels.get(item_data['importance'], 'Unknown')}")
        
        if item_data.get('attachments'):
            print(f"\nAttachments ({len(item_data['attachments'])}):")
            for att in item_data['attachments']:
                size_kb = att['size'] / 1024
                print(f"  - {att['filename']} ({size_kb:.1f} KB)")
                if att.get('saved_path'):
                    print(f"    Saved to: {att['saved_path']}")
        
        print("=" * 80)
        print("\nMessage Content:")
        print("-" * 80)
        print(item_data['body'])
        
    elif item_type == 'calendar':
        print(f"Type: Calendar Event")
        print(f"Subject: {item_data['subject']}")
        print(f"Start: {item_data['start']}")
        print(f"End: {item_data['end']}")
        if item_data['location']:
            print(f"Location: {item_data['location']}")
        if item_data['organizer']:
            print(f"Organizer: {item_data['organizer']}")
        if item_data['required_attendees']:
            print(f"Required Attendees: {item_data['required_attendees']}")
        if item_data['optional_attendees']:
            print(f"Optional Attendees: {item_data['optional_attendees']}")
        
        busy_labels = {0: "Free", 1: "Tentative", 2: "Busy", 3: "Out of Office"}
        print(f"Show as: {busy_labels.get(item_data['busy_status'], 'Unknown')}")
        
        if item_data['all_day_event']:
            print("All Day Event: Yes")
        
        if item_data['is_recurring']:
            print("Recurring: Yes")
        
        if item_data['reminder_set']:
            print(f"Reminder: {item_data['reminder_minutes']} minutes before")
        
        if item_data['categories']:
            print(f"Categories: {item_data['categories']}")
        
        print("=" * 80)
        print("\nDescription:")
        print("-" * 80)
        print(item_data['body'])
        
    elif item_type == 'contact':
        print(f"Type: Contact")
        print(f"Name: {item_data['full_name']}")
        if item_data['company']:
            print(f"Company: {item_data['company']}")
        if item_data['job_title']:
            print(f"Title: {item_data['job_title']}")
        
        print("\nContact Information:")
        if item_data['email1']:
            print(f"  Email 1: {item_data['email1']}")
        if item_data['email2']:
            print(f"  Email 2: {item_data['email2']}")
        if item_data['business_phone']:
            print(f"  Business: {item_data['business_phone']}")
        if item_data['mobile_phone']:
            print(f"  Mobile: {item_data['mobile_phone']}")
        if item_data['home_phone']:
            print(f"  Home: {item_data['home_phone']}")
        
        if item_data['business_address']:
            print(f"\nBusiness Address: {item_data['business_address']}")
        if item_data['home_address']:
            print(f"Home Address: {item_data['home_address']}")
        
        if item_data['birthday']:
            print(f"\nBirthday: {item_data['birthday']}")
        
        if item_data['notes']:
            print("\nNotes:")
            print("-" * 40)
            print(item_data['notes'])
            
    elif item_type == 'task':
        print(f"Type: Task")
        print(f"Subject: {item_data['subject']}")
        print(f"Status: {item_data['status']} ({item_data['percent_complete']}% complete)")
        
        if item_data['start_date']:
            print(f"Start Date: {item_data['start_date']}")
        if item_data['due_date']:
            print(f"Due Date: {item_data['due_date']}")
        if item_data['date_completed']:
            print(f"Completed: {item_data['date_completed']}")
        
        importance_labels = {0: "Low", 1: "Normal", 2: "High"}
        if item_data['importance'] != 1:
            print(f"Importance: {importance_labels.get(item_data['importance'], 'Unknown')}")
        
        if item_data['owner']:
            print(f"Owner: {item_data['owner']}")
        
        if item_data['reminder_set']:
            print(f"Reminder: {item_data['reminder_time']}")
        
        if item_data['categories']:
            print(f"Categories: {item_data['categories']}")
        
        print("\nDescription:")
        print("-" * 40)
        print(item_data['body'])
        
    elif item_type == 'note':
        print(f"Type: Note")
        if item_data['subject']:
            print(f"Subject: {item_data['subject']}")
        print(f"Created: {item_data['created']}")
        print(f"Modified: {item_data['modified']}")
        if item_data['categories']:
            print(f"Categories: {item_data['categories']}")
        print("\nContent:")
        print("-" * 40)
        print(item_data['body'])
        
    else:
        print(f"Type: Unknown (Class {item_data.get('class', 'N/A')})")
        if item_data.get('subject'):
            print(f"Subject: {item_data['subject']}")
        if item_data.get('body'):
            print("\nContent:")
            print("-" * 40)
            print(item_data['body'])
    
    print("=" * 80)


def main():
    parser = argparse.ArgumentParser(
        description='Read full Outlook item content by EntryID (email, calendar, contact, task, etc.)',
        epilog='Get the EntryID from outlook_list, outlook_filter, or outlook_search output'
    )
    parser.add_argument('entry_id', 
                        help='The EntryID of the item to read')
    parser.add_argument('--json', action='store_true',
                        help='Output as JSON instead of formatted text')
    parser.add_argument('--save-html', type=str, metavar='FILE',
                        help='Save HTML version to file (for emails)')
    parser.add_argument('--save-text', type=str, metavar='FILE',
                        help='Save text version to file')
    parser.add_argument('--save-attachments', action='store_true',
                        help='Save all attachments to temp folder (for emails/calendar)')
    
    args = parser.parse_args()
    
    # Read the item
    output_format = 'json' if args.json else 'text'
    item_data = read_outlook_item(args.entry_id, 
                                  save_attachments=args.save_attachments,
                                  output_format=output_format)
    
    if not item_data:
        return 1
    
    # Save HTML if requested (for emails)
    if args.save_html and item_data.get('type') == 'email' and item_data.get('html_body'):
        try:
            with open(args.save_html, 'w', encoding='utf-8') as f:
                f.write(item_data['html_body'])
            if not args.json:
                print(f"\nHTML version saved to: {args.save_html}")
        except Exception as e:
            if not args.json:
                print(f"\nError saving HTML: {e}")
    
    # Save text if requested
    if args.save_text:
        try:
            with open(args.save_text, 'w', encoding='utf-8') as f:
                if item_data.get('type') == 'email':
                    f.write(f"Subject: {item_data['subject']}\n")
                    f.write(f"From: {item_data['sender']} <{item_data['sender_email']}>\n")
                    f.write(f"Date: {item_data['received']}\n")
                    f.write("-" * 80 + "\n")
                    f.write(item_data['body'])
                else:
                    # Generic text export
                    f.write(json.dumps(item_data, indent=2, default=str))
            if not args.json:
                print(f"\nText version saved to: {args.save_text}")
        except Exception as e:
            if not args.json:
                print(f"\nError saving text: {e}")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())