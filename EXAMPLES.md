# Example Prompts for Outlook MCP

This document provides specific working prompts that demonstrate the core functionality of the Outlook MCP server. These examples are designed for MCP Directory reviewers and new users to understand the tool's capabilities.

## Required Setup

Before trying these examples, ensure:
- Windows OS with Microsoft Outlook installed and configured
- Outlook opened at least once to initialize profiles
- outlook-mcp server installed and configured with Claude

## Example 1: Email Management and Search

### Prompt:
```
Use outlook_list to show me my Outlook accounts and then search my Inbox for any emails containing "travel" or "flight" from the last 30 days. If you find any travel-related emails, show me the full content of the most recent one.
```

### Expected Behavior:
1. Lists available Outlook accounts and folder structure
2. Searches Inbox for travel-related keywords using date filtering
3. Displays search results with sender, subject, and date
4. Retrieves and shows full content of most recent match including body text

### Value Demonstration:
- **Account Discovery**: Shows how users can explore their Outlook structure
- **Semantic Search**: Demonstrates intelligent keyword searching across email content
- **Content Retrieval**: Shows full email content access including formatted text

## Example 2: Calendar Integration and Planning

### Prompt:
```
Check my Outlook calendar for any meetings or events scheduled for next week. If there are any travel-related appointments, cross-reference them with emails in my Travel folder to find related booking confirmations or itineraries.
```

### Expected Behavior:
1. Accesses calendar and filters for next week's events
2. Identifies travel-related calendar entries by title/location
3. Searches designated Travel folder for related confirmation emails
4. Correlates calendar events with supporting email documentation

### Value Demonstration:
- **Calendar Access**: Shows integration with Outlook calendar data
- **Cross-Reference Capability**: Demonstrates intelligent connection between calendar and email
- **Travel Use Case**: Highlights practical business travel management scenario

## Example 3: Email Organization and Filtering

### Prompt:
```
Find all unread emails in my Inbox from the last week, group them by sender, and prioritize any that might be urgent (containing words like "urgent", "asap", "deadline", or "immediate"). Show me a summary of what needs my attention.
```

### Expected Behavior:
1. Filters Inbox for unread emails from past 7 days
2. Groups results by sender email/name
3. Applies urgency keyword detection to prioritize messages
4. Provides organized summary with actionable insights

### Value Demonstration:
- **Advanced Filtering**: Shows complex multi-criteria filtering capabilities
- **Intelligent Prioritization**: Demonstrates content analysis for urgency detection
- **Executive Summary**: Highlights ability to synthesize information for decision-making

## Example 4: Attachment and Document Discovery

### Prompt:
```
Search my entire mailbox for emails with PDF attachments that contain "invoice" or "receipt" in the subject line from the last 3 months. List the attachments and their sizes, and help me identify any that might be duplicate receipts based on similar amounts or dates in the email content.
```

### Expected Behavior:
1. Searches across all folders for emails with PDF attachments
2. Filters by subject line keywords and date range
3. Lists attachment details including file names and sizes
4. Analyzes email content to identify potential duplicates

### Value Demonstration:
- **Attachment Management**: Shows comprehensive file attachment discovery
- **Cross-Folder Search**: Demonstrates search across entire mailbox structure
- **Content Analysis**: Highlights intelligent duplicate detection capabilities
- **Business Use Case**: Addresses common expense management workflow

## Example 5: Contact and Communication History

### Prompt:
```
Find all email conversations with john.smith@company.com from the last 6 months, including both sent and received messages. Organize them chronologically and summarize the main topics discussed in our recent communications.
```

### Expected Behavior:
1. Searches both Inbox and Sent Items for specific email address
2. Retrieves all messages in conversation threads
3. Sorts results chronologically to show communication timeline
4. Analyzes content to identify recurring topics and themes

### Value Demonstration:
- **Contact-Centric Search**: Shows person-focused email discovery
- **Bidirectional Communication**: Demonstrates sent/received message correlation
- **Conversation Threading**: Highlights ability to track ongoing discussions
- **Content Summarization**: Shows intelligent topic extraction from communications

## Advanced Usage Patterns

### Complex Query Example:
```
I need to prepare for my quarterly review. Find all emails from my manager Sarah from the last 3 months that mention "project", "deadline", or "performance". Also check my calendar for any 1-on-1 meetings we've had, and cross-reference any action items mentioned in those meeting invitations with follow-up emails.
```

### Integration Example:
```
I'm planning a business trip to New York. Search for any previous travel to NYC in my emails to find hotel recommendations, then check my calendar to see what dates I'm available in the next 4 weeks for a 3-day trip.
```

## Testing Notes for Reviewers

### Sample Data Requirements:
- **Emails**: At least 50+ emails across multiple folders (Inbox, Sent, custom folders)
- **Calendar**: Several calendar events including meetings, appointments, travel
- **Attachments**: Mix of document types (PDF, Word, Excel) for attachment testing
- **Contacts**: Varied sender/recipient addresses for contact-based searches

### Performance Expectations:
- **Simple queries** (< 100 results): Under 3 seconds
- **Complex searches** (> 1000 items): Under 10 seconds  
- **Cross-folder operations**: Under 15 seconds
- **Attachment processing**: Under 5 seconds per file

### Error Handling:
- Graceful handling of missing folders or profiles
- Clear error messages for COM automation issues
- Fallback behavior for Unicode/encoding problems

---

**Note**: These examples assume a typical business Outlook setup. Results will vary based on your specific email volume, folder structure, and Outlook configuration. The MCP server adapts to your existing Outlook organization without requiring any changes to your current setup.