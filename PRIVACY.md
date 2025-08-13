# Privacy Policy

**Last Updated:** January 2025

## Overview

Outlook MCP Server ("the Software") is a local tool that provides Microsoft Outlook integration for AI assistants via the Model Context Protocol. This privacy policy explains how the Software handles your data.

## Data Collection and Usage

### What Data is Accessed
The Software accesses your local Microsoft Outlook data through Windows COM automation, including:
- Email messages (subject, sender, recipient, body, attachments)
- Calendar events (title, location, attendees, description)
- Contacts (name, email, phone, address)
- Tasks and notes

### How Data is Processed
- **Local Only**: All data processing occurs entirely on your local machine
- **No External Transmission**: No Outlook data is sent to external servers, APIs, or third parties
- **No Storage**: The Software does not store, cache, or retain any of your Outlook data
- **No Analytics**: No usage statistics or telemetry data is collected
- **On-Demand Access**: Data is only accessed when you explicitly request it through AI assistant commands

### Data Flow
1. You make a request through your AI assistant (e.g., "show me recent emails")
2. The AI assistant calls the appropriate MCP tool
3. The Software accesses your local Outlook application via COM automation
4. Data is retrieved and returned directly to your AI assistant
5. No data is stored or transmitted elsewhere

## Technical Implementation

### COM Automation Security
- Uses Windows COM (Component Object Model) to interface with Outlook
- Requires your existing Outlook installation and profile
- Operates with the same permissions as your Outlook application
- No additional authentication or credentials required

### Data Handling Safeguards
- Text encoding is handled safely to prevent Unicode errors
- Emoji and special characters are processed securely
- No data parsing that could expose sensitive information
- Error handling prevents data leakage in exception messages

## Your Rights and Control

### You Maintain Full Control
- **Your Data Stays Local**: All Outlook data remains on your device
- **Uninstall Anytime**: Remove the Software without data retention concerns  
- **No Account Required**: No registration or account creation needed
- **Outlook Controls Access**: Your existing Outlook security settings apply

### Permissions
- The Software requests no additional permissions beyond COM access to Outlook
- Operates within the security context of your user account
- Cannot access data from other users or applications

## Third-Party Services

### No External Dependencies
- Does not connect to external APIs or services
- Does not use cloud-based processing
- Does not integrate with external data providers
- Only dependencies are local system components (Node.js, Python, Windows COM)

## Updates and Changes

### Privacy Policy Updates
- Changes to this policy will be reflected in the Software's repository
- Material changes will be highlighted in release notes
- Continued use indicates acceptance of updated terms

## Contact Information

### Support and Questions
- **Issues**: Report via [GitHub Issues](https://github.com/cjmlarson/outlook-mcp/issues)
- **Author**: Connor Larson ([@cjmlarson](https://github.com/cjmlarson))
- **Repository**: https://github.com/cjmlarson/outlook-mcp

## Compliance

### Data Protection Principles
This Software is designed with privacy-by-design principles:
- **Data Minimization**: Only accesses data necessary for requested operations
- **Purpose Limitation**: Data is used only for AI assistant integration
- **Local Processing**: No data leaves your device
- **User Control**: You control all data access through your requests

### Open Source Transparency
- All source code is publicly available for review
- No hidden data collection or transmission
- Community oversight and contribution welcome

---

**Note**: This Software accesses your local Outlook data only. It does not collect personal information about you beyond what's necessary to operate with your Outlook installation. Your privacy and data security are paramount in the design and operation of this tool.