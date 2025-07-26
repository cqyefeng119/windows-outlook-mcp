# Outlook MCP Server

This is an MCP (Model Context Protocol) server for integration with Microsoft Outlook. It operates the local Outlook client on Windows via COM, providing email reading, summarization, and draft generation features. Its main advantage is fast deployment on Windows without complex security authentication.

<a href="https://glama.ai/mcp/servers/@cqyefeng119/windows-outlook-mcp">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/@cqyefeng119/windows-outlook-mcp/badge" alt="Outlook Server MCP server" />
</a>

## Features

### 📧 Email Management
- Retrieve inbox email list
- Get details of a specific email by ID
- Search emails
- Mark emails as read

### 📝 Email Summarization
- Intelligently summarize a single email
- Batch summarize inbox emails
- Automatically detect email priority
- Identify actionable emails

### ✍️ Draft Generation
- Create custom email drafts
- Generate drafts using predefined templates
- Intelligently generate reply drafts
- Generate emails based on context

## Installation

### 1. Install dependencies
```powershell
cd path\to\windows-outlook-mcp
npm install
```

### 2. Compile TypeScript
```powershell
npm run build
```

### 3. Configure Claude Desktop
Add the following to your Claude Desktop configuration file:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["path/to/windows-outlook-mcp/src/index.ts"],
      "env": {}
    }
  }
}
```

## Available Tools

### Email Reading Tools

#### `get_inbox_emails`
Retrieve a list of inbox emails
- `count` (optional): Number of emails to retrieve, default is 10

#### `get_email_by_id`
Get a specific email by ID
- `id` (required): Email ID

#### `search_emails`
Search emails
- `query` (required): Search keyword
- `count` (optional): Number of results to return, default is 10

### Email Summarization Tools

#### `summarize_email`
Summarize a single email
- `email_id` (required): ID of the email to summarize

#### `summarize_inbox`
Summarize inbox emails
- `count` (optional): Number of emails to summarize, default is 10

### Draft Generation Tools

#### `create_draft`
Create an email draft
- `to` (required): Array of recipient email addresses
- `cc` (optional): Array of CC email addresses
- `bcc` (optional): Array of BCC email addresses
- `subject` (required): Email subject
- `body` (required): Email content

#### `generate_draft_from_template`
Generate a draft using a template
- `template_name` (required): Template name
  - `meeting_request`: Meeting invitation
  - `follow_up`: Follow-up email
  - `thank_you`: Thank you email
  - `status_update`: Status update
- `variables` (required): Key-value pairs for template variables
- `recipients` (required): Array of recipient email addresses

#### `generate_reply_draft`
Generate a reply draft
- `original_email_id` (required): Original email ID
- `reply_type` (required): Type of reply
  - `agree`: Agree
  - `decline`: Decline
  - `info_request`: Request for information
  - `custom`: Custom
- `custom_message` (optional): Custom message content

#### `generate_smart_draft`
Intelligently generate a draft
- `context` (required): Email context content
- `intent` (required): Email intent
  - `schedule_meeting`: Schedule a meeting
  - `request_information`: Request information
  - `project_update`: Project update
  - `general`: General email
- `recipients` (required): Array of recipient email addresses

### Auxiliary Tools

#### `get_draft_templates`
Get the list of available templates

#### `mark_email_as_read`
Mark an email as read
- `email_id` (required): Email ID

## Usage Examples

### Retrieve and summarize latest emails
```
Please fetch and summarize the latest 5 emails.
```

### Search for specific emails
```
Search for emails containing the keyword "meeting".
```

### Generate a meeting invitation draft
```
Generate an email using the meeting invitation template with the subject "Project Kickoff Meeting" for tomorrow at 2 PM.
```

### Smart reply to an email
```
Generate an "agree" reply draft for email ID: xxx.
```

## System Requirements

- Windows 10/11
- Microsoft Outlook installed and configured
- Node.js 16.0 or higher
- PowerShell 5.0 or higher

## Notes

1. **Permissions**: This tool requires access to Outlook's COM interface. Make sure Outlook is running.
2. **Security**: The tool uses PowerShell scripts to interact with Outlook. Ensure your system security policy allows this.
3. **Performance**: Processing a large number of emails may take time. Batch processing is recommended.
4. **Error Handling**: If you encounter COM errors, restart Outlook and try again.

## Development Notes

Project structure:
```
outlook/
├── src/
│   ├── index.ts              # Main server file
│   ├── outlook-manager.ts    # Outlook interface manager
│   ├── email-summarizer.ts   # Email summarization functionality
│   └── draft-generator.ts    # Draft generation functionality
├── dist/                     # Compiled output directory
├── package.json
├── tsconfig.json
└── README.md
```

To extend functionality, modify the relevant TypeScript files and recompile.