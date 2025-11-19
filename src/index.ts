#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { OutlookManager, EmailMessage, EmailDraft } from './outlook-manager.js';

const server = new Server(
  {
    name: 'outlook-mcp-server',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

const outlookManager = new OutlookManager();

// List available tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: "get_inbox_emails",
        description: "Get inbox email list",
        inputSchema: {
          type: "object",
          properties: {
            count: {
              type: "number",
              description: "Number of emails to retrieve",
              default: 10
            }
          }
        }
      },
      {
        name: "get_email_by_id", 
        description: "Get specific email by ID",
        inputSchema: {
          type: "object",
          properties: {
            id: {
              type: "string",
              description: "Email ID"
            }
          },
          required: ["id"]
        }
      },
      {
        name: "summarize_email",
        description: "Summarize individual email content",
        inputSchema: {
          type: "object",
          properties: {
            email_id: {
              type: "string",
              description: "Email ID to summarize"
            }
          },
          required: ["email_id"]
        }
      },
      {
        name: "summarize_inbox",
        description: "Summarize inbox emails",
        inputSchema: {
          type: "object",
          properties: {
            count: {
              type: "number",
              description: "Number of emails to summarize",
              default: 10
            }
          }
        }
      },
      {
        name: "create_draft",
        description: "Create new email draft",
        inputSchema: {
          type: "object",
          properties: {
            to: {
              type: "array",
              items: { type: "string" },
              description: "Recipient email addresses"
            },
            subject: {
              type: "string",
              description: "Email subject"
            },
            body: {
              type: "string",
              description: "Email content"
            },
            cc: {
              type: "array",
              items: { type: "string" },
              description: "CC email addresses"
            },
            bcc: {
              type: "array",
              items: { type: "string" },
              description: "BCC email addresses"
            }
          },
          required: ["to", "subject", "body"]
        }
      },
      {
        name: "mark_email_as_read",
        description: "Mark email as read",
        inputSchema: {
          type: "object",
          properties: {
            email_id: {
              type: "string",
              description: "Email ID"
            }
          },
          required: ["email_id"]
        }
      },
      {
        name: "search_inbox_emails",
        description: "Search inbox emails",
        inputSchema: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "Search keywords"
            },
            count: {
              type: "number",
              description: "Number of results to return",
              default: 10
            }
          },
          required: ["query"]
        }
      },
      {
        name: "search_sent_emails",
        description: "Search sent emails",
        inputSchema: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "Search keywords"
            },
            count: {
              type: "number",
              description: "Number of results to return",
              default: 10
            }
          },
          required: ["query"]
        }
      },
      {
        name: "get_sent_emails",
        description: "Get sent emails list",
        inputSchema: {
          type: "object",
          properties: {
            count: {
              type: "number",
              description: "Number of sent emails to retrieve",
              default: 10
            }
          }
        }
      },
      {
        name: "get_draft_emails",
        description: "Get draft emails list",
        inputSchema: {
          type: "object",
          properties: {
            count: {
              type: "number",
              description: "Number of draft emails to retrieve",
              default: 10
            }
          }
        }
      },
      {
        name: "search_draft_emails",
        description: "Search draft emails",
        inputSchema: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "Search keywords"
            },
            count: {
              type: "number",
              description: "Number of results to return",
              default: 10
            }
          },
          required: ["query"]
        }
      },
      {
        name: "duplicate_email_as_draft",
        description: "Duplicate existing email as draft (preserving complete format)",
        inputSchema: {
          type: "object",
          properties: {
            source_email_id: {
              type: "string",
              description: "Original email ID to duplicate"
            },
            store_id: {
              type: "string",
              description: "Store ID for the original email (optional but recommended)"
            },
            new_subject: {
              type: "string",
              description: "New email subject (optional)"
            },
            new_recipients: {
              type: "array",
              items: { type: "string" },
              description: "New recipient list (optional)"
            }
          },
          required: ["source_email_id"]
        }
      },
      {
        name: "set_show_as",
        description: "Set Show As (Free/Busy status) for a calendar event",
        inputSchema: {
          type: "object",
          properties: {
            eventId: {
              type: "string",
              description: "Event ID to update"
            },
            subject: {
              type: "string",
              description: "Subject of the event to find"
            },
            startDate: {
              type: "string",
              description: "Start date of the event to find (ISO 8601 format)"
            },
            showAs: {
              type: "string",
              enum: ["Free", "Tentative", "Busy", "OutOfOffice", "WorkingElsewhere"],
              description: "Show As status to set"
            }
          },
          required: ["showAs"]
        }
      },
      {
        name: "create_event_with_show_as",
        description: "Create a calendar event with specific Show As status (e.g., OutOfOffice for vacation)",
        inputSchema: {
          type: "object",
          properties: {
            subject: {
              type: "string",
              description: "Event subject/title"
            },
            start: {
              type: "string",
              description: "Start date and time (ISO 8601 format)"
            },
            end: {
              type: "string",
              description: "End date and time (ISO 8601 format)"
            },
            location: {
              type: "string",
              description: "Event location"
            },
            body: {
              type: "string",
              description: "Event description/body"
            },
            showAs: {
              type: "string",
              enum: ["Free", "Tentative", "Busy", "OutOfOffice", "WorkingElsewhere"],
              description: "Show As status (default: Busy)"
            },
            reminderMinutes: {
              type: "number",
              description: "Reminder time in minutes before the event"
            }
          },
          required: ["subject", "start", "end"]
        }
      },
      {
        name: "list_events",
        description: "List calendar events within a specified date range",
        inputSchema: {
          type: "object",
          properties: {
            startDate: {
              type: "string",
              description: "Start date (ISO 8601 format)"
            },
            endDate: {
              type: "string",
              description: "End date (ISO 8601 format, optional)"
            },
            calendar: {
              type: "string",
              description: "Calendar name (optional)"
            }
          },
          required: ["startDate"]
        }
      },
      {
        name: "update_event",
        description: "Update an existing calendar event",
        inputSchema: {
          type: "object",
          properties: {
            eventId: {
              type: "string",
              description: "Event ID to update"
            },
            subject: {
              type: "string",
              description: "New event subject/title (optional)"
            },
            startDate: {
              type: "string",
              description: "New start date in MM/DD/YYYY format (optional)"
            },
            startTime: {
              type: "string",
              description: "New start time in HH:MM AM/PM format (optional)"
            },
            endDate: {
              type: "string",
              description: "New end date in MM/DD/YYYY format (optional)"
            },
            endTime: {
              type: "string",
              description: "New end time in HH:MM AM/PM format (optional)"
            },
            location: {
              type: "string",
              description: "New event location (optional)"
            },
            body: {
              type: "string",
              description: "New event description/body (optional)"
            },
            calendar: {
              type: "string",
              description: "Calendar name (optional)"
            }
          },
          required: ["eventId"]
        }
      },
      {
        name: "delete_event",
        description: "Delete a calendar event by its ID",
        inputSchema: {
          type: "object",
          properties: {
            eventId: {
              type: "string",
              description: "Event ID to delete"
            },
            calendar: {
              type: "string",
              description: "Calendar name (optional)"
            }
          },
          required: ["eventId"]
        }
      },
      {
        name: "find_free_slots",
        description: "Find available time slots in the calendar",
        inputSchema: {
          type: "object",
          properties: {
            startDate: {
              type: "string",
              description: "Start date (ISO 8601 format)"
            },
            endDate: {
              type: "string",
              description: "End date (ISO 8601 format, optional, defaults to 7 days from start)"
            },
            duration: {
              type: "number",
              description: "Duration in minutes (optional, defaults to 30)"
            },
            workDayStart: {
              type: "number",
              description: "Work day start hour (0-23) (optional, defaults to 9)"
            },
            workDayEnd: {
              type: "number",
              description: "Work day end hour (0-23) (optional, defaults to 17)"
            },
            calendar: {
              type: "string",
              description: "Calendar name (optional)"
            }
          },
          required: ["startDate"]
        }
      },
      {
        name: "get_attendee_status",
        description: "Check the response status of meeting attendees",
        inputSchema: {
          type: "object",
          properties: {
            eventId: {
              type: "string",
              description: "Event ID"
            },
            calendar: {
              type: "string",
              description: "Calendar name (optional)"
            }
          },
          required: ["eventId"]
        }
      },
      {
        name: "get_calendars",
        description: "List available calendars",
        inputSchema: {
          type: "object",
          properties: {}
        }
      }
    ]
  };
});

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      case 'get_inbox_emails': {
        const count = (args as any)?.count || 10;
        const emails = await outlookManager.getInboxEmails(count);
        return {
          content: [
            {
              type: 'text',
              text: `📊 **Email Overview**\nTotal: ${emails.length} emails\nUnread: ${emails.filter(e => !e.isRead).length} emails\n\n📋 **Email List:**\n` + 
                   emails.map((email, index) => 
                     `${index + 1}. ${email.isRead ? '✅' : '📩'} **${email.subject}**\n   From: ${email.sender}\n   Time: ${email.receivedTime}\n   Preview: ${email.body?.substring(0, 100)}...\n`
                   ).join('\n')
            },
          ],
        };
      }

      case 'get_email_by_id': {
        const id = (args as any)?.id;
        if (!id) {
          throw new Error('Email ID is required');
        }
        const email = await outlookManager.getEmailById(id);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(email, null, 2),
            },
          ],
        };
      }

      case 'summarize_email': {
        const emailId = (args as any)?.email_id;
        if (!emailId) {
          throw new Error('Email ID is required');
        }
        const email = await outlookManager.getEmailById(emailId);
        return {
          content: [
            {
              type: 'text',
              text: `📧 **Email Details**\n\n**Subject:** ${email.subject}\n**From:** ${email.sender}\n**To:** ${email.recipients?.join(', ')}\n**Time:** ${email.receivedTime}\n**Read:** ${email.isRead ? 'Yes' : 'No'}\n\n**Content Summary:**\n${email.body?.substring(0, 500)}${email.body && email.body.length > 500 ? '...' : ''}`,
            },
          ],
        };
      }

      case 'summarize_inbox': {
        const count = (args as any)?.count || 10;
        const emails = await outlookManager.getInboxEmails(count);
        const unreadCount = emails.filter(e => !e.isRead).length;
        
        return {
          content: [
            {
              type: 'text',
              text: `📊 **Inbox Summary**\nTotal: ${emails.length} emails\nUnread: ${unreadCount} emails\n\n📋 **Recent Emails:**\n` +
                   emails.slice(0, 5).map((email, index) => 
                     `${index + 1}. ${email.isRead ? '✅' : '📩'} **${email.subject}**\n   From: ${email.sender}\n   Time: ${email.receivedTime}\n`
                   ).join('\n')
            },
          ],
        };
      }

      case 'create_draft': {
        const draft: EmailDraft = {
          to: (args as any)?.to || [],
          subject: (args as any)?.subject || '',
          body: (args as any)?.body || '',
          cc: (args as any)?.cc,
          bcc: (args as any)?.bcc,
        };
        const result = await outlookManager.createDraft(draft);
        return {
          content: [
            {
              type: 'text',
              text: `✅ **Email Draft Created Successfully**\n\n**Subject:** ${draft.subject}\n**To:** ${draft.to.join(', ')}\n${draft.cc ? `**CC:** ${draft.cc.join(', ')}\n` : ''}**Result:** ${result}`,
            },
          ],
        };
      }

      case 'mark_email_as_read': {
        const emailId = (args as any)?.email_id;
        if (!emailId) {
          throw new Error('Email ID is required');
        }
        await outlookManager.markAsRead(emailId);
        return {
          content: [
            {
              type: 'text',
              text: `✅ **Email marked as read**\nEmail ID: ${emailId}`,
            },
          ],
        };
      }

      case 'search_inbox_emails': {
        const query = (args as any)?.query;
        const count = (args as any)?.count || 10;
        if (!query) {
          throw new Error('Search query is required');
        }
        const emails = await outlookManager.searchInboxEmails(query, count);
        return {
          content: [
            {
              type: 'text',
              text: `🔍 **Search Results: "${query}"**\nTotal: ${emails.length} items\nUnread: ${emails.filter(e => !e.isRead).length} items\n\n📋 **Search Results List:**\n` +
                   emails.map((email, index) => 
                     `${index + 1}. ${email.isRead ? '✅' : '📩'} **${email.subject}**\n   From: ${email.sender}\n   Time: ${email.receivedTime}\n   EntryID: ${email.id}\n   StoreID: ${email.storeId || 'N/A'}\n   Search Context: ${email.body?.includes(query) ? 'Match in content' : 'Match in subject'}: ${email.subject}\n   Preview: ${email.body?.substring(0, 100)}...\n`
                   ).join('\n')
            },
          ],
        };
      }

      case 'search_sent_emails': {
        const query = (args as any)?.query;
        const count = (args as any)?.count || 10;
        if (!query) {
          throw new Error('Search query is required');
        }
        const emails = await outlookManager.searchSentEmails(query, count);
        return {
          content: [
            {
              type: 'text',
              text: `🔍 **Search Results: "${query}"**\nTotal: ${emails.length} items\nUnread: ${emails.filter(e => !e.isRead).length} items\n\n📋 **Search Results List:**\n` +
                   emails.map((email, index) => 
                     `${index + 1}. ${email.isRead ? '✅' : '📩'} **${email.subject}**\n   From: ${email.sender}\n   Time: ${email.receivedTime}\n   EntryID: ${email.id}\n   StoreID: ${email.storeId || 'N/A'}\n   Search Context: ${email.body?.includes(query) ? 'Match in content' : 'Match in subject'}: ${email.subject}\n   Preview: ${email.body?.substring(0, 100)}...\n`
                   ).join('\n')
            },
          ],
        };
      }

      case 'get_sent_emails': {
        const count = (args as any)?.count || 10;
        const emails = await outlookManager.getSentEmails(count);
        return {
          content: [
            {
              type: 'text',
              text: `📊 **Email Overview**\nTotal: ${emails.length} emails\nUnread: ${emails.filter(e => !e.isRead).length} emails\n\n📋 **Email List:**\n` + 
                   emails.map((email, index) => 
                     `${index + 1}. ${email.isRead ? '✅' : '📩'} **${email.subject}**\n   From: ${email.sender}\n   Time: ${email.receivedTime}\n   EntryID: ${email.id}\n   StoreID: ${email.storeId || 'N/A'}\n   Preview: ${email.body?.substring(0, 100)}...\n`
                   ).join('\n')
            },
          ],
        };
      }

      case 'get_draft_emails': {
        const count = (args as any)?.count || 10;
        const drafts = await outlookManager.getDraftEmails(count);
        return {
          content: [
            {
              type: 'text',
              text: `📂 **Draft Email Overview**\nTotal: ${drafts.length} drafts\n\n📝 **Draft Email List:**\n` + 
                   drafts.map((draft, index) => 
                     `${index + 1}. **${draft.subject}**\n   From: ${draft.sender}\n   Time: ${draft.receivedTime}\n   EntryID: ${draft.id}\n   StoreID: ${draft.storeId || 'N/A'}\n   Preview: ${draft.body?.substring(0, 100)}...\n`
                   ).join('\n')
            },
          ],
        };
      }

      case 'search_draft_emails': {
        const query = (args as any)?.query;
        const count = (args as any)?.count || 10;
        if (!query) {
          throw new Error('Search query is required');
        }
        const drafts = await outlookManager.searchDraftEmails(query, count);
        return {
          content: [
            {
              type: 'text',
              text: `🔍 **Draft Search Results: "${query}"**\nTotal: ${drafts.length} items\n\n📋 **Draft Search Results List:**\n` +
                   drafts.map((draft, index) => 
                     `${index + 1}. **${draft.subject}**\n   From: ${draft.sender}\n   Time: ${draft.receivedTime}\n   EntryID: ${draft.id}\n   StoreID: ${draft.storeId || 'N/A'}\n   Search Context: ${draft.body?.includes(query) ? 'Match in content' : 'Match in subject'}: ${draft.subject}\n   Preview: ${draft.body?.substring(0, 100)}...\n`
                   ).join('\n')
            },
          ],
        };
      }

      case 'duplicate_email_as_draft': {
        const sourceEmailId = (args as any)?.source_email_id;
        const newSubject = (args as any)?.new_subject;
        const newRecipients = (args as any)?.new_recipients;
        const storeId = (args as any)?.store_id;
        
        if (!sourceEmailId) {
          throw new Error('Source email ID is required');
        }
        
        const result = await outlookManager.duplicateEmailAsDraft(
          sourceEmailId,
          newSubject,
          newRecipients,
          storeId
        );
        return {
          content: [
            {
              type: 'text',
              text: `✅ **Email duplicated successfully**\n\n**Result:** ${result}${newSubject ? `\n**New Subject:** ${newSubject}` : ''}${newRecipients ? `\n**New Recipients:** ${newRecipients.join(', ')}` : ''}`,
            },
          ],
        };
      }

      case 'set_show_as': {
        const options = {
          eventId: (args as any)?.eventId,
          subject: (args as any)?.subject,
          startDate: (args as any)?.startDate ? new Date((args as any).startDate) : undefined,
          showAs: (args as any)?.showAs
        };
        const result = await outlookManager.setShowAs(options);
        return {
          content: [
            {
              type: 'text',
              text: `${result.success ? '✅' : '❌'} **Show As Updated**\n\n${result.message}`,
            },
          ],
        };
      }

      case 'create_event_with_show_as': {
        const result = await outlookManager.createEventWithShowAs({
          subject: (args as any)?.subject,
          start: new Date((args as any)?.start),
          end: new Date((args as any)?.end),
          location: (args as any)?.location,
          body: (args as any)?.body,
          showAs: (args as any)?.showAs,
          reminderMinutes: (args as any)?.reminderMinutes
        });
        return {
          content: [
            {
              type: 'text',
              text: `${result.success ? '✅' : '❌'} **Calendar Event Created**\n\n${result.message}\n\n**Event ID:** ${result.eventId}`,
            },
          ],
        };
      }

      case 'list_events': {
        const events = await outlookManager.listEvents({
          startDate: new Date((args as any)?.startDate),
          endDate: (args as any)?.endDate ? new Date((args as any)?.endDate) : undefined,
          calendar: (args as any)?.calendar
        });
        return {
          content: [
            {
              type: 'text',
              text: `📅 **Calendar Events**\nTotal: ${events.length} events\n\n` +
                   events.map((event, index) => 
                     `${index + 1}. **${event.Subject}**\n   📍 ${event.Location || 'No location'}\n   🕐 ${event.Start} - ${event.End}\n   📋 ${event.Body ? event.Body.substring(0, 100) + '...' : 'No description'}\n   ID: ${event.Id}\n`
                   ).join('\n')
            },
          ],
        };
      }

      case 'update_event': {
        const result = await outlookManager.updateEvent({
          eventId: (args as any)?.eventId,
          subject: (args as any)?.subject,
          startDate: (args as any)?.startDate,
          startTime: (args as any)?.startTime,
          endDate: (args as any)?.endDate,
          endTime: (args as any)?.endTime,
          location: (args as any)?.location,
          body: (args as any)?.body,
          calendar: (args as any)?.calendar
        });
        return {
          content: [
            {
              type: 'text',
              text: `${result.success ? '✅' : '❌'} **Event Update**\n\n${result.message}`,
            },
          ],
        };
      }

      case 'delete_event': {
        const result = await outlookManager.deleteEvent({
          eventId: (args as any)?.eventId,
          calendar: (args as any)?.calendar
        });
        return {
          content: [
            {
              type: 'text',
              text: `${result.success ? '✅' : '❌'} **Event Deletion**\n\n${result.message}`,
            },
          ],
        };
      }

      case 'find_free_slots': {
        const freeSlots = await outlookManager.findFreeSlots({
          startDate: new Date((args as any)?.startDate),
          endDate: (args as any)?.endDate ? new Date((args as any)?.endDate) : undefined,
          duration: (args as any)?.duration,
          workDayStart: (args as any)?.workDayStart,
          workDayEnd: (args as any)?.workDayEnd,
          calendar: (args as any)?.calendar
        });
        return {
          content: [
            {
              type: 'text',
              text: `🆓 **Free Time Slots**\nTotal: ${freeSlots.length} slots found\n\n` +
                   freeSlots.slice(0, 20).map((slot, index) => 
                     `${index + 1}. ${slot.Start} - ${slot.End}`
                   ).join('\n') +
                   (freeSlots.length > 20 ? `\n\n... and ${freeSlots.length - 20} more slots` : '')
            },
          ],
        };
      }

      case 'get_attendee_status': {
        const attendeeStatus = await outlookManager.getAttendeeStatus({
          eventId: (args as any)?.eventId,
          calendar: (args as any)?.calendar
        });
        return {
          content: [
            {
              type: 'text',
              text: `👥 **Meeting Attendee Status**\n\n**Subject:** ${attendeeStatus.Subject}\n**Time:** ${attendeeStatus.Start} - ${attendeeStatus.End}\n**Location:** ${attendeeStatus.Location || 'N/A'}\n**Organizer:** ${attendeeStatus.Organizer}\n\n**Attendees:**\n` +
                   attendeeStatus.Attendees.map((attendee: any, index: number) => 
                     `${index + 1}. ${attendee.Name} (${attendee.Email})\n   Status: ${attendee.ResponseStatus}`
                   ).join('\n')
            },
          ],
        };
      }

      case 'get_calendars': {
        const calendars = await outlookManager.getCalendars();
        return {
          content: [
            {
              type: 'text',
              text: `📅 **Available Calendars**\nTotal: ${calendars.length} calendars\n\n` +
                   calendars.map((calendar, index) => 
                     `${index + 1}. ${calendar.IsDefault ? '⭐' : '📅'} **${calendar.Name}**\n   Owner: ${calendar.Owner}`
                   ).join('\n')
            },
          ],
        };
      }

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    return {
      content: [
        {
          type: 'text',
          text: `❌ **Error:** ${error instanceof Error ? error.message : String(error)}`,
        },
      ],
      isError: true,
    };
  }
});

async function runServer() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Outlook MCP Server running on stdio');
}

runServer().catch(console.error);
