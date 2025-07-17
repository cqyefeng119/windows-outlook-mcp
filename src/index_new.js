#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema, } from '@modelcontextprotocol/sdk/types.js';
import { OutlookManager } from './outlook-manager.js';
import { EmailSummarizer } from './email-summarizer.js';
import { DraftGenerator } from './draft-generator.js';
const server = new Server({
    name: 'outlook-mcp-server',
    version: '1.0.0',
}, {
    capabilities: {
        tools: {},
    },
});
const outlookManager = new OutlookManager();
// List available tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
        tools: [
            {
                name: "get_inbox_emails",
                description: "取得收件箱邮件列表",
                inputSchema: {
                    type: "object",
                    properties: {
                        count: {
                            type: "number",
                            description: "要获取的邮件数量",
                            default: 10
                        }
                    }
                }
            },
            {
                name: "get_email_by_id",
                description: "根据ID获取特定邮件",
                inputSchema: {
                    type: "object",
                    properties: {
                        id: {
                            type: "string",
                            description: "邮件ID"
                        }
                    },
                    required: ["id"]
                }
            },
            {
                name: "summarize_email",
                description: "总结单个邮件内容",
                inputSchema: {
                    type: "object",
                    properties: {
                        email_id: {
                            type: "string",
                            description: "要总结的邮件ID"
                        }
                    },
                    required: ["email_id"]
                }
            },
            {
                name: "summarize_inbox",
                description: "总结收件箱邮件",
                inputSchema: {
                    type: "object",
                    properties: {
                        count: {
                            type: "number",
                            description: "要总结的邮件数量",
                            default: 10
                        }
                    }
                }
            },
            {
                name: "create_draft",
                description: "创建新的邮件草稿",
                inputSchema: {
                    type: "object",
                    properties: {
                        to: {
                            type: "array",
                            items: { type: "string" },
                            description: "收件人邮箱地址"
                        },
                        subject: {
                            type: "string",
                            description: "邮件主题"
                        },
                        body: {
                            type: "string",
                            description: "邮件内容"
                        },
                        cc: {
                            type: "array",
                            items: { type: "string" },
                            description: "抄送邮箱地址"
                        },
                        bcc: {
                            type: "array",
                            items: { type: "string" },
                            description: "密送邮箱地址"
                        }
                    },
                    required: ["to", "subject", "body"]
                }
            },
            {
                name: "generate_draft_from_template",
                description: "从模板生成邮件草稿",
                inputSchema: {
                    type: "object",
                    properties: {
                        template_name: {
                            type: "string",
                            enum: ["meeting_request", "follow_up", "thank_you", "status_update"],
                            description: "模板名称"
                        },
                        variables: {
                            type: "object",
                            additionalProperties: { type: "string" },
                            description: "模板变量键值对"
                        },
                        recipients: {
                            type: "array",
                            items: { type: "string" },
                            description: "收件人邮箱地址"
                        }
                    },
                    required: ["template_name", "variables", "recipients"]
                }
            },
            {
                name: "generate_reply_draft",
                description: "生成回复邮件草稿",
                inputSchema: {
                    type: "object",
                    properties: {
                        original_email_id: {
                            type: "string",
                            description: "原始邮件ID"
                        },
                        reply_type: {
                            type: "string",
                            enum: ["agree", "decline", "info_request", "custom"],
                            description: "回复类型"
                        },
                        custom_message: {
                            type: "string",
                            description: "自定义消息内容"
                        }
                    },
                    required: ["original_email_id", "reply_type"]
                }
            },
            {
                name: "generate_smart_draft",
                description: "根据上下文智能生成邮件草稿",
                inputSchema: {
                    type: "object",
                    properties: {
                        context: {
                            type: "string",
                            description: "邮件上下文内容"
                        },
                        intent: {
                            type: "string",
                            enum: ["schedule_meeting", "request_information", "project_update", "general"],
                            description: "邮件意图"
                        },
                        recipients: {
                            type: "array",
                            items: { type: "string" },
                            description: "收件人邮箱地址"
                        }
                    },
                    required: ["context", "intent", "recipients"]
                }
            },
            {
                name: "get_draft_templates",
                description: "获取可用的邮件模板列表",
                inputSchema: {
                    type: "object",
                    properties: {}
                }
            },
            {
                name: "mark_email_as_read",
                description: "将邮件标记为已读",
                inputSchema: {
                    type: "object",
                    properties: {
                        email_id: {
                            type: "string",
                            description: "邮件ID"
                        }
                    },
                    required: ["email_id"]
                }
            },
            {
                name: "search_inbox_emails",
                description: "搜索收件箱邮件",
                inputSchema: {
                    type: "object",
                    properties: {
                        query: {
                            type: "string",
                            description: "搜索关键词"
                        },
                        count: {
                            type: "number",
                            description: "返回结果数量",
                            default: 10
                        }
                    },
                    required: ["query"]
                }
            },
            {
                name: "search_sent_emails",
                description: "搜索已发送邮件",
                inputSchema: {
                    type: "object",
                    properties: {
                        query: {
                            type: "string",
                            description: "搜索关键词"
                        },
                        count: {
                            type: "number",
                            description: "返回结果数量",
                            default: 10
                        }
                    },
                    required: ["query"]
                }
            },
            {
                name: "get_sent_emails",
                description: "获取已发送邮件列表",
                inputSchema: {
                    type: "object",
                    properties: {
                        count: {
                            type: "number",
                            description: "要获取的已发送邮件数量",
                            default: 10
                        }
                    }
                }
            },
            {
                name: "duplicate_email_as_draft",
                description: "复制已有邮件为草稿（保持完整格式）",
                inputSchema: {
                    type: "object",
                    properties: {
                        source_email_id: {
                            type: "string",
                            description: "要复制的原始邮件ID"
                        },
                        new_subject: {
                            type: "string",
                            description: "新邮件主题（可选）"
                        },
                        new_recipients: {
                            type: "array",
                            items: { type: "string" },
                            description: "新收件人列表（可选）"
                        }
                    },
                    required: ["source_email_id"]
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
                const emails = await outlookManager.getInboxEmails(args.count || 10);
                return {
                    content: [
                        {
                            type: 'text',
                            text: EmailSummarizer.formatEmailList(emails),
                        },
                    ],
                };
            }
            case 'get_email_by_id': {
                const email = await outlookManager.getEmailById(args.id);
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
                const email = await outlookManager.getEmailById(args.email_id);
                const summary = EmailSummarizer.summarizeEmail(email);
                return {
                    content: [
                        {
                            type: 'text',
                            text: summary,
                        },
                    ],
                };
            }
            case 'summarize_inbox': {
                const emails = await outlookManager.getInboxEmails(args.count || 10);
                const summary = EmailSummarizer.summarizeInbox(emails);
                return {
                    content: [
                        {
                            type: 'text',
                            text: summary,
                        },
                    ],
                };
            }
            case 'create_draft': {
                const draft = {
                    to: args.to,
                    subject: args.subject,
                    body: args.body,
                    cc: args.cc,
                    bcc: args.bcc,
                };
                const result = await outlookManager.createDraft(draft);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `邮件草稿创建成功: ${result}`,
                        },
                    ],
                };
            }
            case 'generate_draft_from_template': {
                const draft = DraftGenerator.generateFromTemplate(args.template_name, args.variables, args.recipients);
                const result = await outlookManager.createDraft(draft);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `模板邮件草稿创建成功: ${result}\n模板: ${args.template_name}`,
                        },
                    ],
                };
            }
            case 'generate_reply_draft': {
                const originalEmail = await outlookManager.getEmailById(args.original_email_id);
                const draft = DraftGenerator.generateReply(originalEmail, args.reply_type, args.custom_message);
                const result = await outlookManager.createDraft(draft);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `回复邮件草稿创建成功: ${result}\n回复类型: ${args.reply_type}`,
                        },
                    ],
                };
            }
            case 'generate_smart_draft': {
                const draft = DraftGenerator.generateSmartDraft(args.context, args.intent, args.recipients);
                const result = await outlookManager.createDraft(draft);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `智能邮件草稿创建成功: ${result}\n意图: ${args.intent}`,
                        },
                    ],
                };
            }
            case 'get_draft_templates': {
                const templates = DraftGenerator.getTemplates();
                return {
                    content: [
                        {
                            type: 'text',
                            text: JSON.stringify(templates, null, 2),
                        },
                    ],
                };
            }
            case 'mark_email_as_read': {
                await outlookManager.markAsRead(args.email_id);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `邮件已标记为已读: ${args.email_id}`,
                        },
                    ],
                };
            }
            case 'search_inbox_emails': {
                const emails = await outlookManager.searchInboxEmails(args.query, args.count || 10);
                return {
                    content: [
                        {
                            type: 'text',
                            text: EmailSummarizer.formatSearchResults(emails, args.query),
                        },
                    ],
                };
            }
            case 'search_sent_emails': {
                const emails = await outlookManager.searchSentEmails(args.query, args.count || 10);
                return {
                    content: [
                        {
                            type: 'text',
                            text: EmailSummarizer.formatSearchResults(emails, args.query),
                        },
                    ],
                };
            }
            case 'get_sent_emails': {
                const emails = await outlookManager.getSentEmails(args.count || 10);
                return {
                    content: [
                        {
                            type: 'text',
                            text: EmailSummarizer.formatEmailList(emails),
                        },
                    ],
                };
            }
            case 'duplicate_email_as_draft': {
                const result = await outlookManager.duplicateEmailAsDraft(args.source_email_id, args.new_subject, args.new_recipients);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `邮件复制成功: ${result}${args.new_subject ? `\n新主题: ${args.new_subject}` : ''}${args.new_recipients ? `\n新收件人: ${args.new_recipients.join(', ')}` : ''}`,
                        },
                    ],
                };
            }
            default:
                throw new Error(`Unknown tool: ${name}`);
        }
    }
    catch (error) {
        return {
            content: [
                {
                    type: 'text',
                    text: `Error: ${error instanceof Error ? error.message : String(error)}`,
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
