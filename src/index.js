#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema, } from '@modelcontextprotocol/sdk/types.js';
import { OutlookManager } from './outlook-manager.js';
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
                description: "获取收件箱邮件列表",
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
                const count = args?.count || 10;
                const emails = await outlookManager.getInboxEmails(count);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `📊 **邮件总览**\n总计: ${emails.length} 封邮件\n未读: ${emails.filter(e => !e.isRead).length} 封\n\n📋 **邮件列表:**\n` +
                                emails.map((email, index) => `${index + 1}. ${email.isRead ? '✅' : '📩'} **${email.subject}**\n   发件人: ${email.sender}\n   时间: ${email.receivedTime}\n   预览: ${email.body?.substring(0, 100)}...\n`).join('\n')
                        },
                    ],
                };
            }
            case 'get_email_by_id': {
                const id = args?.id;
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
                const emailId = args?.email_id;
                if (!emailId) {
                    throw new Error('Email ID is required');
                }
                const email = await outlookManager.getEmailById(emailId);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `📧 **邮件详情**\n\n**主题:** ${email.subject}\n**发件人:** ${email.sender}\n**收件人:** ${email.recipients?.join(', ')}\n**时间:** ${email.receivedTime}\n**已读:** ${email.isRead ? '是' : '否'}\n\n**内容摘要:**\n${email.body?.substring(0, 500)}${email.body && email.body.length > 500 ? '...' : ''}`,
                        },
                    ],
                };
            }
            case 'summarize_inbox': {
                const count = args?.count || 10;
                const emails = await outlookManager.getInboxEmails(count);
                const unreadCount = emails.filter(e => !e.isRead).length;
                return {
                    content: [
                        {
                            type: 'text',
                            text: `📊 **收件箱摘要**\n总计: ${emails.length} 封邮件\n未读: ${unreadCount} 封\n\n📋 **最近邮件:**\n` +
                                emails.slice(0, 5).map((email, index) => `${index + 1}. ${email.isRead ? '✅' : '📩'} **${email.subject}**\n   发件人: ${email.sender}\n   时间: ${email.receivedTime}\n`).join('\n')
                        },
                    ],
                };
            }
            case 'create_draft': {
                const draft = {
                    to: args?.to || [],
                    subject: args?.subject || '',
                    body: args?.body || '',
                    cc: args?.cc,
                    bcc: args?.bcc,
                };
                const result = await outlookManager.createDraft(draft);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `✅ **邮件草稿创建成功**\n\n**主题:** ${draft.subject}\n**收件人:** ${draft.to.join(', ')}\n${draft.cc ? `**抄送:** ${draft.cc.join(', ')}\n` : ''}**结果:** ${result}`,
                        },
                    ],
                };
            }
            case 'mark_email_as_read': {
                const emailId = args?.email_id;
                if (!emailId) {
                    throw new Error('Email ID is required');
                }
                await outlookManager.markAsRead(emailId);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `✅ **邮件已标记为已读**\n邮件ID: ${emailId}`,
                        },
                    ],
                };
            }
            case 'search_inbox_emails': {
                const query = args?.query;
                const count = args?.count || 10;
                if (!query) {
                    throw new Error('Search query is required');
                }
                const emails = await outlookManager.searchInboxEmails(query, count);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `🔍 **搜索结果: "${query}"**\n总计: ${emails.length} 件\n未读: ${emails.filter(e => !e.isRead).length} 件\n\n📋 **搜索结果一览:**\n` +
                                emails.map((email, index) => `${index + 1}. ${email.isRead ? '✅' : '📩'} **${email.subject}**\n   发件人: ${email.sender}\n   时间: ${email.receivedTime}\n   検索コンテキスト: ${email.body?.includes(query) ? '内容に一致' : '件名に一致'}: ${email.subject}\n   预览: ${email.body?.substring(0, 100)}...\n`).join('\n')
                        },
                    ],
                };
            }
            case 'search_sent_emails': {
                const query = args?.query;
                const count = args?.count || 10;
                if (!query) {
                    throw new Error('Search query is required');
                }
                const emails = await outlookManager.searchSentEmails(query, count);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `🔍 **検索結果: "${query}"**\n総計: ${emails.length} 件\n未読: ${emails.filter(e => !e.isRead).length} 件\n\n📋 **検索結果一覧:**\n` +
                                emails.map((email, index) => `${index + 1}. ${email.isRead ? '✅' : '📩'} **${email.subject}**\n   发件人: ${email.sender}\n   时间: ${email.receivedTime}\n   検索コンテキスト: ${email.body?.includes(query) ? '内容に一致' : '件名に一致'}: ${email.subject}\n   预览: ${email.body?.substring(0, 100)}...\n`).join('\n')
                        },
                    ],
                };
            }
            case 'get_sent_emails': {
                const count = args?.count || 10;
                const emails = await outlookManager.getSentEmails(count);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `📊 **邮件总览**\n总计: ${emails.length} 封邮件\n未读: ${emails.filter(e => !e.isRead).length} 封\n\n📋 **邮件列表:**\n` +
                                emails.map((email, index) => `${index + 1}. ${email.isRead ? '✅' : '📩'} **${email.subject}**\n   发件人: ${email.sender}\n   时间: ${email.receivedTime}\n   预览: ${email.body?.substring(0, 100)}...\n`).join('\n')
                        },
                    ],
                };
            }
            case 'duplicate_email_as_draft': {
                const sourceEmailId = args?.source_email_id;
                const newSubject = args?.new_subject;
                const newRecipients = args?.new_recipients;
                if (!sourceEmailId) {
                    throw new Error('Source email ID is required');
                }
                const result = await outlookManager.duplicateEmailAsDraft(sourceEmailId, newSubject, newRecipients);
                return {
                    content: [
                        {
                            type: 'text',
                            text: `✅ **邮件复制成功**\n\n**结果:** ${result}${newSubject ? `\n**新主题:** ${newSubject}` : ''}${newRecipients ? `\n**新收件人:** ${newRecipients.join(', ')}` : ''}`,
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
