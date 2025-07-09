#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { OutlookManager, EmailMessage, EmailDraft } from './outlook-manager.js';
import { EmailSummarizer } from './email-summarizer.js';
import { DraftGenerator } from './draft-generator.js';

class OutlookMCPServer {
  private server: McpServer;
  private outlookManager: OutlookManager;

  constructor() {
    this.server = new McpServer({
      name: 'outlook-mcp-server',
      version: '1.0.0',
    });

    this.outlookManager = new OutlookManager();
    this.setupTools();
  }

  private setupTools(): void {
    // 获取收件箱邮件列表
    this.server.tool(
      'get_inbox_emails',
      {
        count: z.number().optional().default(10).describe('要获取的邮件数量'),
      },
      async ({ count }) => {
        try {
          const emails = await this.outlookManager.getInboxEmails(count);
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify(emails, null, 2),
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // 根据ID获取特定邮件
    this.server.tool(
      'get_email_by_id',
      {
        id: z.string().describe('邮件ID'),
      },
      async ({ id }) => {
        try {
          const email = await this.outlookManager.getEmailById(id);
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify(email, null, 2),
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // 搜索邮件
    this.server.tool(
      'search_emails',
      {
        query: z.string().describe('搜索关键词'),
        count: z.number().optional().default(10).describe('返回结果数量'),
      },
      async ({ query, count }) => {
        try {
          const emails = await this.outlookManager.searchEmails(query, count);
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify(emails, null, 2),
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // 总结单个邮件
    this.server.tool(
      'summarize_email',
      {
        email_id: z.string().describe('要总结的邮件ID'),
      },
      async ({ email_id }) => {
        try {
          const email = await this.outlookManager.getEmailById(email_id);
          const summary = EmailSummarizer.summarizeEmail(email);
          return {
            content: [
              {
                type: 'text' as const,
                text: summary,
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // 总结收件箱邮件
    this.server.tool(
      'summarize_inbox',
      {
        count: z.number().optional().default(10).describe('要总结的邮件数量'),
      },
      async ({ count }) => {
        try {
          const emails = await this.outlookManager.getInboxEmails(count);
          const summary = EmailSummarizer.summarizeMultipleEmails(emails);
          return {
            content: [
              {
                type: 'text' as const,
                text: summary,
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // 创建邮件草稿
    this.server.tool(
      'create_draft',
      {
        to: z.array(z.string()).describe('收件人邮箱地址'),
        cc: z.array(z.string()).optional().describe('抄送邮箱地址'),
        bcc: z.array(z.string()).optional().describe('密送邮箱地址'),
        subject: z.string().describe('邮件主题'),
        body: z.string().describe('邮件内容'),
      },
      async ({ to, cc, bcc, subject, body }) => {
        try {
          const draft: EmailDraft = {
            to,
            cc,
            bcc,
            subject,
            body,
          };
          const result = await this.outlookManager.createDraft(draft);
          return {
            content: [
              {
                type: 'text' as const,
                text: `草稿创建成功: ${result}`,
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // 使用模板生成草稿
    this.server.tool(
      'generate_draft_from_template',
      {
        template_name: z.enum(['meeting_request', 'follow_up', 'thank_you', 'status_update']).describe('模板名称'),
        variables: z.record(z.string()).describe('模板变量键值对'),
        recipients: z.array(z.string()).describe('收件人邮箱地址'),
      },
      async ({ template_name, variables, recipients }) => {
        try {
          const draft = DraftGenerator.generateDraft(template_name, variables, recipients);
          const result = await this.outlookManager.createDraft(draft);
          return {
            content: [
              {
                type: 'text' as const,
                text: `从模板生成的草稿创建成功: ${result}\n\n草稿内容:\n主题: ${draft.subject}\n内容: ${draft.body}`,
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // 生成回复草稿
    this.server.tool(
      'generate_reply_draft',
      {
        original_email_id: z.string().describe('原始邮件ID'),
        reply_type: z.enum(['agree', 'decline', 'info_request', 'custom']).describe('回复类型'),
        custom_message: z.string().optional().describe('自定义消息内容'),
      },
      async ({ original_email_id, reply_type, custom_message }) => {
        try {
          const originalEmail = await this.outlookManager.getEmailById(original_email_id);
          const draft = DraftGenerator.generateReplyDraft(
            originalEmail.subject,
            originalEmail.body,
            reply_type,
            custom_message
          );
          const result = await this.outlookManager.createDraft(draft);
          return {
            content: [
              {
                type: 'text' as const,
                text: `回复草稿创建成功: ${result}\n\n草稿内容:\n主题: ${draft.subject}\n内容: ${draft.body}`,
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // 智能生成草稿
    this.server.tool(
      'generate_smart_draft',
      {
        context: z.string().describe('邮件上下文内容'),
        intent: z.enum(['schedule_meeting', 'request_information', 'project_update', 'general']).describe('邮件意图'),
        recipients: z.array(z.string()).describe('收件人邮箱地址'),
      },
      async ({ context, intent, recipients }) => {
        try {
          const draft = DraftGenerator.generateSmartDraft(context, intent, recipients);
          const result = await this.outlookManager.createDraft(draft);
          return {
            content: [
              {
                type: 'text' as const,
                text: `智能草稿创建成功: ${result}\n\n草稿内容:\n主题: ${draft.subject}\n内容: ${draft.body}`,
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // 获取可用模板
    this.server.tool(
      'get_draft_templates',
      {},
      async () => {
        try {
          const templates = DraftGenerator.getAvailableTemplates();
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify(templates, null, 2),
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // 标记邮件为已读
    this.server.tool(
      'mark_email_as_read',
      {
        email_id: z.string().describe('邮件ID'),
      },
      async ({ email_id }) => {
        try {
          await this.outlookManager.markAsRead(email_id);
          return {
            content: [
              {
                type: 'text' as const,
                text: '邮件已标记为已读',
              },
            ],
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );
  }

  async run(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Outlook MCP Server running on stdio');
  }
}

const server = new OutlookMCPServer();
server.run().catch(console.error);
