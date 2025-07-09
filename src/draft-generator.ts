import { EmailDraft } from './outlook-manager.js';

export interface DraftTemplate {
  name: string;
  subject: string;
  body: string;
  category: string;
}

export class DraftGenerator {
  private static templates: DraftTemplate[] = [
    {
      name: 'meeting_request',
      subject: '会议邀请 - {topic}',
      body: `您好，

我想邀请您参加关于{topic}的会议。

会议详情：
- 日期：{date}
- 时间：{time}
- 地点：{location}
- 预计时长：{duration}

会议议程：
{agenda}

请确认您的参会时间，如有冲突请告知合适的时间安排。

谢谢！

最好的问候，
{sender}`,
      category: 'meeting'
    },
    {
      name: 'follow_up',
      subject: '跟进 - {topic}',
      body: `您好，

我想跟进一下我们之前讨论的{topic}。

{details}

请告知当前进展情况，如有需要协助的地方请随时联系我。

期待您的回复。

最好的问候，
{sender}`,
      category: 'follow_up'
    },
    {
      name: 'thank_you',
      subject: '感谢 - {topic}',
      body: `您好，

感谢您在{topic}方面的帮助和支持。

{details}

您的协助对我们非常重要，我们深表感谢。

最好的问候，
{sender}`,
      category: 'thank_you'
    },
    {
      name: 'status_update',
      subject: '状态更新 - {project}',
      body: `您好，

这是关于{project}的状态更新。

当前进展：
{progress}

下一步计划：
{next_steps}

如有问题或需要讨论的地方，请随时联系我。

最好的问候，
{sender}`,
      category: 'update'
    }
  ];

  static generateDraft(templateName: string, variables: Record<string, string>, recipients: string[]): EmailDraft {
    const template = this.templates.find(t => t.name === templateName);
    if (!template) {
      throw new Error(`Template '${templateName}' not found`);
    }

    const subject = this.replaceVariables(template.subject, variables);
    const body = this.replaceVariables(template.body, variables);

    return {
      to: recipients,
      subject,
      body,
      isHtml: false
    };
  }

  static generateReplyDraft(originalSubject: string, originalBody: string, replyType: 'agree' | 'decline' | 'info_request' | 'custom', customMessage?: string): EmailDraft {
    let subject = originalSubject;
    if (!subject.toLowerCase().startsWith('re:')) {
      subject = `Re: ${subject}`;
    }

    let body = '';
    
    switch (replyType) {
      case 'agree':
        body = `您好，

谢谢您的邮件。我同意您的提议/建议。

${customMessage || ''}

请告知下一步的安排。

最好的问候`;
        break;
      
      case 'decline':
        body = `您好，

谢谢您的邮件。很抱歉，我无法接受您的提议/建议。

${customMessage || ''}

如有其他方案，请随时联系我。

最好的问候`;
        break;
      
      case 'info_request':
        body = `您好，

谢谢您的邮件。我需要更多信息来做出决定。

${customMessage || '请提供更多详细信息。'}

期待您的回复。

最好的问候`;
        break;
      
      case 'custom':
        body = customMessage || '';
        break;
    }

    // Add original email reference
    body += `\n\n---\n原始邮件:\n${this.formatOriginalEmail(originalBody)}`;

    return {
      to: [],
      subject,
      body,
      isHtml: false
    };
  }

  static generateSmartDraft(context: string, intent: string, recipients: string[]): EmailDraft {
    const intents = {
      'schedule_meeting': {
        subject: '会议安排请求',
        body: `您好，

我想安排一个会议讨论以下事项：

${context}

请告知您方便的时间，我会发送正式的会议邀请。

最好的问候`
      },
      'request_information': {
        subject: '信息请求',
        body: `您好，

我需要了解以下信息：

${context}

如果您能提供相关信息，我将不胜感激。

最好的问候`
      },
      'project_update': {
        subject: '项目更新',
        body: `您好，

这是项目的最新更新：

${context}

如有问题或需要讨论的地方，请随时联系我。

最好的问候`
      },
      'general': {
        subject: '邮件',
        body: `您好，

${context}

最好的问候`
      }
    };

    const template = intents[intent as keyof typeof intents] || intents.general;
    
    return {
      to: recipients,
      subject: template.subject,
      body: template.body,
      isHtml: false
    };
  }

  static getAvailableTemplates(): DraftTemplate[] {
    return this.templates;
  }

  private static replaceVariables(text: string, variables: Record<string, string>): string {
    let result = text;
    for (const [key, value] of Object.entries(variables)) {
      result = result.replace(new RegExp(`\\{${key}\\}`, 'g'), value);
    }
    return result;
  }

  private static formatOriginalEmail(originalBody: string): string {
    // Clean up the original email body and limit to first few lines
    const cleanBody = originalBody
      .replace(/<[^>]*>/g, '')
      .replace(/\s+/g, ' ')
      .trim();
    
    const lines = cleanBody.split('\n').slice(0, 5);
    return lines.join('\n') + (lines.length >= 5 ? '\n...' : '');
  }
}
