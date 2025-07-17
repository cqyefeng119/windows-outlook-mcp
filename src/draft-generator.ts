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
      subject: 'Meeting Request - {topic}',
      body: `Hello,

I would like to invite you to a meeting regarding {topic}.

Meeting Details:
- Date: {date}
- Time: {time}
- Location: {location}
- Duration: {duration}

Agenda:
{agenda}

Please confirm your availability. If you have any conflicts, please let me know your preferred time.

Thank you!

Best regards,
{sender}`,
      category: 'meeting'
    },
    {
      name: 'follow_up',
      subject: 'Follow-up - {topic}',
      body: `Hello,

I wanted to follow up on our previous discussion about {topic}.

{details}

Please let me know if you have any questions or if there's anything else I can help with.

Best regards,
{sender}`,
      category: 'follow_up'
    },
    {
      name: 'thank_you',
      subject: 'Thank you - {topic}',
      body: `Hello,

Thank you for your help and support with {topic}.

{details}

Your assistance is greatly appreciated.

Best regards,
{sender}`,
      category: 'thank_you'
    },
    {
      name: 'status_update',
      subject: 'Status Update - {project}',
      body: `Hello,

Here's a status update on {project}.

Current Progress:
{progress}

Next Steps:
{next_steps}

Please let me know if you have any questions or concerns.

Best regards,
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
        body = `Hello,

Thank you for your email. I agree with your proposal/suggestion.

${customMessage || ''}

Please let me know the next steps.

Best regards`;
        break;
      
      case 'decline':
        body = `Hello,

Thank you for your email. I'm sorry, but I cannot accept your proposal/suggestion.

${customMessage || ''}

If there are alternative options, please feel free to contact me.

Best regards`;
        break;
      
      case 'info_request':
        body = `Hello,

Thank you for your email. I need more information to make a decision.

${customMessage || 'Please provide more details.'}

Looking forward to your reply.

Best regards`;
        break;
      
      case 'custom':
        body = customMessage || '';
        break;
    }

    // Add original email reference
    body += `\n\n---\nOriginal Email:\n${this.formatOriginalEmail(originalBody)}`;

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
        subject: 'Meeting Request',
        body: `Hello,

I would like to schedule a meeting to discuss the following:

${context}

Please let me know your available time, and I will send a formal meeting invitation.

Best regards`
      },
      'request_information': {
        subject: 'Information Request',
        body: `Hello,

I need to understand the following information:

${context}

I would appreciate it if you could provide relevant information.

Best regards`
      },
      'project_update': {
        subject: 'Project Update',
        body: `Hello,

Here is the latest project update:

${context}

If you have any questions or need to discuss anything, please feel free to contact me.

Best regards`
      },
      'general': {
        subject: 'Email',
        body: `Hello,

${context}

Best regards`
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
