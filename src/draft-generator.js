export class DraftGenerator {
    static templates = [
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
    static generateDraft(templateName, variables, recipients) {
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
    static generateReplyDraft(originalSubject, originalBody, replyType, customMessage) {
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
    static generateSmartDraft(context, intent, recipients) {
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
        const template = intents[intent] || intents.general;
        return {
            to: recipients,
            subject: template.subject,
            body: template.body,
            isHtml: false
        };
    }
    static getAvailableTemplates() {
        return this.templates;
    }
    static replaceVariables(text, variables) {
        let result = text;
        for (const [key, value] of Object.entries(variables)) {
            result = result.replace(new RegExp(`\\{${key}\\}`, 'g'), value);
        }
        return result;
    }
    static formatOriginalEmail(originalBody) {
        // Clean up the original email body and limit to first few lines
        const cleanBody = originalBody
            .replace(/<[^>]*>/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const lines = cleanBody.split('\n').slice(0, 5);
        return lines.join('\n') + (lines.length >= 5 ? '\n...' : '');
    }
}
