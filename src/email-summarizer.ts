import { EmailMessage } from './outlook-manager.js';

export class EmailSummarizer {
  
  static summarizeEmail(email: EmailMessage): string {
    const summary = {
      subject: email.subject,
      sender: email.sender,
      receivedTime: email.receivedTime,
      isRead: email.isRead,
      hasAttachments: email.hasAttachments,
      bodyPreview: this.extractBodyPreview(email.body),
      keyPoints: this.extractKeyPoints(email.body),
      priority: this.detectPriority(email.subject, email.body),
      actionRequired: this.detectActionRequired(email.body)
    };

    return this.formatSummary(summary);
  }

  static summarizeMultipleEmails(emails: EmailMessage[]): string {
    const summaries = emails.map(email => ({
      id: email.id,
      subject: email.subject,
      sender: email.sender,
      receivedTime: email.receivedTime,
      isRead: email.isRead,
      priority: this.detectPriority(email.subject, email.body),
      actionRequired: this.detectActionRequired(email.body),
      bodyPreview: this.extractBodyPreview(email.body)
    }));

    return this.formatMultipleSummaries(summaries);
  }

  private static extractBodyPreview(body: string, maxLength: number = 200): string {
    // Remove HTML tags and excessive whitespace
    const cleanBody = body
      .replace(/<[^>]*>/g, '')
      .replace(/\s+/g, ' ')
      .trim();

    if (cleanBody.length <= maxLength) {
      return cleanBody;
    }

    return cleanBody.substring(0, maxLength) + '...';
  }

  private static extractKeyPoints(body: string): string[] {
    const keyPoints: string[] = [];
    const cleanBody = body.replace(/<[^>]*>/g, '').toLowerCase();

    // Look for numbered lists
    const numberedLists = cleanBody.match(/\d+\.\s+[^\n]+/g);
    if (numberedLists) {
      keyPoints.push(...numberedLists.slice(0, 3));
    }

    // Look for bullet points
    const bulletPoints = cleanBody.match(/[•\-\*]\s+[^\n]+/g);
    if (bulletPoints) {
      keyPoints.push(...bulletPoints.slice(0, 3));
    }

    // Look for sentences with action words
    const actionWords = ['please', 'request', 'need', 'urgent', 'important', 'deadline', 'meeting', 'call'];
    const sentences = cleanBody.split(/[.!?]+/);
    
    for (const sentence of sentences) {
      for (const word of actionWords) {
        if (sentence.includes(word) && sentence.length > 20) {
          keyPoints.push(sentence.trim());
          break;
        }
      }
      if (keyPoints.length >= 5) break;
    }

    return keyPoints.slice(0, 5);
  }

  private static detectPriority(subject: string, body: string): 'high' | 'medium' | 'low' {
    const text = (subject + ' ' + body).toLowerCase();
    
    const highPriorityWords = ['urgent', 'asap', 'emergency', 'critical', 'immediate', 'deadline'];
    const mediumPriorityWords = ['important', 'please review', 'meeting', 'call'];
    
    for (const word of highPriorityWords) {
      if (text.includes(word)) return 'high';
    }
    
    for (const word of mediumPriorityWords) {
      if (text.includes(word)) return 'medium';
    }
    
    return 'low';
  }

  private static detectActionRequired(body: string): boolean {
    const text = body.toLowerCase();
    const actionWords = [
      'please', 'can you', 'could you', 'would you', 'need you to',
      'respond', 'reply', 'confirm', 'approve', 'review', 'check',
      'call', 'meeting', 'schedule', 'deadline', 'urgent'
    ];

    return actionWords.some(word => text.includes(word));
  }

  private static formatSummary(summary: any): string {
    return `
📧 **邮件摘要**
主题: ${summary.subject}
发件人: ${summary.sender}
接收时间: ${summary.receivedTime}
状态: ${summary.isRead ? '已读' : '未读'}
${summary.hasAttachments ? '📎 有附件' : ''}
优先级: ${summary.priority === 'high' ? '🔴 高' : summary.priority === 'medium' ? '🟡 中' : '🟢 低'}
${summary.actionRequired ? '⚠️ 需要回复或行动' : ''}

**内容预览:**
${summary.bodyPreview}

**关键点:**
${summary.keyPoints.length > 0 ? summary.keyPoints.map((point: string) => `• ${point}`).join('\n') : '无特殊关键点'}
`;
  }

  private static formatMultipleSummaries(summaries: any[]): string {
    const unreadCount = summaries.filter(s => !s.isRead).length;
    const highPriorityCount = summaries.filter(s => s.priority === 'high').length;
    const actionRequiredCount = summaries.filter(s => s.actionRequired).length;

    let result = `📊 **邮件总览**\n`;
    result += `总计: ${summaries.length} 封邮件\n`;
    result += `未读: ${unreadCount} 封\n`;
    result += `高优先级: ${highPriorityCount} 封\n`;
    result += `需要行动: ${actionRequiredCount} 封\n\n`;

    result += `📋 **邮件列表:**\n`;
    summaries.forEach((summary, index) => {
      const priorityIcon = summary.priority === 'high' ? '🔴' : summary.priority === 'medium' ? '🟡' : '🟢';
      const statusIcon = summary.isRead ? '✅' : '📬';
      const actionIcon = summary.actionRequired ? '⚠️' : '';
      
      result += `${index + 1}. ${statusIcon} ${priorityIcon} ${actionIcon} **${summary.subject}**\n`;
      result += `   发件人: ${summary.sender}\n`;
      result += `   时间: ${summary.receivedTime}\n`;
      result += `   预览: ${summary.bodyPreview}\n\n`;
    });

    return result;
  }
}
