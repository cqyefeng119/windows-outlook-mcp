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

  /**
   * Search emails with full-text search
   * @param emails Array of emails to search
   * @param searchTerm Search term (case-insensitive)
   * @returns Array of emails containing the search term
   */
  static searchEmails(emails: EmailMessage[], searchTerm: string): EmailMessage[] {
    if (!searchTerm.trim()) {
      return emails;
    }

    const normalizedSearchTerm = searchTerm.toLowerCase();
    
    return emails.filter(email => {
      // Search in subject
      const subjectMatch = email.subject.toLowerCase().includes(normalizedSearchTerm);
      
      // Search in sender
      const senderMatch = email.sender.toLowerCase().includes(normalizedSearchTerm);
      
      // Search in body content (remove HTML tags before searching)
      const cleanBody = email.body.replace(/<[^>]*>/g, '').toLowerCase();
      const bodyMatch = cleanBody.includes(normalizedSearchTerm);
      
      return subjectMatch || senderMatch || bodyMatch;
    });
  }

  /**
   * Generate summary of emails containing search term (with search context)
   * @param emails Array of search result emails
   * @param searchTerm Search term
   * @returns Summary string of search results
   */
  static summarizeSearchResults(emails: EmailMessage[], searchTerm: string): string {
    if (emails.length === 0) {
      return `🔍 **Search Results**\nNo emails found matching "${searchTerm}".`;
    }

    const summaries = emails.map(email => ({
      id: email.id,
      subject: email.subject,
      sender: email.sender,
      receivedTime: email.receivedTime,
      isRead: email.isRead,
      priority: this.detectPriority(email.subject, email.body),
      actionRequired: this.detectActionRequired(email.body),
      bodyPreview: this.extractBodyPreview(email.body),
      searchContext: this.extractSearchContext(email, searchTerm)
    }));

    return this.formatSearchResults(summaries, searchTerm);
  }

  /**
   * Extract search context around the search term
   * @param email Email object
   * @param searchTerm Search term
   * @returns Excerpt of sentence containing the search term
   */
  private static extractSearchContext(email: EmailMessage, searchTerm: string): string {
    const normalizedSearchTerm = searchTerm.toLowerCase();
    const cleanBody = email.body.replace(/<[^>]*>/g, '');
    const sentences = cleanBody.split(/[.!?。！？]+/);
    
    // Find sentence containing the search term
    for (const sentence of sentences) {
      if (sentence.toLowerCase().includes(normalizedSearchTerm)) {
        // If sentence is too long, extract context around search term
        const trimmedSentence = sentence.trim();
        if (trimmedSentence.length > 150) {
          const index = trimmedSentence.toLowerCase().indexOf(normalizedSearchTerm);
          const start = Math.max(0, index - 50);
          const end = Math.min(trimmedSentence.length, index + searchTerm.length + 50);
          return '...' + trimmedSentence.substring(start, end) + '...';
        }
        return trimmedSentence;
      }
    }
    
    // If matched in subject or sender
    if (email.subject.toLowerCase().includes(normalizedSearchTerm)) {
      return `Match in subject: ${email.subject}`;
    }
    if (email.sender.toLowerCase().includes(normalizedSearchTerm)) {
      return `Match in sender: ${email.sender}`;
    }
    
    return 'No search context found';
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
      keyPoints.push(...numberedLists.slice(0, 10));
    }

    // Look for bullet points
    const bulletPoints = cleanBody.match(/[•\-\*]\s+[^\n]+/g);
    if (bulletPoints) {
      keyPoints.push(...bulletPoints.slice(0, 10));
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
📧 **Email Summary**
Subject: ${summary.subject}
Sender: ${summary.sender}
Received: ${summary.receivedTime}
Status: ${summary.isRead ? 'Read' : 'Unread'}
${summary.hasAttachments ? '📎 Has Attachments' : ''}
Priority: ${summary.priority === 'high' ? '🔴 High' : summary.priority === 'medium' ? '🟡 Medium' : '🟢 Low'}
${summary.actionRequired ? '⚠️ Action Required' : ''}

**Content Preview:**
${summary.bodyPreview}

**Key Points:**
${summary.keyPoints.length > 0 ? summary.keyPoints.map((point: string) => `• ${point}`).join('\n') : 'No special key points'}
`;
  }

  private static formatMultipleSummaries(summaries: any[]): string {
    const unreadCount = summaries.filter(s => !s.isRead).length;
    const highPriorityCount = summaries.filter(s => s.priority === 'high').length;
    const actionRequiredCount = summaries.filter(s => s.actionRequired).length;

    let result = `📊 **Email Overview**\n`;
    result += `Total: ${summaries.length} emails\n`;
    result += `Unread: ${unreadCount} emails\n`;
    result += `High Priority: ${highPriorityCount} emails\n`;
    result += `Action Required: ${actionRequiredCount} emails\n\n`;

    result += `📋 **Email List:**\n`;
    summaries.forEach((summary, index) => {
      const priorityIcon = summary.priority === 'high' ? '🔴' : summary.priority === 'medium' ? '🟡' : '🟢';
      const statusIcon = summary.isRead ? '✅' : '📬';
      const actionIcon = summary.actionRequired ? '⚠️' : '';
      
      result += `${index + 1}. ${statusIcon} ${priorityIcon} ${actionIcon} **${summary.subject}**\n`;
      result += `   From: ${summary.sender}\n`;
      result += `   Time: ${summary.receivedTime}\n`;
      result += `   Preview: ${summary.bodyPreview}\n\n`;
    });

    return result;
  }

  /**
   * Format search results
   * @param summaries Array of search result summaries
   * @param searchTerm Search term
   * @returns Formatted search results string
   */
  private static formatSearchResults(summaries: any[], searchTerm: string): string {
    const unreadCount = summaries.filter(s => !s.isRead).length;
    const highPriorityCount = summaries.filter(s => s.priority === 'high').length;
    const actionRequiredCount = summaries.filter(s => s.actionRequired).length;

    let result = `🔍 **Search Results: "${searchTerm}"**\n`;
    result += `Total: ${summaries.length} items\n`;
    result += `Unread: ${unreadCount} items\n`;
    result += `High Priority: ${highPriorityCount} items\n`;
    result += `Action Required: ${actionRequiredCount} items\n\n`;

    result += `📋 **Search Results List:**\n`;
    summaries.forEach((summary, index) => {
      const priorityIcon = summary.priority === 'high' ? '🔴' : summary.priority === 'medium' ? '🟡' : '🟢';
      const statusIcon = summary.isRead ? '✅' : '📬';
      const actionIcon = summary.actionRequired ? '⚠️' : '';
      
      result += `${index + 1}. ${statusIcon} ${priorityIcon} ${actionIcon} **${summary.subject}**\n`;
      result += `   From: ${summary.sender}\n`;
      result += `   Time: ${summary.receivedTime}\n`;
      result += `   Search Context: ${summary.searchContext}\n`;
      result += `   Preview: ${summary.bodyPreview}\n\n`;
    });

    return result;
  }
}
