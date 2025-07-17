export class EmailSummarizer {
    static summarizeEmail(email) {
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
    static summarizeMultipleEmails(emails) {
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
     * 全文検索でメールをフィルタリング
     * @param emails 検索対象のメール配列
     * @param searchTerm 検索語（大文字小文字を区別しない）
     * @returns 検索語を含むメールの配列
     */
    static searchEmails(emails, searchTerm) {
        if (!searchTerm.trim()) {
            return emails;
        }
        const normalizedSearchTerm = searchTerm.toLowerCase();
        return emails.filter(email => {
            // 件名での検索
            const subjectMatch = email.subject.toLowerCase().includes(normalizedSearchTerm);
            // 送信者での検索
            const senderMatch = email.sender.toLowerCase().includes(normalizedSearchTerm);
            // 本文での検索（HTMLタグを除去して検索）
            const cleanBody = email.body.replace(/<[^>]*>/g, '').toLowerCase();
            const bodyMatch = cleanBody.includes(normalizedSearchTerm);
            return subjectMatch || senderMatch || bodyMatch;
        });
    }
    /**
     * 検索語を含むメールのサマリーを生成（検索コンテキスト付き）
     * @param emails 検索結果のメール配列
     * @param searchTerm 検索語
     * @returns 検索結果のサマリー文字列
     */
    static summarizeSearchResults(emails, searchTerm) {
        if (emails.length === 0) {
            return `🔍 **検索結果**\n検索語「${searchTerm}」に一致するメールが見つかりませんでした。`;
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
     * 検索語の周辺コンテキストを抽出
     * @param email メールオブジェクト
     * @param searchTerm 検索語
     * @returns 検索語を含む文章の抜粋
     */
    static extractSearchContext(email, searchTerm) {
        const normalizedSearchTerm = searchTerm.toLowerCase();
        const cleanBody = email.body.replace(/<[^>]*>/g, '');
        const sentences = cleanBody.split(/[.!?。！？]+/);
        // 検索語を含む文を探す
        for (const sentence of sentences) {
            if (sentence.toLowerCase().includes(normalizedSearchTerm)) {
                // 文が長すぎる場合は検索語周辺を抽出
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
        // 件名または送信者で一致した場合
        if (email.subject.toLowerCase().includes(normalizedSearchTerm)) {
            return `件名に一致: ${email.subject}`;
        }
        if (email.sender.toLowerCase().includes(normalizedSearchTerm)) {
            return `送信者に一致: ${email.sender}`;
        }
        return '検索語を含むコンテキストが見つかりません';
    }
    static extractBodyPreview(body, maxLength = 200) {
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
    static extractKeyPoints(body) {
        const keyPoints = [];
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
            if (keyPoints.length >= 5)
                break;
        }
        return keyPoints.slice(0, 5);
    }
    static detectPriority(subject, body) {
        const text = (subject + ' ' + body).toLowerCase();
        const highPriorityWords = ['urgent', 'asap', 'emergency', 'critical', 'immediate', 'deadline'];
        const mediumPriorityWords = ['important', 'please review', 'meeting', 'call'];
        for (const word of highPriorityWords) {
            if (text.includes(word))
                return 'high';
        }
        for (const word of mediumPriorityWords) {
            if (text.includes(word))
                return 'medium';
        }
        return 'low';
    }
    static detectActionRequired(body) {
        const text = body.toLowerCase();
        const actionWords = [
            'please', 'can you', 'could you', 'would you', 'need you to',
            'respond', 'reply', 'confirm', 'approve', 'review', 'check',
            'call', 'meeting', 'schedule', 'deadline', 'urgent'
        ];
        return actionWords.some(word => text.includes(word));
    }
    static formatSummary(summary) {
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
${summary.keyPoints.length > 0 ? summary.keyPoints.map((point) => `• ${point}`).join('\n') : '无特殊关键点'}
`;
    }
    static formatMultipleSummaries(summaries) {
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
    /**
     * 検索結果のフォーマット
     * @param summaries 検索結果のサマリー配列
     * @param searchTerm 検索語
     * @returns フォーマットされた検索結果文字列
     */
    static formatSearchResults(summaries, searchTerm) {
        const unreadCount = summaries.filter(s => !s.isRead).length;
        const highPriorityCount = summaries.filter(s => s.priority === 'high').length;
        const actionRequiredCount = summaries.filter(s => s.actionRequired).length;
        let result = `🔍 **検索結果: "${searchTerm}"**\n`;
        result += `総計: ${summaries.length} 件\n`;
        result += `未読: ${unreadCount} 件\n`;
        result += `高優先級: ${highPriorityCount} 件\n`;
        result += `需要行動: ${actionRequiredCount} 件\n\n`;
        result += `📋 **検索結果一覧:**\n`;
        summaries.forEach((summary, index) => {
            const priorityIcon = summary.priority === 'high' ? '🔴' : summary.priority === 'medium' ? '🟡' : '🟢';
            const statusIcon = summary.isRead ? '✅' : '📬';
            const actionIcon = summary.actionRequired ? '⚠️' : '';
            result += `${index + 1}. ${statusIcon} ${priorityIcon} ${actionIcon} **${summary.subject}**\n`;
            result += `   发件人: ${summary.sender}\n`;
            result += `   时间: ${summary.receivedTime}\n`;
            result += `   検索コンテキスト: ${summary.searchContext}\n`;
            result += `   预览: ${summary.bodyPreview}\n\n`;
        });
        return result;
    }
}
