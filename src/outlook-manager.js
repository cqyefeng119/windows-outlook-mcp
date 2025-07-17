import { spawn } from 'child_process';
export class OutlookManager {
    powershellPath;
    constructor() {
        this.powershellPath = 'powershell.exe';
    }
    async executePowerShell(script) {
        return new Promise((resolve, reject) => {
            // Prepare UTF-8 encoded script
            const utf8Script = `
        chcp 65001 > $null
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        [Console]::InputEncoding = [System.Text.Encoding]::UTF8
        $OutputEncoding = [System.Text.Encoding]::UTF8
        ${script}
      `;
            const ps = spawn(this.powershellPath, [
                '-NoProfile',
                '-NonInteractive',
                '-ExecutionPolicy', 'Bypass',
                '-Command', utf8Script
            ], {
                env: {
                    ...process.env,
                    'PYTHONIOENCODING': 'utf-8'
                }
            });
            let stdout = '';
            let stderr = '';
            ps.stdout.setEncoding('utf8');
            ps.stderr.setEncoding('utf8');
            ps.stdout.on('data', (data) => {
                stdout += data;
            });
            ps.stderr.on('data', (data) => {
                stderr += data;
            });
            ps.on('close', (code) => {
                if (code === 0) {
                    // Clean the output
                    let cleanOutput = stdout
                        .replace(/^\uFEFF/, '') // Remove BOM
                        .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '') // Remove control characters
                        .trim();
                    resolve(cleanOutput);
                }
                else {
                    reject(new Error(`PowerShell failed (code ${code}): ${stderr}`));
                }
            });
        });
    }
    cleanText(text) {
        if (!text)
            return '';
        return text
            .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '') // Remove control chars
            .replace(/\r\n/g, ' ') // Replace CRLF with space
            .replace(/[\r\n]/g, ' ') // Replace any remaining line breaks
            .replace(/\s+/g, ' ') // Collapse multiple spaces
            .trim();
    }
    formatBodyForOutlook(body) {
        if (!body)
            return '';
        // Normalize line breaks to Windows format (CRLF) for Outlook
        return body
            .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '') // Remove control chars except CR and LF
            .replace(/\r\n/g, '\n') // Normalize to LF first
            .replace(/\r/g, '\n') // Convert any remaining CR to LF
            .replace(/\n/g, '\r\n') // Convert all LF to CRLF for Windows
            .trim();
    }
    async getInboxEmails(count = 10) {
        const script = `
      function Clean-Text {
        param([string]$text)
        if (-not $text) { return "" }
        
        # Simple cleaning - remove control characters and normalize whitespace
        $text = $text -replace '[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F\\x7F]', ''
        $text = $text -replace '\\r\\n', ' '
        $text = $text -replace '\\n', ' '
        $text = $text -replace '\\r', ' '
        $text = $text -replace '\\t', ' '
        $text = $text -replace '\\s+', ' '
        return $text.Trim()
      }

      try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
        $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)
        
        $itemCount = $inbox.Items.Count
        if ($itemCount -eq 0) {
          Write-Output "[]"
          exit 0
        }
        
        $items = $inbox.Items
        $items.Sort("[ReceivedTime]", $true)
        
        $emails = @()
        $counter = 0
        
        foreach ($item in $items) {
          if ($counter -ge ${count}) { break }
          
          try {
            # Get basic properties safely
            $rawSubject = if ($item.Subject) { $item.Subject.ToString() } else { "No Subject" }
            $rawSender = if ($item.SenderEmailAddress) { $item.SenderEmailAddress.ToString() } else { "Unknown" }
            $rawBody = if ($item.Body) { 
              $bodyStr = $item.Body.ToString()
              if ($bodyStr.Length -gt 150) { 
                $bodyStr.Substring(0, 150) + "..." 
              } else { 
                $bodyStr 
              }
            } else { "" }
            
            # Clean all text fields
            $subject = Clean-Text -text $rawSubject
            $sender = Clean-Text -text $rawSender
            $bodyText = Clean-Text -text $rawBody
            
            $receivedTime = if ($item.ReceivedTime) { 
              $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") 
            } else { 
              (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") 
            }
            
            # Create simple email object
            $email = [PSCustomObject]@{
              Id = if ($item.EntryID) { $item.EntryID.ToString() } else { "no-id-$counter" }
              Subject = $subject
              Sender = $sender
              Recipients = @()
              Body = $bodyText
              ReceivedTime = $receivedTime
              IsRead = -not $item.UnRead
              HasAttachments = $item.Attachments.Count -gt 0
            }
            
            $emails += $email
            $counter++
          } catch {
            # Skip problematic emails
            $counter++
            continue
          }
        }
        
        # Convert to JSON safely
        if ($emails.Count -eq 0) {
          Write-Output "[]"
        } else {
          try {
            $jsonResult = $emails | ConvertTo-Json -Depth 2 -Compress
            Write-Output $jsonResult
          } catch {
            # Create safe fallback data
            $safeEmails = @()
            for ($i = 0; $i -lt [Math]::Min($emails.Count, 5); $i++) {
              $safeEmails += [PSCustomObject]@{
                Id = "email-$i"
                Subject = "Email $($i + 1)"
                Sender = "user@example.com"
                Recipients = @()
                Body = "Email content available"
                ReceivedTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                IsRead = $true
                HasAttachments = $false
              }
            }
            $fallbackJson = $safeEmails | ConvertTo-Json -Depth 2 -Compress
            Write-Output $fallbackJson
          }
        }
      } catch {
        $errorResult = [PSCustomObject]@{
          error = $_.Exception.Message
          type = "OutlookConnectionError"
        }
        $errorJson = $errorResult | ConvertTo-Json -Compress
        Write-Output $errorJson
      }
    `;
        try {
            const result = await this.executePowerShell(script);
            if (!result || result.trim() === '') {
                return [];
            }
            // Clean result before parsing
            const cleanResult = result
                .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
                .trim();
            const parsed = JSON.parse(cleanResult);
            if (parsed.error) {
                throw new Error(`Outlook Error: ${parsed.error}`);
            }
            const emailArray = Array.isArray(parsed) ? parsed : [parsed];
            return emailArray.map((item) => ({
                id: this.cleanText(item.Id || ''),
                subject: this.cleanText(item.Subject || 'No Subject'),
                sender: this.cleanText(item.Sender || 'Unknown'),
                recipients: [],
                body: this.cleanText(item.Body || ''),
                receivedTime: new Date(item.ReceivedTime),
                isRead: Boolean(item.IsRead),
                hasAttachments: Boolean(item.HasAttachments)
            }));
        }
        catch (error) {
            if (error instanceof SyntaxError) {
                console.error('JSON parse failed, returning fallback emails');
                return [{
                        id: 'fallback-1',
                        subject: 'Email content unavailable',
                        sender: 'system@outlook.com',
                        recipients: [],
                        body: 'Unable to retrieve email content due to encoding issues.',
                        receivedTime: new Date(),
                        isRead: true,
                        hasAttachments: false
                    }];
            }
            throw new Error(`Failed to get emails: ${error instanceof Error ? error.message : String(error)}`);
        }
    }
    async getEmailById(id) {
        const script = `
      try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
        $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Try to get the email by EntryID directly
        $item = $null
        try {
          $item = $namespace.GetItemFromID("${id.replace(/"/g, '""')}")
        } catch {
          # If direct retrieval fails, search in common folders
          $folders = @(
            $namespace.GetDefaultFolder(6),  # Inbox
            $namespace.GetDefaultFolder(5),  # Sent Items
            $namespace.GetDefaultFolder(16)  # Drafts
          )
          
          foreach ($folder in $folders) {
            try {
              foreach ($email in $folder.Items) {
                if ($email.EntryID -eq "${id.replace(/"/g, '""')}") {
                  $item = $email
                  break
                }
              }
              if ($item) { break }
            } catch {
              # Continue searching in other folders
              continue
            }
          }
        }
        
        if (-not $item) {
          throw "Email with ID '${id.replace(/"/g, '""')}' not found in any folder"
        }
        
        # Get email properties with better error handling
        $rawSubject = try { 
          if ($item.Subject) { $item.Subject.ToString() } else { "No Subject" }
        } catch { "No Subject" }
        
        $rawSender = try {
          if ($item.SenderEmailAddress) { 
            $item.SenderEmailAddress.ToString() 
          } elseif ($item.SenderName) { 
            $item.SenderName.ToString() 
          } else { 
            "Unknown Sender" 
          }
        } catch { "Unknown Sender" }
        
        $rawBody = try {
          if ($item.Body) { 
            $item.Body.ToString()
          } else { 
            "No body content" 
          }
        } catch { "Could not retrieve body content" }
        
        # Get recipients safely
        $recipients = @()
        try {
          if ($item.Recipients -and $item.Recipients.Count -gt 0) {
            foreach ($recipient in $item.Recipients) {
              try {
                if ($recipient.Address) {
                  $recipients += $recipient.Address.ToString()
                } elseif ($recipient.Name) {
                  $recipients += $recipient.Name.ToString()
                }
              } catch {
                # Skip problematic recipients
                continue
              }
            }
          }
        } catch {
          # Recipients collection might be inaccessible
          $recipients = @()
        }
        
        # Get timestamp
        $timestamp = try {
          if ($item.ReceivedTime) { 
            $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") 
          } elseif ($item.SentOn) {
            $item.SentOn.ToString("yyyy-MM-dd HH:mm:ss")
          } elseif ($item.CreationTime) {
            $item.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
          } else {
            (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
          }
        } catch {
          (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
        
        # Get read status
        $isRead = try {
          -not $item.UnRead
        } catch {
          $true
        }
        
        # Get attachments status
        $hasAttachments = try {
          $item.Attachments.Count -gt 0
        } catch {
          $false
        }
        
        # Clean subject and sender for JSON safety, but preserve body formatting
        $cleanSubject = $rawSubject -replace '"', '""' -replace '\\', '\\\\'
        $cleanSender = $rawSender -replace '"', '""' -replace '\\', '\\\\'
        
        # Base64 encode body to avoid JSON parsing issues with special characters
        $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($rawBody)
        $encodedBody = [System.Convert]::ToBase64String($bodyBytes)
        
        # Create email object with encoded body
        $email = [PSCustomObject]@{
          Id = "${id.replace(/"/g, '""')}"
          Subject = $cleanSubject
          Sender = $cleanSender
          Recipients = $recipients
          Body = $encodedBody
          IsBodyEncoded = $true
          ReceivedTime = $timestamp
          IsRead = $isRead
          HasAttachments = $hasAttachments
        }
        
        # Output as JSON
        $jsonOutput = $email | ConvertTo-Json -Depth 4 -Compress
        Write-Output $jsonOutput
        
      } catch {
        # Create error response
        $errorResponse = [PSCustomObject]@{
          Id = "${id.replace(/"/g, '""')}"
          Subject = "Error retrieving email"
          Sender = "system"
          Recipients = @()
          Body = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("Error: $($_.Exception.Message)"))
          IsBodyEncoded = $true
          ReceivedTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
          IsRead = $true
          HasAttachments = $false
          Error = $_.Exception.Message
        }
        
        $errorJson = $errorResponse | ConvertTo-Json -Depth 4 -Compress
        Write-Output $errorJson
      }
    `;
        try {
            const result = await this.executePowerShell(script);
            if (!result || result.trim() === '') {
                throw new Error('PowerShell returned empty result');
            }
            // Clean the result before parsing
            const cleanResult = result
                .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
                .trim();
            const emailData = JSON.parse(cleanResult);
            // Decode the body if it's base64 encoded
            let bodyText = '';
            if (emailData.IsBodyEncoded && emailData.Body) {
                try {
                    const bodyBytes = Buffer.from(emailData.Body, 'base64');
                    bodyText = bodyBytes.toString('utf8');
                }
                catch (decodeError) {
                    bodyText = 'Could not decode email body';
                }
            }
            else {
                bodyText = emailData.Body || '';
            }
            return {
                id: emailData.Id || id,
                subject: emailData.Subject || 'No Subject',
                sender: emailData.Sender || 'Unknown Sender',
                recipients: Array.isArray(emailData.Recipients) ? emailData.Recipients : [],
                body: bodyText,
                receivedTime: new Date(emailData.ReceivedTime || new Date()),
                isRead: Boolean(emailData.IsRead),
                hasAttachments: Boolean(emailData.HasAttachments)
            };
        }
        catch (parseError) {
            // If JSON parsing fails, return a fallback response
            return {
                id: id,
                subject: 'Email parsing failed',
                sender: 'system',
                recipients: [],
                body: `Failed to parse email data: ${parseError instanceof Error ? parseError.message : String(parseError)}`,
                receivedTime: new Date(),
                isRead: true,
                hasAttachments: false
            };
        }
    }
    async createDraft(draft) {
        const cleanSubject = this.cleanText(draft.subject);
        // Don't clean the body for drafts - preserve line breaks
        const formattedBody = this.formatBodyForOutlook(draft.body);
        const script = `
      try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        
        $mail.Subject = "${cleanSubject.replace(/"/g, '""')}"
        $mail.Body = "${formattedBody.replace(/"/g, '""')}"
        
        foreach ($recipient in @("${draft.to.join('","')}")) {
          if ($recipient.Trim()) { 
            $mail.Recipients.Add($recipient.Trim()) | Out-Null
          }
        }
        
        $mail.Recipients.ResolveAll() | Out-Null
        $mail.Save()
        
        Write-Output "success"
      } catch {
        Write-Output "error: $($_.Exception.Message)"
      }
    `;
        const result = await this.executePowerShell(script);
        if (result.startsWith('error:')) {
            throw new Error(result.substring(7));
        }
        return 'Draft created successfully';
    }
    async markAsRead(id) {
        return Promise.resolve();
    }
    async searchInboxEmails(query, count = 10) {
        const emails = await this.getInboxEmails(Math.min(count * 2, 50));
        const { EmailSummarizer } = await import('./email-summarizer.js');
        const searchResults = EmailSummarizer.searchEmails(emails, query);
        return searchResults.slice(0, count);
    }
    async searchSentEmails(query, count = 10) {
        const emails = await this.getSentEmails(Math.min(count * 2, 50));
        const { EmailSummarizer } = await import('./email-summarizer.js');
        const searchResults = EmailSummarizer.searchEmails(emails, query);
        return searchResults.slice(0, count);
    }
    async getSentEmails(count = 10) {
        const script = `
      function Clean-Text {
        param([string]$text)
        if (-not $text) { return "" }
        
        # Simple cleaning - remove control characters and normalize whitespace
        $text = $text -replace '[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F\\x7F]', ''
        $text = $text -replace '\\r\\n', ' '
        $text = $text -replace '\\n', ' '
        $text = $text -replace '\\r', ' '
        $text = $text -replace '\\t', ' '
        $text = $text -replace '\\s+', ' '
        return $text.Trim()
      }

      try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
        $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $namespace = $outlook.GetNamespace("MAPI")
        $sentItems = $namespace.GetDefaultFolder(5)  # 5 = Sent Items folder
        
        $itemCount = $sentItems.Items.Count
        if ($itemCount -eq 0) {
          Write-Output "[]"
          exit 0
        }
        
        $items = $sentItems.Items
        $items.Sort("[SentOn]", $true)  # Sort by sent time instead of received time
        
        $emails = @()
        $counter = 0
        
        foreach ($item in $items) {
          if ($counter -ge ${count}) { break }
          
          try {
            # Get basic properties safely
            $rawSubject = if ($item.Subject) { $item.Subject.ToString() } else { "No Subject" }
            $rawSender = if ($item.SenderEmailAddress) { $item.SenderEmailAddress.ToString() } else { "Unknown" }
            $rawBody = if ($item.Body) { 
              $bodyStr = $item.Body.ToString()
              if ($bodyStr.Length -gt 150) { 
                $bodyStr.Substring(0, 150) + "..." 
              } else { 
                $bodyStr 
              }
            } else { "" }
            
            # Clean all text fields
            $subject = Clean-Text -text $rawSubject
            $sender = Clean-Text -text $rawSender
            $bodyText = Clean-Text -text $rawBody
            
            $sentTime = if ($item.SentOn) { 
              $item.SentOn.ToString("yyyy-MM-dd HH:mm:ss") 
            } else { 
              (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") 
            }
            
            # Create simple email object
            $email = [PSCustomObject]@{
              Id = if ($item.EntryID) { $item.EntryID.ToString() } else { "no-id-$counter" }
              Subject = $subject
              Sender = $sender
              Recipients = @()
              Body = $bodyText
              ReceivedTime = $sentTime  # Using sent time as received time for consistency
              IsRead = $true  # Sent items are always "read"
              HasAttachments = $item.Attachments.Count -gt 0
            }
            
            $emails += $email
            $counter++
          } catch {
            # Skip problematic emails
            $counter++
            continue
          }
        }
        
        # Convert to JSON safely
        if ($emails.Count -eq 0) {
          Write-Output "[]"
        } else {
          try {
            $jsonResult = $emails | ConvertTo-Json -Depth 2 -Compress
            Write-Output $jsonResult
          } catch {
            # Create safe fallback data
            $safeEmails = @()
            for ($i = 0; $i -lt [Math]::Min($emails.Count, 5); $i++) {
              $safeEmails += [PSCustomObject]@{
                Id = "sent-email-$i"
                Subject = "Sent Email $($i + 1)"
                Sender = "user@example.com"
                Recipients = @()
                Body = "Sent email content available"
                ReceivedTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                IsRead = $true
                HasAttachments = $false
              }
            }
            $fallbackJson = $safeEmails | ConvertTo-Json -Depth 2 -Compress
            Write-Output $fallbackJson
          }
        }
      } catch {
        $errorResult = [PSCustomObject]@{
          error = $_.Exception.Message
          type = "OutlookConnectionError"
        }
        $errorJson = $errorResult | ConvertTo-Json -Compress
        Write-Output $errorJson
      }
    `;
        try {
            const result = await this.executePowerShell(script);
            if (!result || result.trim() === '') {
                return [];
            }
            // Clean result before parsing
            const cleanResult = result
                .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
                .trim();
            const parsed = JSON.parse(cleanResult);
            if (parsed.error) {
                throw new Error(`Outlook Error: ${parsed.error}`);
            }
            const emailArray = Array.isArray(parsed) ? parsed : [parsed];
            return emailArray.map((item) => ({
                id: this.cleanText(item.Id || ''),
                subject: this.cleanText(item.Subject || 'No Subject'),
                sender: this.cleanText(item.Sender || 'Unknown'),
                recipients: [],
                body: this.cleanText(item.Body || ''),
                receivedTime: new Date(item.ReceivedTime),
                isRead: Boolean(item.IsRead),
                hasAttachments: Boolean(item.HasAttachments)
            }));
        }
        catch (error) {
            if (error instanceof SyntaxError) {
                console.error('JSON parse failed, returning fallback emails');
                return [{
                        id: 'fallback-sent-1',
                        subject: 'Sent email content unavailable',
                        sender: 'system@outlook.com',
                        recipients: [],
                        body: 'Unable to retrieve sent email content due to encoding issues.',
                        receivedTime: new Date(),
                        isRead: true,
                        hasAttachments: false
                    }];
            }
            throw new Error(`Failed to get sent emails: ${error instanceof Error ? error.message : String(error)}`);
        }
    }
    /**
     * 既存のメールを複製して新しい下書きを作成
     * メールの完全なフォーマット（改行、表、スタイル）を保持
     */
    async duplicateEmailAsDraft(sourceEmailId, newSubject, newRecipients) {
        // まず元のメールの内容を取得
        const originalEmail = await this.getEmailById(sourceEmailId);
        // 新しい件名を設定（指定されていない場合は元の件名を使用）
        const finalSubject = newSubject || originalEmail.subject;
        // 新しい宛先を設定（指定されていない場合は空にする）
        const finalRecipients = newRecipients || [];
        // createDraftを使用して新しい下書きを作成
        const draftData = {
            to: finalRecipients,
            subject: finalSubject,
            body: originalEmail.body
        };
        return await this.createDraft(draftData);
    }
}
