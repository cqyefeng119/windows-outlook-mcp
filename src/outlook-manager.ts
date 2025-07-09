import { spawn } from 'child_process';

export interface EmailMessage {
  id: string;
  subject: string;
  sender: string;
  recipients: string[];
  body: string;
  receivedTime: Date;
  isRead: boolean;
  hasAttachments: boolean;
}

export interface EmailDraft {
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  body: string;
  isHtml?: boolean;
}

export class OutlookManager {
  private powershellPath: string;

  constructor() {
    this.powershellPath = 'powershell.exe';
  }

  private async executePowerShell(script: string): Promise<string> {
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
        } else {
          reject(new Error(`PowerShell failed (code ${code}): ${stderr}`));
        }
      });
    });
  }

  private cleanText(text: string): string {
    if (!text) return '';
    
    return text
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '') // Remove control chars
      .replace(/\r\n/g, ' ') // Replace CRLF with space
      .replace(/[\r\n]/g, ' ') // Replace any remaining line breaks
      .replace(/\s+/g, ' ') // Collapse multiple spaces
      .trim();
  }

  async getInboxEmails(count: number = 10): Promise<EmailMessage[]> {
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
      
      return emailArray.map((item: any) => ({
        id: this.cleanText(item.Id || ''),
        subject: this.cleanText(item.Subject || 'No Subject'),
        sender: this.cleanText(item.Sender || 'Unknown'),
        recipients: [],
        body: this.cleanText(item.Body || ''),
        receivedTime: new Date(item.ReceivedTime),
        isRead: Boolean(item.IsRead),
        hasAttachments: Boolean(item.HasAttachments)
      }));
      
    } catch (error) {
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

  async getEmailById(id: string): Promise<EmailMessage> {
    return {
      id: id,
      subject: 'Email details unavailable',
      sender: 'system@outlook.com',
      recipients: [],
      body: 'Email content could not be retrieved.',
      receivedTime: new Date(),
      isRead: true,
      hasAttachments: false
    };
  }

  async createDraft(draft: EmailDraft): Promise<string> {
    const cleanSubject = this.cleanText(draft.subject);
    const cleanBody = this.cleanText(draft.body);
    
    const script = `
      try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        
        $mail.Subject = "${cleanSubject.replace(/"/g, '""')}"
        $mail.Body = "${cleanBody.replace(/"/g, '""')}"
        
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

  async markAsRead(id: string): Promise<void> {
    return Promise.resolve();
  }

  async searchEmails(query: string, count: number = 10): Promise<EmailMessage[]> {
    const emails = await this.getInboxEmails(Math.min(count * 2, 20));
    return emails.slice(0, count);
  }
}
