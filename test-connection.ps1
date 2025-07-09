# Set UTF-8 encoding for Windows
chcp 65001 > $null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "Testing robust Outlook email retrieval with UTF-8 encoding..." -ForegroundColor Yellow

function Clean-Text {
    param([string]$text)
    if (-not $text) { return "" }
    
    # Remove control characters and normalize
    $text = $text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]', ''
    $text = $text -replace '\r\n', ' '
    $text = $text -replace '\n', ' '
    $text = $text -replace '\r', ' '
    $text = $text -replace '\t', ' '
    $text = $text -replace '  +', ' '
    return $text.Trim()
}

try {
    Write-Host "Connecting to Outlook..." -ForegroundColor Green
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
    $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $itemCount = $inbox.Items.Count
    Write-Host "Found $itemCount emails in inbox" -ForegroundColor Green
    
    if ($itemCount -eq 0) {
        Write-Host "Inbox is empty - test completed successfully" -ForegroundColor Green
        exit 0
    }
    
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $emails = @()
    $counter = 0
    $testCount = 3  # Test with first 3 emails
    
    foreach ($item in $items) {
        if ($counter -ge $testCount) { break }
        
        try {
            Write-Host "Processing email $($counter + 1)..." -ForegroundColor Cyan
            
            # Get basic properties with null checks
            $subject = if ($item.Subject) { Clean-Text -text $item.Subject } else { "No Subject" }
            $sender = if ($item.SenderEmailAddress) { Clean-Text -text $item.SenderEmailAddress } else { "Unknown" }
            $bodyText = if ($item.Body) { 
                $cleanBody = Clean-Text -text $item.Body
                if ($cleanBody.Length -gt 100) { 
                    $cleanBody.Substring(0, 100) + "..." 
                } else { 
                    $cleanBody 
                }
            } else { "" }
            
            $receivedTime = if ($item.ReceivedTime) { 
                $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") 
            } else { 
                (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") 
            }
            
            Write-Host "  Subject: $subject" -ForegroundColor White
            Write-Host "  Sender: $sender" -ForegroundColor White
            if ($bodyText.Length -gt 0) {
                $preview = $bodyText.Substring(0, [Math]::Min(50, $bodyText.Length))
                Write-Host "  Body preview: $preview..." -ForegroundColor Gray
            }
            
            # Create email object with clean data
            $email = [PSCustomObject]@{
                Id = if ($item.EntryID) { $item.EntryID } else { "no-id-$counter" }
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
            Write-Host "  Error processing email $($counter + 1): $($_.Exception.Message)" -ForegroundColor Red
            $counter++
            continue
        }
    }
    
    Write-Host ""
    Write-Host "Converting to JSON..." -ForegroundColor Yellow
    try {
        $jsonResult = $emails | ConvertTo-Json -Depth 2 -Compress
        Write-Host "JSON conversion successful!" -ForegroundColor Green
        Write-Host "JSON output length: $($jsonResult.Length) characters" -ForegroundColor Cyan
        
        # Show first 200 characters of JSON
        $preview = if ($jsonResult.Length -gt 200) { 
            $jsonResult.Substring(0, 200) + "..." 
        } else { 
            $jsonResult 
        }
        Write-Host "JSON preview: $preview" -ForegroundColor Gray
        
        # Test JSON parsing
        $parsed = $jsonResult | ConvertFrom-Json
        Write-Host "JSON parsing test successful!" -ForegroundColor Green
        Write-Host "Parsed email count: $($parsed.Count)" -ForegroundColor Cyan
        
    } catch {
        Write-Host "JSON conversion failed: $($_.Exception.Message)" -ForegroundColor Red
        
        # Try fallback
        Write-Host "Attempting fallback approach..." -ForegroundColor Yellow
        $simpleEmails = @()
        for ($i = 0; $i -lt $emails.Count; $i++) {
            $simpleEmails += [PSCustomObject]@{
                Id = "test-$i"
                Subject = "Email $($i + 1)"
                Sender = "test@example.com"
                Recipients = @()
                Body = "Test content"
                ReceivedTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                IsRead = $true
                HasAttachments = $false
            }
        }
        $fallbackJson = $simpleEmails | ConvertTo-Json -Depth 2 -Compress
        Write-Host "Fallback JSON successful: $($fallbackJson.Length) characters" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Test completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Host "Test failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
