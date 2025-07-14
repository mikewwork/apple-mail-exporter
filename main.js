const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { exec } = require('child_process');
const { promisify } = require('util');
const fs = require('fs');
const path = require('path');
const EmlParser = require('eml-parser');
const EmailReplyParser = require('email-reply-parser');
const emailRegex = require('email-regex-safe');

const execAsync = promisify(exec);

function createWindow() {
  const win = new BrowserWindow({
    width: 1000,
    height: 700,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  win.loadFile('index.html');
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// Handle folder selection dialog
ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openDirectory'],
    title: 'Select Output Folder'
  });
  
  if (!result.canceled && result.filePaths.length > 0) {
    return result.filePaths[0];
  }
  return null;
});

// Handle scanning mail folders using AppleScript
ipcMain.handle('scan-mail-folders', async () => {
  try {
    const script = `
      tell application "Mail"
        set allMailboxes to {}
        
        -- Get local mailboxes (On My Mac) with full paths
        repeat with currentMailbox in mailboxes
          if account of currentMailbox is missing value then
            set mailboxPath to name of currentMailbox
            set end of allMailboxes to mailboxPath
            
            -- Also get submailboxes with full path
            try
              repeat with subMailbox in mailboxes of currentMailbox
                set subMailboxPath to mailboxPath & "/" & name of subMailbox
                set end of allMailboxes to subMailboxPath
                
                -- Get sub-submailboxes (3 levels deep)
                try
                  repeat with subSubMailbox in mailboxes of subMailbox
                    set subSubMailboxPath to subMailboxPath & "/" & name of subSubMailbox
                    set end of allMailboxes to subSubMailboxPath
                  end repeat
                end try
              end repeat
            end try
          end if
        end repeat
        
        return allMailboxes
      end tell
    `;
    
    const { stdout } = await execAsync(`osascript -e '${script}'`);
    
    const folders = stdout.trim()
      .split(', ')
      .filter(folder => folder.length > 0)
      .map(folder => folder.replace(/"/g, ''))
      .sort();
    
    if (folders.length === 0) {
      throw new Error('No local mailboxes found. Make sure you have "On My Mac" mailboxes set up in Apple Mail.');
    }
    
    return folders;
    
  } catch (error) {
    console.error('Error scanning mail folders:', error);
    throw error;
  }
});

// Function to generate CSV from .eml files
async function generateCSVFromEmails(emailsFolder, csvFilePath, event) {
  try {
    const files = fs.readdirSync(emailsFolder).filter(file => file.endsWith('.eml'));
    
    if (files.length === 0) {
      console.log('No .eml files found to process');
      return;
    }
    
    const csvRows = ['Date,Sender Email,Subject,Name,Company,Phone,Domain,Filename,Content'];
    
    // Send initial progress update
    if (event) {
      event.sender.send('progress-update', 50, `Processing ${files.length} emails...`);
    }
    
    // Calculate batch size for progress updates
    // For small counts (≤50): update every email
    // For medium counts (51-500): update every 10 emails
    // For large counts (>500): update every 50 emails
    let batchSize;
    if (files.length <= 50) {
      batchSize = 1; // Update every email for small counts
    } else if (files.length <= 500) {
      batchSize = 10; // Update every 10 emails for medium counts
    } else {
      batchSize = 50; // Update every 50 emails for large counts
    }
    let lastProgressUpdate = 0;
    
    for (let i = 0; i < files.length; i++) {
      const fileName = files[i];
      const filePath = path.join(emailsFolder, fileName);
      const emailContent = fs.readFileSync(filePath, 'utf8');
      
      // Parse email content
      const emailData = await parseEmailContent(emailContent, fileName);
      
      // Add to CSV rows
      csvRows.push([
        emailData.timestamp,
        emailData.fromEmail,
        emailData.subject,
        emailData.senderName,
        emailData.companyName,
        emailData.contactPhone,
        emailData.websiteUrl,
        emailData.fileName,
        emailData.content
      ].map(field => `"${field.replace(/"/g, '""')}"`).join(','));
      
      // Update progress in batches to avoid overwhelming the UI
      if (event && (i + 1) % batchSize === 0) {
        const progress = 50 + Math.floor((i + 1) / files.length * 40); // 50-90%
        const processedCount = Math.min(i + 1, files.length);
        event.sender.send('progress-update', progress, `Processed ${processedCount} of ${files.length} emails...`);
        lastProgressUpdate = progress;
      }
    }
    
    // Send final processing update if we haven't sent one recently
    if (event && lastProgressUpdate < 90) {
      event.sender.send('progress-update', 90, `Processed all ${files.length} emails...`);
    }
    
    // Write CSV file
    if (event) {
      event.sender.send('progress-update', 90, 'Creating CSV file...');
    }
    fs.writeFileSync(csvFilePath, csvRows.join('\n'), 'utf8');
    console.log(`CSV file created with ${files.length} email records`);
    
  } catch (error) {
    console.error('Error generating CSV:', error);
    throw error;
  }
}

// Function to parse email content and extract required data
async function parseEmailContent(emailContent, fileName) {
  try {
    const eml = new EmlParser(Buffer.from(emailContent));
    const data = await eml.parseEml();
    
    // Extract basic email data
    const fromEmail = data.from?.text || data.from?.value?.[0]?.address || 'Unknown';
    const fromHeader = data.from?.text || '';
    const subject = data.subject || '';
    const emailDate = data.date ? new Date(data.date) : new Date();
    const timestamp = emailDate.toISOString().replace('T', ' ').substring(0, 19);
    
    // Get text body (simple approach like your example)
    const textBody = data.text || data.htmlAsText || 'No content available';
    
    // Clean up content for CSV (convert line breaks to spaces)
    const content = textBody
      .replace(/\r\n/g, ' ')
      .replace(/\r/g, ' ')
      .replace(/\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    
    // Extract website URLs from content
    const urlRegex = /(https?:\/\/[^\s<>"]+|www\.[^\s<>"]+)/gi;
    const urls = content.match(urlRegex) || [];
    const websiteUrl = urls.length > 0 ? extractDomain(urls[0]) : '';
    
    // Use email-reply-parser for signature extraction
    const corporateInfo = parseCorporateEmail(textBody, fromHeader);
    
    return {
      fromEmail,
      fileName,
      subject,
      websiteUrl,
      senderName: corporateInfo.senderName || '',
      companyName: corporateInfo.company || '',
      contactPhone: corporateInfo.phone || '',
      timestamp,
      content
    };
    
  } catch (error) {
    console.error('Error parsing email:', error);
    return {
      fromEmail: 'Error',
      fileName,
      subject: '',
      websiteUrl: '',
      senderName: '',
      companyName: '',
      contactPhone: '',
      timestamp: new Date().toISOString().replace('T', ' ').substring(0, 19),
      content: 'Error parsing email content'
    };
  }
}

// Helper function to extract domain from URL
function extractDomain(url) {
  try {
    let domain = url;
    if (url.startsWith('http://')) {
      domain = url.substring(7);
    } else if (url.startsWith('https://')) {
      domain = url.substring(8);
    } else if (url.startsWith('www.')) {
      domain = url.substring(4);
    }
    
    // Remove path and query parameters
    const slashIndex = domain.indexOf('/');
    if (slashIndex > 0) {
      domain = domain.substring(0, slashIndex);
    }
    
    return domain;
  } catch (error) {
    return '';
  }
}

// Superior email parsing functions using email-reply-parser
function extractSenderName(fromHeader) {
  // Example: "Jane Doe <jane@acme.com>"
  const nameMatch = fromHeader.match(/^([^<]+)</);
  return nameMatch ? nameMatch[1].trim() : fromHeader;
}

function extractPhone(signatureBlock) {
  if (!signatureBlock) return null;
  const phoneMatch = signatureBlock.match(
    /(\+?\d{1,3}[-.\s]?)?(\(?\d{2,4}\)?[-.\s]?)?\d{3,4}[-.\s]?\d{4}/
  );
  return phoneMatch ? phoneMatch[0].trim() : null;
}

function extractCompanyName(signatureBlock) {
  if (!signatureBlock) return null;
  const lines = signatureBlock.trim().split('\n').map(l => l.trim()).filter(Boolean);

  const signoffIndex = lines.findIndex(line =>
    /^(Regards|Thanks|Best|Sincerely)[\s,]*$/i.test(line)
  );

  // Assume company name is 2 lines after the sign-off
  if (signoffIndex !== -1 && lines[signoffIndex + 2]) {
    return lines[signoffIndex + 2];
  }

  // Fallback: last non-name line
  return lines.length > 1 ? lines.slice(-1)[0] : null;
}

function parseCorporateEmail(rawEmail, fromHeader = null) {
  try {
    const parser = new EmailReplyParser();
    const parsed = parser.read(rawEmail);
    const fragments = parsed.getFragments();
    const signature = fragments.filter(f => f.isSignature()).map(f => f.getContent()).join('\n');

    return {
      senderName: fromHeader ? extractSenderName(fromHeader) : null,
      phone: extractPhone(signature),
      company: extractCompanyName(signature)
    };
  } catch (error) {
    console.error('Error parsing corporate email:', error);
    return {
      senderName: null,
      phone: null,
      company: null
    };
  }
}



// Handle exporting emails as .eml files using AppleScript
ipcMain.handle('export-emails', async (event, selectedFolder, outputPath) => {
  try {
    // Create emails subfolder in the output directory
    const emailsFolder = path.join(outputPath, 'emails');
    if (!fs.existsSync(emailsFolder)) {
      fs.mkdirSync(emailsFolder, { recursive: true });
    }
    
    // Create CSV file path
    const csvFilePath = path.join(outputPath, 'contacts.csv');
    
    const script = `
      tell application "Mail"
        set exportedCount to 0
        
        -- Get the selected mailbox
        set selectedMailbox to missing value
        
        -- Parse the folder path to handle nested mailboxes
        set folderParts to words of "${selectedFolder}"
        
        -- Navigate to the mailbox
        repeat with i from 1 to count of folderParts
          set folderName to item i of folderParts
          if i is 1 then
            set selectedMailbox to mailbox folderName
          else
            set selectedMailbox to mailbox folderName of selectedMailbox
          end if
        end repeat
        
        log "Processing mailbox: ${selectedFolder}"
        
        -- Process main mailbox
        set allMessages to messages of selectedMailbox
        set messageCount to count of allMessages
        
        if messageCount > 0 then
          log "Found " & messageCount & " messages in ${selectedFolder}"
          
          repeat with i from 1 to messageCount
            set currentMessage to item i of allMessages
            
            -- Get message details
            set messageSender to sender of currentMessage
            set messageDate to date received of currentMessage
            set messageSource to source of currentMessage
            
            -- Create timestamp for filename
            set timestamp to do shell script "date +%Y%m%d_%H%M%S"
            set fileName to "email_" & timestamp & "_" & exportedCount + i & ".eml"
            set filePath to "${emailsFolder}/" & fileName
            
            -- Write raw message source directly to .eml file
            try
              do shell script "echo " & quoted form of messageSource & " > " & quoted form of filePath
              log "Exported: " & fileName
            on error writeError
              log "Error writing file: " & writeError
            end try
          end repeat
          
          set exportedCount to exportedCount + messageCount
        end if
        
        -- Process submailboxes
        try
          repeat with subMailbox in mailboxes of selectedMailbox
            set subMailboxName to name of subMailbox
            log "Processing submailbox: " & subMailboxName
            
            -- Get all messages in this submailbox
            set subMessages to messages of subMailbox
            set subMessageCount to count of subMessages
            
            if subMessageCount > 0 then
              log "Found " & subMessageCount & " messages in " & subMailboxName
              
              repeat with j from 1 to subMessageCount
                set currentMessage to item j of subMessages
                
                -- Get message details
                set messageSender to sender of currentMessage
                set messageDate to date received of currentMessage
                set messageSource to source of currentMessage
                
                -- Create timestamp for filename
                set timestamp to do shell script "date +%Y%m%d_%H%M%S"
                set fileName to "email_" & timestamp & "_" & exportedCount + j & ".eml"
                set filePath to "${emailsFolder}/" & fileName
                
                -- Write raw message source directly to .eml file
                try
                  do shell script "echo " & quoted form of messageSource & " > " & quoted form of filePath
                  log "Exported: " & fileName
                on error writeError
                  log "Error writing file: " & writeError
                end try
              end repeat
              
              set exportedCount to exportedCount + subMessageCount
            end if
          end repeat
        end try
        
        log "Total exported: " & exportedCount
        return exportedCount
      end tell
    `;
    
    const { stdout, stderr } = await execAsync(`osascript -e '${script}'`);
    
    if (stderr) {
      console.log('AppleScript debug output:', stderr);
    }
    
    const exportedCount = parseInt(stdout.trim()) || 0;
    
    // Now process the .eml files to generate CSV
    if (exportedCount > 0) {
      try {
        await generateCSVFromEmails(emailsFolder, csvFilePath, event);
      } catch (csvError) {
        console.error('Error generating CSV:', csvError);
        // Don't fail the entire operation if CSV generation fails
      }
    }
    
    return {
      success: true,
      messageCount: exportedCount,
      message: `Successfully exported ${exportedCount} emails as .eml files from "${selectedFolder}" and all subfolders to "${emailsFolder}". Contact information saved to "contacts.csv" in the output folder.`
    };
    
  } catch (error) {
    console.error('Error exporting emails:', error);
    throw error;
  }
});



 