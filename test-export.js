const { exec } = require('child_process');

function testEmailExport(accountName = "Google", folderName = "INBOX", emailCount = 5) {
  const script = `
    tell application "Mail"
      try
        set targetAccount to account \"${accountName}\"
        set targetFolder to mailbox \"${folderName}\" of targetAccount
        set emailList to messages of targetFolder
        set emailCount to count of emailList
        
        log "Found " & emailCount & " emails in ${folderName}"
        
        if emailCount > 0 then
          repeat with i from 1 to ${emailCount}
            if i ≤ emailCount then
              set currentEmail to item i of emailList
              
              -- Robust extraction of sender (From)
              set emailFrom to ""
              try
                set senderName to sender of currentEmail as string
                set senderAddress to address of sender of currentEmail as string
                set emailFrom to senderName & " <" & senderAddress & ">"
              end try

              -- Robust extraction of recipients (To)
              set emailTo to ""
              try
                set recipientList to to recipients of currentEmail
                set toAddresses to {}
                repeat with r in recipientList
                  set end of toAddresses to (name of r as string) & " <" & (address of r as string) & ">"
                end repeat
                set emailTo to (toAddresses as string)
              end try

              set emailSubject to subject of currentEmail as string
              set emailDate to date received of currentEmail as string
              set emailContent to content of currentEmail as string

              set emlText to "From: " & emailFrom & return & ¬
                  "To: " & emailTo & return & ¬
                  "Subject: " & emailSubject & return & ¬
                  "Date: " & emailDate & return & return & ¬
                  emailContent

              set fileName to "test_email_" & i & ".eml"
              set filePath to "/tmp/" & fileName

              set fileRef to open for access file filePath with write permission
              write emlText to fileRef
              close access fileRef
            end if
          end repeat
          return "Exported " & ${emailCount} & " emails"
        else
          return "No emails found in ${folderName}"
        end if
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;
  
  console.log(`Testing export from ${accountName}/${folderName}...`);
  
  exec(`osascript -e '${script}'`, (error, stdout, stderr) => {
    if (error) {
      console.error('Export test failed:', error);
      return;
    }
    console.log('Export test result:', stdout.trim());
  });
}

// Test with different folders
console.log('Testing email export functionality...\n');

// Test INBOX
testEmailExport("Google", "INBOX", 3);

setTimeout(() => {
  // Test Sent Mail
  testEmailExport("Google", "Sent Mail", 2);
}, 2000);

setTimeout(() => {
  // Test with iCloud account
  testEmailExport("iCloud", "INBOX", 2);
}, 4000); 