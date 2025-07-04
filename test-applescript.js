const { exec } = require('child_process');

// Test AppleScript connection to Mail
function testMailConnection() {
  const script = `
    tell application "Mail"
      try
        set accountCount to count of accounts
        return "Mail is running with " & accountCount & " accounts"
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;
  
  exec(`osascript -e '${script}'`, (error, stdout, stderr) => {
    if (error) {
      console.error('AppleScript test failed:', error);
      return;
    }
    console.log('AppleScript test result:', stdout.trim());
  });
}

// Test email export functionality
function testEmailExport() {
  const script = `
    tell application "Mail"
      try
        if (count of accounts) > 0 then
          set firstAccount to account 1
          set accountName to name of firstAccount
          return "First account: " & accountName
        else
          return "No accounts found"
        end if
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;
  
  exec(`osascript -e '${script}'`, (error, stdout, stderr) => {
    if (error) {
      console.error('Email export test failed:', error);
      return;
    }
    console.log('Email export test result:', stdout.trim());
  });
}

console.log('Testing AppleScript functionality...');
testMailConnection();
setTimeout(testEmailExport, 1000); 