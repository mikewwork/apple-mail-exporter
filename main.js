const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { exec, execFile } = require('child_process');
const fs = require('fs-extra');
const EmlParser = require('eml-parser');
const { simpleParser } = require('mailparser');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const puppeteer = require('puppeteer');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
    icon: path.join(__dirname, 'assets/icon.png'),
    title: 'Apple Mail Exporter'
  });

  mainWindow.loadFile('index.html');

  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools();
  }
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

// IPC Handlers
ipcMain.handle('get-mail-accounts', async () => {
  try {
    // AppleScript as array of lines for execFile
    const scriptLines = [
      'tell application "Mail"',
      'set output to ""',
      'repeat with acc in accounts',
      'set output to output & "Account: " & (name of acc) & linefeed',
      'repeat with mbox in mailboxes of acc',
      'set output to output & "  Folder: " & (name of mbox) & linefeed',
      'end repeat',
      'end repeat',
      'return output',
      'end tell'
    ];
    const args = scriptLines.flatMap(line => ['-e', line]);
    return new Promise((resolve, reject) => {
      execFile('osascript', args, (error, stdout, stderr) => {
        if (error) {
          console.error('AppleScript error:', error);
          reject(error);
          return;
        }
        try {
          const accounts = parseAccountsAndFolders(stdout);
          resolve(accounts);
        } catch (parseError) {
          reject(parseError);
        }
      });
    });
  } catch (error) {
    console.error('Error getting mail accounts:', error);
    throw error;
  }
});

function parseAccountsAndFolders(output) {
  const lines = output.split(/\r?\n/).filter(Boolean);
  const accounts = [];
  let currentAccount = null;
  for (const line of lines) {
    if (line.startsWith('Account: ')) {
      if (currentAccount) accounts.push(currentAccount);
      currentAccount = { name: line.replace('Account: ', '').trim(), folders: [] };
    } else if (line.startsWith('  Folder: ')) {
      if (currentAccount) {
        currentAccount.folders.push({ name: line.replace('  Folder: ', '').trim() });
      }
    }
  }
  if (currentAccount) accounts.push(currentAccount);
  return accounts;
}

ipcMain.handle('select-output-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
    title: 'Select Output Folder for PDF Files'
  });

  console.log('Dialog result:', result);

  if (!result.canceled) {
    console.log('Selected folder:', result.filePaths[0]);
    return result.filePaths[0];
  }
  return null;
});

ipcMain.handle('openFolder', async (event, folderPath) => {
  const { shell } = require('electron');
  await shell.openPath(folderPath);
});

function extractClientName(body) {
  const signOffs = [
    'regards,', 'best regards,', 'thanks,', 'thank you,', 'sincerely,', 'cheers,', 'kind regards,'
  ];
  const lines = body.split(/\r?\n/).map(line => line.trim());
  for (let i = 0; i < lines.length; i++) {
    const lower = lines[i].toLowerCase();
    if (signOffs.some(sign => lower.startsWith(sign))) {
      for (let j = i + 1; j < lines.length; j++) {
        if (lines[j]) return lines[j];
      }
    }
  }
  return '';
}

function extractInfoFromEmail(parsed) {
  const from = parsed.from?.value?.[0]?.address || '';
  const date = parsed.date || '';
  const body = parsed.text || '';
  // Extract base URL (protocol + domain)
  let website = '';
  const urlMatch = body.match(/https?:\/\/[^\s"']+/);
  if (urlMatch) {
    try {
      const url = new URL(urlMatch[0]);
      website = url.origin;
    } catch {}
  }
  const phoneMatch = body.match(/\+?\d[\d\s\-()]{7,}/g);
  const phone = phoneMatch ? phoneMatch[0] : '';
  const clientName = extractClientName(body);
  return { from, date, clientName, website, phone };
}

function logDebug(msg, outputPath) {
  const logFile = path.join(outputPath, 'export_debug.log');
  fs.appendFileSync(logFile, `[${new Date().toISOString()}] ${msg}\n`);
}

async function emlToPdfWithPuppeteer(emlPath, pdfPath) {
  const emlContent = await fs.readFile(emlPath, 'utf8');
  const parsed = await simpleParser(emlContent);
  // Build HTML for the email
  const html = `
    <html><head><meta charset='utf-8'><title>${parsed.subject || ''}</title></head><body>
    <div style='font-family:sans-serif;'>
      <h2>${parsed.subject || ''}</h2>
      <div><b>From:</b> ${parsed.from?.text || ''}</div>
      <div><b>To:</b> ${parsed.to?.text || ''}</div>
      <div><b>Date:</b> ${parsed.date || ''}</div>
      <hr/>
      <div>${parsed.html || parsed.textAsHtml || `<pre>${parsed.text || ''}</pre>`}</div>
    </div>
    </body></html>
  `;
  const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle0' });
  await page.pdf({ path: pdfPath, format: 'A4', printBackground: true });
  await browser.close();
}

ipcMain.handle('export-emails', async (event, { accountName, folderName, outputPath, emailCount = 10 }) => {
  try {
    await fs.ensureDir(outputPath);
    const pdfFiles = [];
    const timestampCounts = {};
    const clientInfoRows = [];
    logDebug('Starting export-emails', outputPath);
    // 1. Get all message count in folder
    const getCountScriptLines = [
      `tell application "Mail"`,
      `set targetAccount to account \"${accountName.replace(/"/g, '\\"')}\"`,
      `set targetFolder to mailbox \"${folderName.replace(/"/g, '\\"')}\" of targetAccount`,
      `set emailList to messages of targetFolder`,
      `return count of emailList`,
      `end tell`
    ];
    const countArgs = getCountScriptLines.flatMap(line => ['-e', line]);
    const totalCount = parseInt(await new Promise((resolve, reject) => {
      execFile('osascript', countArgs, (error, stdout, stderr) => {
        if (error) resolve('0');
        else resolve(stdout.trim());
      });
    }), 10);
    // 2. Loop over the last N messages (most recent)
    const startIdx = Math.max(1, totalCount - emailCount + 1);
    let matchCount = 0;
    for (let i = totalCount; i >= startIdx; --i) {
      // Get date for this email
      const getDateScriptLines = [
        `tell application "Mail"`,
        `set targetAccount to account \"${accountName.replace(/"/g, '\\"')}\"`,
        `set targetFolder to mailbox \"${folderName.replace(/"/g, '\\"')}\" of targetAccount`,
        `set emailList to messages of targetFolder`,
        `set currentEmail to item ${i} of emailList`,
        `set emailDate to date received of currentEmail`,
        `set isoDate to (year of emailDate as string) & "-" & text -2 thru -1 of ("0" & (month of emailDate as integer)) & "-" & text -2 thru -1 of ("0" & day of emailDate as integer) & " " & text -2 thru -1 of ("0" & hours of emailDate as integer) & ":" & text -2 thru -1 of ("0" & minutes of emailDate as integer) & ":" & text -2 thru -1 of ("0" & seconds of emailDate as integer)`,
        `return isoDate`,
        `end tell`
      ];
      const dateArgs = getDateScriptLines.flatMap(line => ['-e', line]);
      let emailDateStr = '';
      try {
        emailDateStr = await new Promise((resolve, reject) => {
          execFile('osascript', dateArgs, (error, stdout, stderr) => {
            if (error) resolve('');
            else resolve(stdout.trim());
          });
        });
      } catch { emailDateStr = ''; }
      let dateObj = emailDateStr ? new Date(emailDateStr.replace(' ', 'T')) : new Date();
      const pad = n => n.toString().padStart(2, '0');
      let timestamp = `${dateObj.getFullYear()}-${pad(dateObj.getMonth()+1)}-${pad(dateObj.getDate())}-${pad(dateObj.getHours())}-${pad(dateObj.getMinutes())}-${pad(dateObj.getSeconds())}`;
      let baseTimestamp = timestamp;
      let uniqueTimestamp = baseTimestamp;
      if (timestampCounts[baseTimestamp] === undefined) {
        timestampCounts[baseTimestamp] = 1;
      } else {
        timestampCounts[baseTimestamp] += 1;
        uniqueTimestamp = `${baseTimestamp}-${timestampCounts[baseTimestamp]}`;
      }
      const emlPath = path.join(outputPath, `${uniqueTimestamp}.eml`);
      const pdfPath = path.join(outputPath, `${uniqueTimestamp}.pdf`);
      logDebug(`Exporting EML for email ${i} to ${emlPath}`, outputPath);
      // Export EML
      const exportScriptLines = [
        `tell application "Mail"`,
        `set exportPath to "${outputPath.replace(/"/g, '\\"')}"`,
        `set targetAccount to account \"${accountName.replace(/"/g, '\\"')}\"`,
        `set targetFolder to mailbox \"${folderName.replace(/"/g, '\\"')}\" of targetAccount`,
        `set emailList to messages of targetFolder`,
        `set currentEmail to item ${i} of emailList`,
        `set fileName to "${uniqueTimestamp}.eml"`,
        `set filePath to exportPath & "/" & fileName`,
        `set emlSource to source of currentEmail`,
        `set emlBase64 to do shell script "echo " & quoted form of emlSource & " | base64"`,
        `do shell script "echo " & quoted form of emlBase64 & " | base64 -D > " & quoted form of filePath`,
        `end tell`
      ];
      const args = exportScriptLines.flatMap(line => ['-e', line]);
      await new Promise((resolve, reject) => {
        execFile('osascript', args, (error, stdout, stderr) => {
          if (error) {
            reject(error);
            return;
          }
          resolve();
        });
      });
      logDebug(`Exported EML for email ${i}`, outputPath);
      // 1. Read EML and extract info for CSV
      try {
        const emlContent = await fs.readFile(emlPath, 'utf8');
        const parsed = await simpleParser(emlContent);
        const info = extractInfoFromEmail(parsed);
        clientInfoRows.push(info);
      } catch (parseErr) {
        logDebug(`Error parsing EML for CSV: ${parseErr.stack || parseErr}`, outputPath);
      }
      // 2. Convert EML to PDF
      try {
        await emlToPdfWithPuppeteer(emlPath, pdfPath);
        await fs.remove(emlPath);
        pdfFiles.push(path.basename(pdfPath));
        logDebug(`Converted PDF for email ${i}: ${pdfPath}`, outputPath);
      } catch (fileError) {
        // Continue with next email even if one fails
        logDebug(`Error processing email ${i}: ${fileError.stack || fileError}`, outputPath);
      }
      matchCount++;
      // Send progress update
      const percent = Math.round((matchCount / emailCount) * 100);
      event.sender.send('export-progress', {
        type: 'progress',
        percent,
        status: `Exported ${matchCount} of ${emailCount} emails to PDF`
      });
    }
    if (matchCount === 0) {
      logDebug('No emails found for export.', outputPath);
      return { success: true, pdfFiles: [] };
    }
    if (clientInfoRows.length > 0) {
      logDebug('Writing client_info.csv', outputPath);
      const csvWriter = createCsvWriter({
        path: path.join(outputPath, 'client_info.csv'),
        header: [
          {id: 'from', title: 'From'},
          {id: 'date', title: 'Date'},
          {id: 'clientName', title: 'ClientName'},
          {id: 'website', title: 'Website'},
          {id: 'phone', title: 'Phone'}
        ]
      });
      await csvWriter.writeRecords(clientInfoRows);
    }
    logDebug('Export complete', outputPath);
    return { success: true, pdfFiles };
  } catch (error) {
    console.error('Error exporting emails:', error);
    throw error;
  }
});

ipcMain.handle('openPDF', async (event, pdfPath) => {
  const { shell } = require('electron');
  await shell.openPath(pdfPath);
});



 