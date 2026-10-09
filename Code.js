// ==========================================
//             CLOUD CONFIGURATION
// ==========================================
const API_KEY       = PropertiesService.getScriptProperties().getProperty('UNIFI_API_KEY') || '';
const DEPLOYMENT_ID = PropertiesService.getScriptProperties().getProperty('UNIFI_DEPLOYMENT_ID') || '';
const SITES_JSON    = PropertiesService.getScriptProperties().getProperty('UNIFI_SITES_CONFIG') || '';
// ==========================================

/**
 * AUTOMATION TARGET: Attach this to your weekly background timer trigger.
 * Runs a complete network-wide top-up loop and emails every configured site.
 */
function weeklyAutoReplenishAndEmail() {
  console.log("Starting master multi-site automated replenishment execution...");
  
  const sites = getParsedSitesConfig();
  if (sites.length === 0) return;

  const weeklySecureToken = getOrGenerateWeeklyToken(true);

  sites.forEach((site) => {
    try {
      console.log(`Processing ecosystem top-up for: ${site.siteName}`);
      
      // 1. SECURITY PURGE: Wipes the pool if your custom JSON property says true
      if (site.purgeBeforeSync === true) {
        console.warn(`SECURITY TRIGGER: Purge flag detected for ${site.siteName}. Wiping unused vouchers...`);
        purgeAllUnusedVouchersDirect(site);
      }
      
      // 2. Run the standard synchronization/top-up loop
      syncSingleSiteVouchers(site);
      
      // 3. Build the unique profile-neutral URL carrying both token and school parameters
      const escapedSiteName = encodeURIComponent(site.siteName);
      const webAppUrl = `https://script.google.com/macros/s/${DEPLOYMENT_ID}/exec?t=${weeklySecureToken}&site=${escapedSiteName}`;
      
      const subject = `🎫 GUEST WI-FI VOUCHERS REGENERATED: ${site.siteName}`;
      const body = `🎫 GUEST WI-FI VOUCHERS REGENERATED\n\n` +
                   `Hi there,\n\nThe weekly guest Wi-Fi vouchers for ${site.siteName} have been successfully updated.\n\nA fresh batch of 20 standard (8-Hour) and 5 extended (3-Day for Ofsted, etc) active passes is now ready to use.\n\nYou can view and directly print your new voucher tickets here:\n${webAppUrl}\n\n` +
                   `💡 Tip: While you can print these off, the page will automatically remove used vouchers in real-time so you can always click this link throughout the week to see exactly what's actually available!\n\n` +
                   `Please note: This link is unique to this batch and is only valid for this week. Please do not reply to this automated message.`;
      
      const htmlBody = `
        <div style="font-family: sans-serif; max-width: 600px; margin: auto; border: 1px solid #eee; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
          <div style="background: #007bff; padding: 25px; text-align: center;">
            <h2 style="color: white; margin: 0; letter-spacing: 1px;">🎫 GUEST WI-FI VOUCHERS REGENERATED</h2>
            <div style="color: #d1e7dd; font-size: 14px; margin-top: 5px; font-weight: bold;">${site.siteName}</div>
          </div>
          <div style="padding: 30px; color: #333; line-height: 1.5;">
            <p>Hi there,</p>
            <p>The weekly guest Wi-Fi vouchers for your location have been updated. A fresh batch of 25 passes (20 Standard & 5 Extended for Ofsted,etc) has been generated and is ready to be printed:</p>
            
            <div style="background: #f9f9f9; padding: 15px; border-left: 4px solid #007bff; margin: 15px 0;">
              <strong>School:</strong> ${site.siteName}<br>
              <strong>Total Active Passes:</strong> 25 Vouchers<br>
              <strong>Voucher Types:</strong> 20 x 8-Hour Passes | 5 x 3-Day Passes<br>
              <strong>Device Limit:</strong> Single Use (Per Token)
            </div>
            
            <p style="margin-top: 25px; margin-bottom: 25px; text-align: center;">
              <a href="${webAppUrl}" target="_blank" style="background: #28a745; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">🖨️ Open Printable Tickets Page</a>
            </p>

            <div style="background: #e0f2fe; padding: 15px; border-radius: 6px; border: 1px solid #bae6fd; font-size: 13px; color: #0369a1; margin-bottom: 25px; line-height: 1.6;">
              💡 <strong>Handy Tip:</strong> While you can print these off all at once, you can also keep this link handy! The page connects directly to our Wi-Fi controller, meaning used vouchers will automatically disappear from the list in real-time so you can check back anytime to see exactly what is still available.
            </div>
            
            <p style="font-size: 13px; color: #555;">If the button above does not open, you can copy and paste this link into your browser:<br>
            <span style="font-family: monospace; font-size: 11px; color: #0369a1; word-break: break-all;">${webAppUrl}</span></p>
            
            <div style="background: #fff9db; padding: 15px; border-radius: 6px; border: 1px solid #ffe066; font-size: 12px; color: #856404; margin-top: 25px;">
              <strong>Please note:</strong> This print link is only valid for this week. A new link will be sent automatically next week.
            </div>
          </div>
        </div>
      `;
      
      MailApp.sendEmail({ to: site.recipients, subject: subject, body: body, htmlBody: htmlBody, noReply: true });
      console.log(`Successfully dispatched update notifications for ${site.siteName}`);
    } catch (siteError) {
      console.error(`Critical failure compiling profile loop for ${site.siteName}: ${siteError.toString()}`);
    }
  });
  
  logToUiToast("Master automated synchronization completed across all schools!", "Batch Complete 🚀");
}

/**
 * CONTEXTUAL TRIGGER: Syncs and tops up codes ONLY for the school corresponding to the active tab.
 */
function syncActiveSheetTabVouchers() {
  const sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const sites = getParsedSitesConfig();
  const activeSite = sites.find(s => s.siteName === sheetName);
  
  if (!activeSite) {
    SpreadsheetApp.getUi().alert(
      "Operation Cancelled", 
      "The active spreadsheet tab name does not match any configured school profile in your JSON variables properties box.", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  logToUiToast(`Syncing ${activeSite.siteName} with UniFi hardware...`, "Processing 🔄");
  syncSingleSiteVouchers(activeSite);
  logToUiToast(`Sync complete for ${activeSite.siteName}!`, "Complete ✅");
}

/**
 * CONTEXTUAL TRIGGER: Performs a full top-up (including a purge if your flag is true) 
 * and dispatches the styled email layout ONLY for the school tab you are currently looking at.
 */
function syncAndEmailActiveTabSite() {
  const sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const sites = getParsedSitesConfig();
  const site = sites.find(s => s.siteName === sheetName);
  
  if (!site) {
    SpreadsheetApp.getUi().alert(
      "Operation Cancelled", 
      "The active spreadsheet tab name does not match any configured school profile in your JSON properties box.", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  if (!DEPLOYMENT_ID) {
    SpreadsheetApp.getUi().alert("Setup Required", "Please save your UNIFI_DEPLOYMENT_ID to Script Properties first.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  logToUiToast(`Running dedicated sync & email loop for ${site.siteName}...`, "Processing 🔄");

  try {
    if (site.purgeBeforeSync === true) {
      console.warn(`SECURITY TRIGGER: Manual single-site run caught purge flag for ${site.siteName}. Wiping vouchers...`);
      purgeAllUnusedVouchersDirect(site);
    }
    
    syncSingleSiteVouchers(site);
    
    const currentSecureToken = getOrGenerateWeeklyToken(false);
    const webAppUrl = `https://script.google.com/macros/s/${DEPLOYMENT_ID}/exec?t=${currentSecureToken}&site=${encodeURIComponent(site.siteName)}`;
    
    const subject = `🎫 GUEST WI-FI VOUCHERS REGENERATED: ${site.siteName}`;
    const body = `🎫 GUEST WI-FI VOUCHERS REGENERATED\n\n` +
                 `Hi there,\n\nThe guest Wi-Fi vouchers for ${site.siteName} have been successfully updated.\n\nA fresh batch of 20 standard (8-Hour) and 5 extended (3-Day) active passes is now ready to use.\n\nYou can view and directly print your new voucher tickets here:\n${webAppUrl}\n\n` +
                 `💡 Tip: While you can print these off, the page will automatically remove used vouchers in real-time so you can always click this link throughout the week to see exactly what's actually available!\n\n` +
                 `Please note: This link is unique to this batch and is only valid for this week. Please do not reply to this automated message.`;
    
    const htmlBody = `
      <div style="font-family: sans-serif; max-width: 600px; margin: auto; border: 1px solid #eee; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
        <div style="background: #007bff; padding: 25px; text-align: center;">
          <h2 style="color: white; margin: 0; letter-spacing: 1px;">🎫 GUEST WI-FI VOUCHERS REGENERATED</h2>
          <div style="color: #d1e7dd; font-size: 14px; margin-top: 5px; font-weight: bold;">${site.siteName}</div>
        </div>
        <div style="padding: 30px; color: #333; line-height: 1.5;">
          <p>Hi there,</p>
          <p>The guest Wi-Fi vouchers for your location have been updated. A fresh batch of 25 passes (20 Standard & 5 extended (3-Day for Ofsted, etc) has been generated and is ready to be printed:</p>
          
          <div style="background: #f9f9f9; padding: 15px; border-left: 4px solid #007bff; margin: 15px 0;">
            <strong>School:</strong> ${site.siteName}<br>
            <strong>Total Active Passes:</strong> 25 Vouchers<br>
            <strong>Voucher Types:</strong> 20 x 8-Hour Passes | 5 x 3-Day Passes<br>
            <strong>Device Limit:</strong> Single Use (Per Token)
          </div>
          
          <p style="margin-top: 25px; margin-bottom: 25px; text-align: center;">
            <a href="${webAppUrl}" target="_blank" style="background: #28a745; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">🖨️ Open Printable Tickets Page</a>
          </p>

          <div style="background: #e0f2fe; padding: 15px; border-radius: 6px; border: 1px solid #bae6fd; font-size: 13px; color: #0369a1; margin-bottom: 25px; line-height: 1.6;">
            💡 <strong>Handy Tip:</strong> While you can print these off all at once, you can also keep this link handy! The page connects directly to our Wi-Fi controller, meaning used vouchers will automatically disappear from the list in real-time so you can check back anytime to see exactly what is still available.
          </div>
          
          <p style="font-size: 13px; color: #555;">If the button above does not open, you can copy and paste this link into your browser:<br>
          <span style="font-family: monospace; font-size: 11px; color: #0369a1; word-break: break-all;">${webAppUrl}</span></p>
          
          <div style="background: #fff9db; padding: 15px; border-radius: 6px; border: 1px solid #ffe066; font-size: 12px; color: #856404; margin-top: 25px;">
            <strong>Please note:</strong> This print link is only valid for this week. A new link will be sent automatically next week.
          </div>
        </div>
      </div>
    `;
    
    MailApp.sendEmail({ to: site.recipients, subject: subject, body: body, htmlBody: htmlBody, noReply: true });
    
    console.log(`Dedicated manual sync and notification sequence successful for ${site.siteName}`);
    logToUiToast(`Sync complete and email dispatched for ${site.siteName}!`, "Success ✉️");
    
  } catch (err) {
    console.error(`Dedicated sync and email process failed for ${site.siteName}: ${err.toString()}`);
    logToUiToast(`Process failed: ${err.message}`, "Error ❌");
  }
}

/**
 * Background execution helper to run the database wipeout directly.
 */
function purgeAllUnusedVouchersDirect(siteConfig) {
  try {
    const realSiteId = getValidSiteId(siteConfig.consoleId);
    const cloudUrl = `https://api.ui.com/v1/connector/consoles/${siteConfig.consoleId}/proxy/network/integration/v1/sites/${realSiteId}/hotspot/vouchers`;
    const unused = fetchAllVouchersFromCloud(cloudUrl).filter(v => v.authorizedGuestCount === 0);
    
    for (let i = 0; i < unused.length; i++) {
      UrlFetchApp.fetch(`${cloudUrl}/${unused[i].id}`, { method: 'delete', headers: { 'X-API-Key': API_KEY, 'Accept': 'application/json' }, muteHttpExceptions: true });
    }
    console.log(`Security Purge Complete: Wiped ${unused.length} vouchers from the controller at ${siteConfig.siteName}`);
  } catch (e) {
    console.error(`Failed to execute background security purge for ${siteConfig.siteName}: ${e.message}`);
  }
}

/**
 * Manual Dashboard Trigger: Syncs all configured sites at once.
 */
function masterManualSyncAll() {
  const sites = getParsedSitesConfig();
  if (sites.length === 0) return;
  
  logToUiToast("Initializing batch top-up loop...", "Sync Running 🔄");
  sites.forEach(site => syncSingleSiteVouchers(site));
  logToUiToast("All sheets updated to match controller states!", "Complete ✅");
}

/**
 * Core processor for a singular specified site profile. Handles sheet tab creation and API requests.
 */
function syncSingleSiteVouchers(siteConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(siteConfig.siteName);
  
  if (!sheet) {
    sheet = ss.insertSheet(siteConfig.siteName);
    sheet.getRange(1, 1, 1, 3).setValues([["Voucher Code", "Duration", "Created At"]]).setFontWeight("bold");
    SpreadsheetApp.flush();
  } else {
    sheet.getRange(1, 1, 1, 3).setValues([["Voucher Code", "Duration", "Created At"]]).setFontWeight("bold");
  }

  try {
    const realSiteId = getValidSiteId(siteConfig.consoleId);
    const cloudUrl = `https://api.ui.com/v1/connector/consoles/${siteConfig.consoleId}/proxy/network/integration/v1/sites/${realSiteId}/hotspot/vouchers`;
    
    let vouchers = fetchAllVouchersFromCloud(cloudUrl);
    
    // 1. Replenish Standard 8-Hour Passes (480 minutes, Target: 20)
    const target8h = 20;
    let valid8h = vouchers.filter(v => v.expired === false && v.authorizedGuestCount === 0 && v.timeLimitMinutes === 480);
    if (valid8h.length < target8h) {
      const needed8h = target8h - valid8h.length;
      const payload8h = { count: needed8h, name: "Google Sheet Sync (8h)", authorizedGuestLimit: 1, timeLimitMinutes: 480 };
      UrlFetchApp.fetch(cloudUrl, {
        method: 'post', contentType: 'application/json',
        headers: { 'X-API-Key': API_KEY, 'Accept': 'application/json' },
        payload: JSON.stringify(payload8h), muteHttpExceptions: true
      });
    }

    // 2. Replenish Extended 3-Day Passes (4320 minutes, Target: 5)
    const target3d = 5;
    let valid3d = vouchers.filter(v => v.expired === false && v.authorizedGuestCount === 0 && v.timeLimitMinutes === 4320);
    if (valid3d.length < target3d) {
      const needed3d = target3d - valid3d.length;
      const payload3d = { count: needed3d, name: "Google Sheet Sync (3d)", authorizedGuestLimit: 1, timeLimitMinutes: 4320 };
      UrlFetchApp.fetch(cloudUrl, {
        method: 'post', contentType: 'application/json',
        headers: { 'X-API-Key': API_KEY, 'Accept': 'application/json' },
        payload: JSON.stringify(payload3d), muteHttpExceptions: true
      });
    }

    Utilities.sleep(2500); 
    
    vouchers = fetchAllVouchersFromCloud(cloudUrl);
    let validVouchers = vouchers.filter(v => v.expired === false && v.authorizedGuestCount === 0 && (v.timeLimitMinutes === 480 || v.timeLimitMinutes === 4320));
    validVouchers.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    }
    
    const outputData = validVouchers.map(v => {
      const formattedDate = v.createdAt ? new Date(v.createdAt).toLocaleString() : 'N/A';
      let code = v.code;
      if (code && code.length === 10 && !code.includes('-')) {
        code = code.substring(0, 5) + '-' + code.substring(5);
      }
      const durationText = v.timeLimitMinutes === 4320 ? '3 Days' : '8 Hours';
      return [code, durationText, formattedDate];
    });
    
    if (outputData.length > 0) {
      sheet.getRange(2, 1, outputData.length, 3).setValues(outputData);
    }
  } catch (err) {
    console.error(`Error syncing hardware for ${siteConfig.siteName}: ${err.message}`);
  }
}

/**
 * Contextual Clear Routine: Clears unused codes ONLY for the school corresponding to the active tab.
 */
function clearActiveSheetTabVouchers() {
  const sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const sites = getParsedSitesConfig();
  const activeSite = sites.find(s => s.siteName === sheetName);
  
  if (!activeSite) {
    SpreadsheetApp.getUi().alert("Operation Cancelled", "The active spreadsheet tab name does not match any configured school profile in your JSON variables properties box.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('⚠️ Purge Target Matrix', `Permanently delete ALL legacy unused vouchers from the UniFi hardware for: ${activeSite.siteName}?`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  try {
    logToUiToast(`Scanning UniFi controller database for ${activeSite.siteName}...`, "Processing 🔄");
    purgeAllUnusedVouchersDirect(activeSite);
    logToUiToast("Wipe complete! Re-syncing tab profile...", "Success 🧹");
    syncSingleSiteVouchers(activeSite);
  } catch (e) {
    logToUiToast(`Purge Error: ${e.message}`, "Failed ❌");
  }
}

/**
 * WEB APP ROUTER: Enforces tokens and determines which tab profile to query via the URL parameter.
 */
function doGet(e) {
  const currentToken = PropertiesService.getScriptProperties().getProperty('WEEKLY_ACCESS_TOKEN');
  const inboundToken = (e && e.parameter) ? e.parameter.t : null;
  const targetSiteName = (e && e.parameter) ? e.parameter.site : null;

  if (!currentToken || inboundToken !== currentToken || !targetSiteName) {
    return HtmlService.createHtmlOutput(`
      <div style="font-family:sans-serif; text-align:center; padding-top:60px; background-color:#f8fafc; height:100vh; margin:0;">
        <div style="max-width:450px; margin:0 auto; background:white; padding:40px; border-radius:8px; border:1px solid #e2e8f0;">
          <h1 style="color:#ef4444; font-size:36px; margin:0 0 15px 0;">🛑 Access Denied</h1>
          <p style="font-size:15px; color:#475569; line-height:1.6;">This voucher printing link has officially expired or been rotated for security.</p>
        </div>
      </div>
    `).setTitle('Access Expired').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const sites = getParsedSitesConfig();
  const matchedSite = sites.find(s => s.siteName === targetSiteName);
  let liveVouchersData = [];

  if (matchedSite) {
    try {
      const realSiteId = getValidSiteId(matchedSite.consoleId);
      const cloudUrl = `https://api.ui.com/v1/connector/consoles/${matchedSite.consoleId}/proxy/network/integration/v1/sites/${realSiteId}/hotspot/vouchers`;
      let validVouchers = fetchAllVouchersFromCloud(cloudUrl).filter(v => v.expired === false && v.authorizedGuestCount === 0 && (v.timeLimitMinutes === 480 || v.timeLimitMinutes === 4320));
      validVouchers.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
      
      liveVouchersData = validVouchers.map(v => {
        let code = v.code;
        if (code && code.length === 10 && !code.includes('-')) code = code.substring(0, 5) + '-' + code.substring(5);
        const durationText = v.timeLimitMinutes === 4320 ? '3 Days' : '8 Hours';
        return [code, v.createdAt ? new Date(v.createdAt).toLocaleString() : 'N/A', durationText];
      });
    } catch (error) {
      liveVouchersData = getTabFallbackData(targetSiteName);
    }
  } else {
    liveVouchersData = getTabFallbackData(targetSiteName);
  }

  const template = HtmlService.createTemplateFromFile('vouchers');
  template.vouchers = liveVouchersData; 
  template.siteTitle = targetSiteName; 
  return template.evaluate().setTitle(`Vouchers - ${targetSiteName}`).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Builds a dynamic launcher listing links for ALL schools matching your JSON definitions.
 */
function openPrintablePage() {
  if (!DEPLOYMENT_ID) {
    SpreadsheetApp.getUi().alert("Setup Required", "Please save your UNIFI_DEPLOYMENT_ID to Script Properties first.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const currentToken = getOrGenerateWeeklyToken(false);
  const sites = getParsedSitesConfig();
  
  let htmlLinks = `<div style="font-family:sans-serif; padding:10px; color:#1e293b;">
                    <p style="margin-bottom:15px; font-size:13px;">Select a location dashboard portal to access live printable ticket forms:</p>`;
  
  sites.forEach(s => {
    const link = `https://script.google.com/macros/s/${DEPLOYMENT_ID}/exec?t=${currentToken}&site=${encodeURIComponent(s.siteName)}`;
    htmlLinks += `<div style="margin-bottom:12px;">
                    <a href="${link}" target="_blank" style="background:#007bff; color:white; padding:8px 14px; text-decoration:none; border-radius:4px; font-weight:bold; display:block; text-align:center; font-size:13px;">${s.siteName} Tickets</a>
                  </div>`;
  });
  
  htmlLinks += `</div>`;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlLinks).setWidth(380).setHeight(Math.min(250, 80 + (sites.length * 45)));
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🖨️ Multi-Site Printable Launchpad');
}

/**
 * Safe non-blocking toast notification router context wrapper shield.
 */
function logToUiToast(msg, title) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, 6);
  } catch (contextError) {}
}

// ==========================================
//            CORE UTILITY HELPERS
// ==========================================

function getParsedSitesConfig() {
  if (!SITES_JSON) return [];
  try {
    const parsed = JSON.parse(SITES_JSON);
    return parsed.map(site => {
      if (!site.hasOwnProperty('purgeBeforeSync')) {
        site.purgeBeforeSync = false;
      }
      return site;
    });
  } catch(e) { return []; }
}

function getTabFallbackData(tabName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
      return data.map(row => [row[0], row[2], row[1]]);
    }
  } catch(e) {}
  return [];
}

function getOrGenerateWeeklyToken(forceNew) {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('WEEKLY_ACCESS_TOKEN');
  if (!token || forceNew) {
    token = Utilities.getUuid(); 
    props.setProperty('WEEKLY_ACCESS_TOKEN', token);
  }
  return token;
}

function fetchAllVouchersFromCloud(baseUrl) {
  let combinedVouchers = []; let offset = 0; const pageLimit = 500; let keepsQuerying = true;
  while (keepsQuerying) {
    const response = UrlFetchApp.fetch(`${baseUrl}?limit=${pageLimit}&offset=${offset}`, { method: 'get', headers: { 'X-API-Key': API_KEY, 'Accept': 'application/json' }, muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) throw new Error("Database read failed.");
    const result = JSON.parse(response.getContentText());
    const dataBatch = result.data || result.vouchers || [];
    combinedVouchers = combinedVouchers.concat(dataBatch);
    offset += pageLimit;
    if (combinedVouchers.length >= (result.totalCount || 0) || dataBatch.length === 0) keepsQuerying = false;
  }
  return combinedVouchers;
}

function getValidSiteId(consoleIdString) {
  const response = UrlFetchApp.fetch(`https://api.ui.com/v1/connector/consoles/${consoleIdString}/proxy/network/integration/v1/sites`, { method: 'get', headers: { 'X-API-Key': API_KEY, 'Accept': 'application/json' }, muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error("Site verification failed.");
  const sites = JSON.parse(response.getContentText()).data || [];
  const match = sites.find(s => s.internalReference === 'default' || s.name.toLowerCase() === 'default');
  return match ? match.id : (sites.length > 0 ? sites[0].id : null);
}

/**
 * Ensures required script properties exist in Script Properties, creating them as blank strings if missing.
 */
function ensureScriptPropertiesExist() {
  const props = PropertiesService.getScriptProperties();
  const requiredKeys = ['UNIFI_API_KEY', 'UNIFI_DEPLOYMENT_ID', 'UNIFI_SITES_CONFIG', 'WEEKLY_ACCESS_TOKEN'];
  const existingProps = props.getProperties();
  
  requiredKeys.forEach(key => {
    if (!Object.prototype.hasOwnProperty.call(existingProps, key)) {
      props.setProperty(key, '');
    }
  });
}

/**
 * Installs a time-based trigger to run weeklyAutoReplenishAndEmail every Monday at 7 AM.
 */
// ⚠️ Installs an automated background trigger running under your account - verify all site recipient email addresses before installing.
function installWeeklyTrigger() {
  const handler = 'weeklyAutoReplenishAndEmail';
  
  // Clean up any existing triggers for this handler to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handler)
    .forEach(t => ScriptApp.deleteTrigger(t));
    
  // Create a weekly trigger for Mondays at 7:00 AM
  ScriptApp.newTrigger(handler)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
    
  logToUiToast("Weekly automated trigger successfully installed for Mondays at 7:00 AM!", "Trigger Installed ⏰");
}

/**
 * Removes any installed time-based triggers for weeklyAutoReplenishAndEmail.
 */
function uninstallWeeklyTrigger() {
  const handler = 'weeklyAutoReplenishAndEmail';
  
  const existingTriggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handler);
    
  existingTriggers.forEach(t => ScriptApp.deleteTrigger(t));
    
  if (existingTriggers.length > 0) {
    logToUiToast("Weekly automated trigger successfully removed.", "Trigger Removed 🗑️");
  } else {
    logToUiToast("No active weekly trigger was found to remove.", "Notice ℹ️");
  }
}

/**
 * MASTER BOOT Hook: Builds your complete operational interface on spreadsheet load.
 */
function onOpen() {
  ensureScriptPropertiesExist();
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('UniFi Wi-Fi')
    .addItem('Sync All Configured Sites', 'masterManualSyncAll')
    .addSeparator()
    .addItem('Sync ONLY Active Tab Site', 'syncActiveSheetTabVouchers')
    .addItem('Force Sync & Email Active Tab Site', 'syncAndEmailActiveTabSite')
    .addSeparator()
    .addItem('🖨️ View Printable Tickets Launchpad', 'openPrintablePage')
    .addSeparator()
    .addItem('✉️ Force Run Weekly Email Test', 'weeklyAutoReplenishAndEmail')
    .addItem('⏰ Install Weekly Trigger (Mondays @ 7 AM)', 'installWeeklyTrigger')
    .addItem('🗑️ Remove Weekly Trigger', 'uninstallWeeklyTrigger')
    .addSeparator()
    .addItem('Clear Active Tab Unused Vouchers from Controller', 'clearActiveSheetTabVouchers')
    .addToUi();
}
