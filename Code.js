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
  if (sites.length === 0) {
    console.warn("[Weekly Auto Replenish] No site configurations found in UNIFI_SITES_CONFIG.");
    return;
  }

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
  console.log(`[Sync Active Tab] Triggered for tab: "${sheetName}"`);
  const sites = getParsedSitesConfig();
  const activeSite = sites.find(s => s.siteName === sheetName);
  
  if (!activeSite) {
    console.warn(`[Sync Active Tab] No site configuration matching active tab name: "${sheetName}"`);
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
  console.log(`[Sync & Email Active Tab] Triggered for tab: "${sheetName}"`);
  const sites = getParsedSitesConfig();
  const site = sites.find(s => s.siteName === sheetName);
  
  if (!site) {
    console.warn(`[Sync & Email Active Tab] No site configuration matching active tab name: "${sheetName}"`);
    SpreadsheetApp.getUi().alert(
      "Operation Cancelled", 
      "The active spreadsheet tab name does not match any configured school profile in your JSON properties box.", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  if (!DEPLOYMENT_ID) {
    console.error(`[Sync & Email Active Tab] Missing UNIFI_DEPLOYMENT_ID script property.`);
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
    console.log(`[Purge Direct] Initiating purge for site: ${siteConfig.siteName}`);
    const realSiteId = getValidSiteId(siteConfig.consoleId);
    const cloudUrl = 'https://api.ui.com/v1/connector/consoles/' + siteConfig.consoleId + '/proxy/network/integration/v1/sites/' + realSiteId + '/hotspot/vouchers';
    console.log(`[Purge Direct] Fetching all vouchers to identify unused entries...`);
    const unused = fetchAllVouchersFromCloud(cloudUrl).filter(v => v.authorizedGuestCount === 0);
    console.log(`[Purge Direct] Found ${unused.length} unused vouchers to purge.`);
    
    for (let i = 0; i < unused.length; i++) {
      console.log(`[Purge Direct API] Deleting voucher ID: ${unused[i].id} (${i + 1}/${unused.length})...`);
      const resp = UrlFetchApp.fetch(cloudUrl + '/' + unused[i].id, { method: 'delete', headers: { 'X-API-Key': API_KEY, 'Accept': 'application/json' }, muteHttpExceptions: true });
      const code = resp.getResponseCode();
      if (code >= 400) {
        console.warn(`[Purge Direct API Warning] Failed to delete voucher ${unused[i].id} (HTTP ${code}): ${resp.getContentText()}`);
      }
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
  console.log("[Master Sync All] Starting manual sync for all configured sites...");
  const sites = getParsedSitesConfig();
  if (sites.length === 0) {
    console.warn("[Master Sync All] No site configurations found.");
    return;
  }
  
  logToUiToast("Initializing batch top-up loop...", "Sync Running 🔄");
  sites.forEach(site => syncSingleSiteVouchers(site));
  logToUiToast("All sheets updated to match controller states!", "Complete ✅");
  console.log("[Master Sync All] Finished sync across all site configurations.");
}

/**
 * Core processor for a singular specified site profile. Handles sheet tab creation and API requests.
 */
function syncSingleSiteVouchers(siteConfig) {
  console.log(`[Sync Single Site] Beginning voucher sync process for: ${siteConfig.siteName}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(siteConfig.siteName);
  
  if (!sheet) {
    console.log(`[Sheet Ops] Tab "${siteConfig.siteName}" not found. Inserting new sheet.`);
    sheet = ss.insertSheet(siteConfig.siteName);
    console.log(`[Sheet Ops] Writing header row to new sheet "${siteConfig.siteName}".`);
    sheet.getRange(1, 1, 1, 3).setValues([["Voucher Code", "Duration", "Created At"]]).setFontWeight("bold");
    SpreadsheetApp.flush();
  } else {
    console.log(`[Sheet Ops] Found existing tab "${siteConfig.siteName}". Formatting header row.`);
    sheet.getRange(1, 1, 1, 3).setValues([["Voucher Code", "Duration", "Created At"]]).setFontWeight("bold");
  }

  try {
    console.log(`[Sync Single Site] Resolving site ID for console: ${siteConfig.consoleId}`);
    const realSiteId = getValidSiteId(siteConfig.consoleId);
    const cloudUrl = 'https://api.ui.com/v1/connector/consoles/' + siteConfig.consoleId + '/proxy/network/integration/v1/sites/' + realSiteId + '/hotspot/vouchers';
    
    console.log(`[Sync Single Site] Fetching initial voucher inventory from controller...`);
    let vouchers = fetchAllVouchersFromCloud(cloudUrl);
    
    // 1. Replenish Standard 8-Hour Passes (480 minutes, Target: 20)
    const target8h = 20;
    let valid8h = vouchers.filter(v => v.expired === false && v.authorizedGuestCount === 0 && v.timeLimitMinutes === 480);
    console.log(`[Sync Single Site] 8-Hour Passes: Found ${valid8h.length} valid unused (Target: ${target8h})`);
    if (valid8h.length < target8h) {
      const needed8h = target8h - valid8h.length;
      console.log(`[API POST] Replenishing 8-Hour Passes: Generating ${needed8h} new vouchers...`);
      const payload8h = { count: needed8h, name: "Google Sheet Sync (8h)", authorizedGuestLimit: 1, timeLimitMinutes: 480 };
      const resp8h = UrlFetchApp.fetch(cloudUrl, {
        method: 'post', contentType: 'application/json',
        headers: { 'X-API-Key': API_KEY, 'Accept': 'application/json' },
        payload: JSON.stringify(payload8h), muteHttpExceptions: true
      });
      console.log(`[API POST] 8-Hour replenishment HTTP status: ${resp8h.getResponseCode()}`);
      if (resp8h.getResponseCode() >= 400) {
        console.error(`[API POST Error] 8-Hour generation failed: ${resp8h.getContentText()}`);
      }
    }

    // 2. Replenish Extended 3-Day Passes (4320 minutes, Target: 5)
    const target3d = 5;
    let valid3d = vouchers.filter(v => v.expired === false && v.authorizedGuestCount === 0 && v.timeLimitMinutes === 4320);
    console.log(`[Sync Single Site] 3-Day Passes: Found ${valid3d.length} valid unused (Target: ${target3d})`);
    if (valid3d.length < target3d) {
      const needed3d = target3d - valid3d.length;
      console.log(`[API POST] Replenishing 3-Day Passes: Generating ${needed3d} new vouchers...`);
      const payload3d = { count: needed3d, name: "Google Sheet Sync (3d)", authorizedGuestLimit: 1, timeLimitMinutes: 4320 };
      const resp3d = UrlFetchApp.fetch(cloudUrl, {
        method: 'post', contentType: 'application/json',
        headers: { 'X-API-Key': API_KEY, 'Accept': 'application/json' },
        payload: JSON.stringify(payload3d), muteHttpExceptions: true
      });
      console.log(`[API POST] 3-Day replenishment HTTP status: ${resp3d.getResponseCode()}`);
      if (resp3d.getResponseCode() >= 400) {
        console.error(`[API POST Error] 3-Day generation failed: ${resp3d.getContentText()}`);
      }
    }

    console.log(`[Sync Single Site] Pausing 2.5 seconds for controller database propagation...`);
    Utilities.sleep(2500); 
    
    console.log(`[Sync Single Site] Re-fetching voucher inventory from controller after generation...`);
    vouchers = fetchAllVouchersFromCloud(cloudUrl);
    let validVouchers = vouchers.filter(v => v.expired === false && v.authorizedGuestCount === 0 && (v.timeLimitMinutes === 480 || v.timeLimitMinutes === 4320));
    validVouchers.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      console.log(`[Sheet Ops] Clearing existing data rows 2 to ${lastRow} in tab "${siteConfig.siteName}".`);
      sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
    } else {
      console.log(`[Sheet Ops] No existing data rows to clear in tab "${siteConfig.siteName}".`);
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
      console.log(`[Sheet Ops] Writing ${outputData.length} active voucher rows to sheet range A2:C${outputData.length + 1} on tab "${siteConfig.siteName}".`);
      sheet.getRange(2, 1, outputData.length, 3).setValues(outputData);
      console.log(`[Sheet Ops Writeback Success] Sheet update completed for "${siteConfig.siteName}".`);
    } else {
      console.warn(`[Sheet Ops Warning] No valid active vouchers found to write for "${siteConfig.siteName}".`);
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
  console.log(`[Clear Active Tab] Triggered for tab: "${sheetName}"`);
  const sites = getParsedSitesConfig();
  const activeSite = sites.find(s => s.siteName === sheetName);
  
  if (!activeSite) {
    console.warn(`[Clear Active Tab] No site configuration matching tab name: "${sheetName}"`);
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
    console.error(`[Clear Active Tab Error] Failed to clear vouchers for "${activeSite.siteName}": ${e.message}`);
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

  console.log(`[Web App doGet] Request received. Target site: "${targetSiteName}", Token provided: ${!!inboundToken}`);

  if (!currentToken || inboundToken !== currentToken || !targetSiteName) {
    console.warn(`[Web App doGet] Access Denied: Token validation failed or missing parameters.`);
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
      console.log(`[Web App doGet] Fetching live vouchers from controller for site: ${targetSiteName}`);
      const realSiteId = getValidSiteId(matchedSite.consoleId);
      const cloudUrl = 'https://api.ui.com/v1/connector/consoles/' + matchedSite.consoleId + '/proxy/network/integration/v1/sites/' + realSiteId + '/hotspot/vouchers';
      let validVouchers = fetchAllVouchersFromCloud(cloudUrl).filter(v => v.expired === false && v.authorizedGuestCount === 0 && (v.timeLimitMinutes === 480 || v.timeLimitMinutes === 4320));
      validVouchers.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
      
      liveVouchersData = validVouchers.map(v => {
        let code = v.code;
        if (code && code.length === 10 && !code.includes('-')) code = code.substring(0, 5) + '-' + code.substring(5);
        const durationText = v.timeLimitMinutes === 4320 ? '3 Days' : '8 Hours';
        return [code, v.createdAt ? new Date(v.createdAt).toLocaleString() : 'N/A', durationText];
      });
      console.log(`[Web App doGet] Prepared ${liveVouchersData.length} live vouchers for HTML output.`);
    } catch (error) {
      console.warn(`[Web App doGet Warning] Live API fetch failed for "${targetSiteName}": ${error.message}. Falling back to sheet tab data.`);
      liveVouchersData = getTabFallbackData(targetSiteName);
    }
  } else {
    console.warn(`[Web App doGet Warning] Site profile "${targetSiteName}" not found in JSON config. Falling back to sheet tab data.`);
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
    console.error("[openPrintablePage] UNIFI_DEPLOYMENT_ID is not configured in Script Properties.");
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
  if (!SITES_JSON) {
    console.warn("[getParsedSitesConfig] UNIFI_SITES_CONFIG Script Property is empty.");
    return [];
  }
  try {
    const parsed = JSON.parse(SITES_JSON);
    return parsed.map(site => {
      if (!Object.prototype.hasOwnProperty.call(site, 'purgeBeforeSync')) {
        site.purgeBeforeSync = false;
      }
      return site;
    });
  } catch(e) {
    console.error(`[getParsedSitesConfig Error] Failed to parse UNIFI_SITES_CONFIG JSON: ${e.message}`);
    return [];
  }
}

function getTabFallbackData(tabName) {
  console.log(`[Fallback Data] Reading fallback data from sheet tab: "${tabName}"`);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
    if (sheet && sheet.getLastRow() > 1) {
      const lastRow = sheet.getLastRow();
      console.log(`[Fallback Data] Found ${lastRow - 1} data rows on tab "${tabName}".`);
      const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      return data.map(row => [row[0], row[2], row[1]]);
    } else {
      console.warn(`[Fallback Data] Tab "${tabName}" does not exist or has no rows past header.`);
    }
  } catch(e) {
    console.error(`[Fallback Data Error] Reading tab "${tabName}" failed: ${e.message}`);
  }
  return [];
}

function getOrGenerateWeeklyToken(forceNew) {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('WEEKLY_ACCESS_TOKEN');
  if (!token || forceNew) {
    token = Utilities.getUuid(); 
    props.setProperty('WEEKLY_ACCESS_TOKEN', token);
    console.log(`[Weekly Token] Generated new weekly security access token.`);
  } else {
    console.log(`[Weekly Token] Using existing weekly security access token.`);
  }
  return token;
}

function fetchAllVouchersFromCloud(baseUrl) {
  let combinedVouchers = []; let offset = 0; const pageLimit = 500; let keepsQuerying = true;
  console.log(`[API Fetch] Starting paginated voucher retrieval from: ${baseUrl}`);
  while (keepsQuerying) {
    const fetchUrl = baseUrl + '?limit=' + pageLimit + '&offset=' + offset;
    console.log(`[API Fetch] Querying page at offset ${offset} (limit ${pageLimit})...`);
    const response = UrlFetchApp.fetch(fetchUrl, { method: 'get', headers: { 'X-API-Key': API_KEY, 'Accept': 'application/json' }, muteHttpExceptions: true });
    const code = response.getResponseCode();
    console.log(`[API Fetch] HTTP Status: ${code}`);
    if (code !== 200) {
      console.error(`[API Fetch Error] Database read failed (HTTP ${code}): ${response.getContentText()}`);
      throw new Error('Database read failed (HTTP ' + code + ')');
    }
    const result = JSON.parse(response.getContentText());
    const dataBatch = result.data || result.vouchers || [];
    console.log(`[API Fetch] Retrieved ${dataBatch.length} vouchers in batch (Total available on controller: ${result.totalCount || 'N/A'})`);
    combinedVouchers = combinedVouchers.concat(dataBatch);
    offset += pageLimit;
    if (combinedVouchers.length >= (result.totalCount || 0) || dataBatch.length === 0) keepsQuerying = false;
  }
  console.log(`[API Fetch Complete] Total vouchers gathered across all pages: ${combinedVouchers.length}`);
  return combinedVouchers;
}

function getValidSiteId(consoleIdString) {
  const apiUrl = 'https://api.ui.com/v1/connector/consoles/' + consoleIdString + '/proxy/network/integration/v1/sites';
  console.log(`[API Site Verification] Requesting site definitions for console: ${consoleIdString}`);
  const response = UrlFetchApp.fetch(apiUrl, { method: 'get', headers: { 'X-API-Key': API_KEY, 'Accept': 'application/json' }, muteHttpExceptions: true });
  const code = response.getResponseCode();
  console.log(`[API Site Verification] HTTP Status: ${code}`);
  if (code !== 200) {
    console.error(`[API Site Verification Error] HTTP ${code}: ${response.getContentText()}`);
    throw new Error('Site verification failed (HTTP ' + code + ')');
  }
  const sites = JSON.parse(response.getContentText()).data || [];
  console.log(`[API Site Verification] Found ${sites.length} site profiles on console.`);
  const match = sites.find(s => s.internalReference === 'default' || s.name.toLowerCase() === 'default');
  const resolvedId = match ? match.id : (sites.length > 0 ? sites[0].id : null);
  console.log(`[API Site Verification] Selected site ID: ${resolvedId}`);
  return resolvedId;
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
