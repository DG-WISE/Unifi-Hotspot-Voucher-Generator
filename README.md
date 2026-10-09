# UniFi Guest Wi-Fi Voucher Generator

An automated, multi-site Google Apps Script and Web App ecosystem for managing, replenishing, and distributing UniFi Guest Wi-Fi vouchers.
This system automatically maintains a steady pool of 20 Standard (8-Hour) and 5 Extended (3-Day) passes for multiple school sites, dispatches weekly stylized emails with live dashboard links, and provides a custom Google Sheets UI for on-demand syncing.

---

## Features
- **Multi-Site Support:** Manage unlimited consoles/sites from a single JSON configuration.
- **Automated Replenishment:** Weekly background triggers top-up missing vouchers rather than wiping existing active ones.
- **Dual Voucher Types:** Automatically generates both 8-hour standard passes and 3-day extended passes (ideal for inspections/Ofsted or multi-day visitors).
- **Dynamic Printable Dashboard:** A secure, live-updating web app portal for receptionists and administrators to view and print tickets.
- **Granular Security:** Per-site `purgeBeforeSync` flags allow for immediate, targeted wipeouts of unused codes across individual consoles.
- **Trigger Management:** Built-in UI menu options to install or remove weekly background automation triggers without touching code.

---

## 🔑 Creating a UniFi API Key

To allow Google Apps Script to communicate with your UniFi consoles via Ubiquiti's Official Integration API, you must generate an API key from your UniFi Cloud account:

1. Log in to the UniFi Site Manager portal at [unifi.ui.com](https://unifi.ui.com).
2. Click on your account profile icon and navigate to **Account Settings** (or go directly to [account.ui.com](https://account.ui.com)).
3. Select **API Keys** or **Integrations** from the menu.
4. Click **Create API Key**. Assign a clear name (e.g., `Google Sheets Voucher Generator`).
5. Ensure the key has full permissions over the **Network Application** component across all required consoles.
6. Copy the generated token immediately and store it securely. You will paste this key into the `UNIFI_API_KEY` Script Property inside Google Apps Script.

> ⚠️ **Permission Requirement:** Personal API Keys inherit the permissions of the account that generated them. Ensure the creating account is designated as the **Owner** or **Super Admin** on every target console in your setup. If managing consoles under an enterprise structure, generate an **Organization API Key** instead.

---
## ⚙️ Setup Instructions

You can deploy this project using either the standard manual method (**Option A**) or via the command line using Google's `clasp` CLI tool (**Option B**).

### Option A: Manual Setup (Google Sheets UI)

1. **Create the Google Sheet:**
   - Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet. Name it something like `UniFi Guest Wi-Fi Manager`.
2. **Open the Apps Script Editor:**
   - In the top menu, click **Extensions** > **Apps Script**.
3. **Add the Code:**
   - Delete any default code in `Code.gs` and paste the contents of the repository's `Code.gs` file.
   - Click the **+** icon next to **Files** in the sidebar, select **HTML**, and name it exactly `vouchers` (Google will save it as `vouchers.html`).
   - Paste the contents of the repository's `vouchers.html` file into this new tab.
   - Click the **Save** (floppy disk) icon.
4. **Initialize Script Properties:**
   - Select `ensureScriptPropertiesExist` from the function dropdown in the top editor toolbar and click **Run**. Grant any requested authorization permissions.
5. **Deploy as a Web App:**
   - Click **Deploy** (top right) > **New deployment**.
   - Select **Web app** as the deployment type.
   - Set **Execute as**: `Me`
   - Set **Who has access**: `Domain` (Security is handled dynamically via generated session access tokens).
   - Click **Deploy** and copy the resulting **Web App URL**. Extract your `DEPLOYMENT_ID` (the string located between `/s/` and `/exec` in the URL).
6. **Configure Environment Variables:**
   - Click **Project Settings** (⚙️ gear icon in the left sidebar).
   - Scroll down to **Script Properties** and edit the required keys:
     - `UNIFI_API_KEY`: Your official UniFi API key.
     - `UNIFI_DEPLOYMENT_ID`: The deployment UUID extracted in step 5.
     - `UNIFI_SITES_CONFIG`: Your JSON array mapping site configurations (see schema below).
7. **Build the Custom UI Menu:**
   - Return to the editor, select `onOpen` from the toolbar dropdown, and click **Run**.
   - Switch back to your Google Sheet tab to see your new **UniFi Wi-Fi** custom menu on the toolbar.

---

### Option B: Command Line Setup (Using `clasp`)

If you manage your scripts locally using Node.js and Google's [clasp CLI](https://developers.google.com/apps-script/guides/clasp):

1. **Clone the Repository:**
   ```bash
   git clone [https://github.com/your-org/unifi-voucher-generator.git](https://github.com/your-org/unifi-voucher-generator.git)
   cd unifi-voucher-generator
   ```
2. **Authenticate with Google:**
   ```bash
   clasp login
   ```
3. **Create a Bound Apps Script Project:**
   ```bash
   clasp create --type sheets --title "UniFi Guest Wi-Fi Vouchers"
   ```
4. **Push Project Files:**
   ```bash
   clasp push
   ```
5. **Configure Properties & Deploy:**
   - Open the web editor via `clasp open`.
   - Configure your **Script Properties** in Project Settings (⚙️).
   - Click **Deploy** > **New deployment** > **Web app** (`Execute as: Me`, `Access: Anyone`).
   - Run the `onOpen` function once in the editor to initialize the spreadsheet interface.

---

## 🛠️ Configuration Schema (`UNIFI_SITES_CONFIG`)

Multi-site targets are defined inside a JSON array stored in the `UNIFI_SITES_CONFIG` Script Property.

```json
[
  {
    "siteName": "Primary Academy",
    "consoleId": "YOUR_CONSOLE_UUID_1:YOUR_SITE_ID_1",
    "recipients": "it@yourdomain.com, office@yourdomain.com",
    "purgeBeforeSync": false
  },
  {
    "siteName": "Secondary Academy",
    "consoleId": "YOUR_CONSOLE_UUID_2:YOUR_SITE_ID_2",
    "recipients": "admin@yourdomain.com, reception@yourdomain.com",
    "purgeBeforeSync": false
  }
]
```

### Property Descriptions
- `siteName`: Must match the exact name of the tab in your Google Sheet.
- `consoleId`: The UniFi hardware Console UUID and Site ID (`<console_uuid>:<site_id>`) found in your UniFi Site Manager URL.
- `recipients`: Comma-separated list of email addresses that receive weekly voucher notifications.
- `purgeBeforeSync`: Security flag. When set to `true`, the script will permanently delete all active unused vouchers on that console during the next sync pass before generating a fresh batch.

---

## 🚀 Usage Guidance

All manual controls are accessible directly from the **UniFi Wi-Fi** menu inside Google Sheets:

### Menu Actions

| Menu Option | Function | Description |
| :--- | :--- | :--- |
| **Sync All Configured Sites** | `masterManualSyncAll` | Runs a silent top-up across all sites defined in your JSON array. No emails sent. |
| **Sync ONLY Active Tab Site** | `syncActiveSheetTabVouchers` | Performs a silent background top-up strictly for the site matching your current tab. |
| **Force Sync & Email Active Tab Site** | `syncAndEmailActiveTabSite` | Tops up the current tab's site and dispatches an updated notification email to its recipients. |
| **🖨️ View Printable Launchpad** | `openPrintablePage` | Opens a modal dialog with direct links to live printable voucher web pages for all sites. |
| **✉️ Force Run Weekly Email Test** | `weeklyAutoReplenishAndEmail` | Manually triggers the master weekly automated workflow (syncs all sites and sends all emails). |
| **⏰ Install Weekly Trigger** | `installWeeklyTrigger` | Installs an automated background trigger set to run every Monday at 7:00 AM. |
| **🗑️ Remove Weekly Trigger** | `uninstallWeeklyTrigger` | Uninstalls any active automated background replenishment triggers. |
| **Clear Active Tab Unused Vouchers** | `clearActiveSheetTabVouchers` | Prompts for confirmation and permanently purges all unused codes from the active tab's console. |
