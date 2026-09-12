<<<<<<< HEAD
# location-registration-form
=======
# Location Registration Form with Cascading Location Hierarchy

A production-ready Google Apps Script Web App for collecting location-based registration data with real-time cascading dropdowns, worldwide country support, Indian State/UT → District → Pincode → Tehsil/Block auto-resolution, duplicate phone & location detection, and automatic Google Sheets database management.

---

## 🌟 Key Features

* **Cascading Dropdowns**: Dynamic JavaScript loading:
  * Country → State/UT
  * State/UT → District
  * District → Pincode
  * Pincode → Sub-district / Tehsil / Block (auto-filled)
* **Worldwide & India Coverage**: Includes 190+ worldwide countries with generic state/postal code fallbacks, plus dedicated deep location cascading for India.
* **Duplicate Detection**:
  * **Duplicate Phone**: Server-side block on re-registering existing phone numbers.
  * **Duplicate Location**: Flags identical exact addresses/locations, highlights the row in yellow (`#FFF3CD`) in Google Sheets, and alerts admins.
* **Admin Email Notifications**: Automated emails sent to the configured admin upon every submission (with flags for duplicates).
* **High Performance**: Location lists are cached using `CacheService` (6-hour cache) for fast response times.
* **Auto-managed Database**: Script automatically creates and formats required Google Sheet tabs (`Responses`, `Countries`, `IN_States`, `IN_Districts`, `IN_Pincodes`, `Config`, `RawImport`).
* **Bulk Import Engine**: Built-in `bulkImportLocationData()` function to easily import full All-India official pincode datasets from data.gov.in.
* **Responsive Frontend**: Mobile-optimized design with modern visual aesthetics, loading spinners, and status feedback.

---

## 📁 Repository Structure

```
├── Code.gs      # Google Apps Script backend logic, database management, RPC endpoints & validations
├── Index.html   # Web app frontend HTML, CSS styling, and client-side JavaScript logic
└── README.md    # Setup, deployment, data import, and GitHub hosting guide
```

---

## 🚀 Step-by-Step Deployment Guide

### 1. Create your Google Sheet Database
1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet.
2. Name the spreadsheet **Location Registrations**.

### 2. Open Apps Script Editor
1. In the Google Sheet menu, click **Extensions → Apps Script**.
2. Rename the project to **Location Registration Web App**.

### 3. Add Project Code
1. Open `Code.gs` in the editor, delete any default code, and paste the contents of [`Code.gs`](Code.gs).
2. Click the **+** button next to **Files** in the sidebar, select **HTML**, name it **`Index`** (exact match), and paste the contents of [`Index.html`](Index.html).

### 4. Initialize Database Schema
1. In the function dropdown at the top of the editor, select **`setupSheets`** and click **Run**.
2. Authorize the OAuth prompts when requested (allows the script to manage the sheet and send emails).
3. Select **`seedIndiaSampleData`** from the function dropdown and click **Run**. This loads demo states, districts, and pincodes for immediate testing.

### 5. Deploy as Web App
1. Click **Deploy → New deployment**.
2. Click the gear icon next to **Select type** and choose **Web app**.
3. Configure settings:
   * **Description**: Location Registration Form v1.0
   * **Execute as**: `Me` (your Google account)
   * **Who has access**: `Anyone` (or *Anyone within [your organization]* for internal use)
4. Click **Deploy**, copy the generated **Web App URL**, and open it in your browser.

---

## 📊 How to Import the Full All-India Location Dataset

The demo dataset covers sample districts and pincodes. To populate the complete official dataset across India:

1. **Download Official Data**: Obtain the *All India Pincode Directory* published by the Department of Posts via [data.gov.in](https://data.gov.in).
2. **Format Source File**: Clean the dataset to **exactly four columns** in this specific order:
   * Column A: `State/UT`
   * Column B: `District`
   * Column C: `Pincode`
   * Column D: `Tehsil/Block` (or Sub-district)
3. **Paste into Sheet**: Open your Google Sheet, switch to the automatically generated **`RawImport`** sheet, and paste your formatted data starting at **Row 2** (keep Row 1 headers intact).
4. **Run Bulk Import**:
   * Open the Apps Script editor.
   * Select **`bulkImportLocationData`** from the function dropdown and click **Run**.
   * The script will deduplicate entries, build `IN_States`, `IN_Districts`, and `IN_Pincodes` tabs, and automatically flush the cache.
5. **Verify Form**: Refresh your Web App URL. The dropdown cascade will now reflect the full dataset.

---

## 🛠 Database Schema Overview

| Sheet Name | Purpose |
| :--- | :--- |
| **`Responses`** | Stores all form submissions with timestamp, contact info, location details, duplicate flags, and review status. |
| **`Countries`** | List of worldwide countries. |
| **`IN_States`** | Unique list of Indian States and Union Territories. |
| **`IN_Districts`** | Mapping of State/UT to District. |
| **`IN_Pincodes`** | Mapping of District to Pincode and Tehsil/Block. |
| **`Config`** | Configuration settings (e.g., `AdminEmail`, `FormTitle`). |
| **`RawImport`** | Staging tab for bulk importing location directory data. |

---

## 🌐 Hosting on GitHub (`https://github.com/tesseractthou-code`)

To host and manage this project repository under your GitHub account:

### Push Repository to GitHub:
```bash
# Initialize local git repository
git init

# Add all files
git add Code.gs Index.html README.md

# Commit changes
git commit -m "Initial commit: Location Registration Web App with cascading dropdowns"

# Set main branch
git branch -M main

# Add remote repository link (create repository on github.com first)
git remote add origin https://github.com/tesseractthou-code/location-registration-form.git

# Push to GitHub
git push -u origin main
```

---

## 📄 License
This project is open-source and available under the [MIT License](LICENSE).
>>>>>>> 3cb10b4 (Initial commit: Location Registration Web App with cascading dropdowns)
