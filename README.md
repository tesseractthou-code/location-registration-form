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
* **Pre-formatted All-India Dataset**: Includes `RawImport_ready.csv` containing 34,979 unique official Indian pincodes ready to load into Google Sheets.
* **Duplicate Detection**:
  * **Duplicate Phone**: Server-side block on re-registering existing phone numbers.
  * **Duplicate Location**: Flags identical exact addresses/locations, highlights the row in yellow (`#FFF3CD`) in Google Sheets, and alerts admins.
* **Admin Email Notifications**: Automated emails sent to the configured admin upon every submission (with flags for duplicates).
* **High Performance**: Location lists are cached using `CacheService` (6-hour cache) for fast response times.
* **Auto-managed Database**: Script automatically creates and formats required Google Sheet tabs (`Responses`, `Countries`, `IN_States`, `IN_Districts`, `IN_Pincodes`, `Config`, `RawImport`).
* **Bulk Import Engine**: Built-in `bulkImportLocationData()` function to easily import full All-India official pincode datasets.

---

## 📁 Repository Structure

```
├── Code.gs                 # Google Apps Script backend logic, database management, RPC endpoints & validations
├── Index.html              # Web app frontend HTML, CSS styling, and client-side JavaScript logic
├── RawImport_ready.csv     # Pre-formatted All-India Pincode dataset (34,979 records ready for RawImport sheet)
├── format_pincode_csv.py   # Python utility script to format custom raw pincode CSV files
└── README.md               # Setup, deployment, data import, and GitHub hosting guide
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
2. Authorize the OAuth prompts when requested.

### 5. Populate Full All-India Location Dataset
1. Open your Google Sheet and select the **`RawImport`** tab.
2. Import or paste the records from [`RawImport_ready.csv`](RawImport_ready.csv) starting at **Row 2** (keep Row 1 headers intact).
3. In the Apps Script editor, select **`bulkImportLocationData`** from the function dropdown and click **Run**.

### 6. Deploy as Web App
1. Click **Deploy → New deployment**.
2. Click the gear icon next to **Select type** and choose **Web app**.
3. Configure settings:
   * **Description**: Location Registration Form v1.0
   * **Execute as**: `Me` (your Google account)
   * **Who has access**: `Anyone` (or *Anyone within [your organization]*)
4. Click **Deploy**, copy the generated **Web App URL**, and open it in your browser.

---

## 🌐 Hosting on GitHub (`https://github.com/tesseractthou-code`)

To host and manage this project repository under your GitHub account:

```bash
cd /Users/abhyuday/Desktop/Tesseract

# Add remote repository link
git remote add origin https://github.com/tesseractthou-code/location-registration-form.git

# Push to GitHub
git push -u origin main
```

---

## 📄 License
This project is open-source and available under the [MIT License](LICENSE).
