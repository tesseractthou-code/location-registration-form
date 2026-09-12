/**
 * ============================================================
 *  CASCADING LOCATION REGISTRATION FORM — Apps Script Web App
 * ============================================================
 * Sheets used (auto-created by setupSheets()):
 *   Responses    - all form submissions
 *   Countries    - worldwide country list
 *   IN_States    - India State/UT list
 *   IN_Districts - State/UT -> District mapping
 *   IN_Pincodes  - District -> Pincode -> Tehsil/Block mapping
 *   Config       - AdminEmail, FormTitle, etc.
 *   RawImport    - scratch sheet used by bulkImportLocationData()
 *
 * SETUP (run once from the Apps Script editor):
 *   1. Select "setupSheets" from the function dropdown, click Run.
 *   2. Approve the authorization prompts.
 *   3. (Optional) Run "seedIndiaSampleData" to load a small demo
 *      dataset so you can test the State->District->Pincode cascade
 *      immediately.
 *   4. Replace the sample data with the full official dataset using
 *      bulkImportLocationData() — see README for instructions.
 *   5. Deploy > New deployment > Web app.
 * ============================================================
 */

var SHEET_NAMES = {
  RESPONSES: 'Responses',
  COUNTRIES: 'Countries',
  STATES: 'IN_States',
  DISTRICTS: 'IN_Districts',
  PINCODES: 'IN_Pincodes',
  CONFIG: 'Config',
  RAW_IMPORT: 'RawImport'
};

var CACHE_SECONDS = 21600; // 6 hours

// ------------------------------------------------------------------
// WEB APP ENTRY POINT & REST API (FOR GITHUB PAGES COMPATIBILITY)
// ------------------------------------------------------------------

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiGetRequest(e.parameter);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(getConfig('FormTitle') || 'Location Registration Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var payload;
  try {
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }
  } catch (err) {
    payload = (e && e.parameter) || {};
  }
  var result = submitForm(payload);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiGetRequest(params) {
  var action = params.action;
  var result;
  try {
    if (action === 'getCountries') {
      result = getCountries();
    } else if (action === 'getStates') {
      result = getStates(params.country);
    } else if (action === 'getDistricts') {
      result = getDistricts(params.state);
    } else if (action === 'getPincodes') {
      result = getPincodes(params.district);
    } else if (action === 'getTehsil') {
      result = getTehsil(params.district, params.pincode);
    } else if (action === 'submitForm') {
      var p = params.payload ? JSON.parse(params.payload) : params;
      result = submitForm(p);
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------------
// SETUP / SEEDING
// ------------------------------------------------------------------

function getSs() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(name, headers) {
  var ss = getSs();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) {
      sh.appendRow(headers);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sh;
}

function setupSheets() {
  getOrCreateSheet(SHEET_NAMES.RESPONSES, [
    'Timestamp', 'Name', 'Phone', 'Email', 'Country', 'State/UT',
    'District', 'Pincode', 'Tehsil/Block', 'Address',
    'Duplicate Phone', 'Duplicate Location', 'Status'
  ]);
  getOrCreateSheet(SHEET_NAMES.COUNTRIES, ['Country']);
  getOrCreateSheet(SHEET_NAMES.STATES, ['State/UT']);
  getOrCreateSheet(SHEET_NAMES.DISTRICTS, ['State/UT', 'District']);
  getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  getOrCreateSheet(SHEET_NAMES.RAW_IMPORT, ['State/UT', 'District', 'Pincode', 'Tehsil/Block']);

  var cfg = getOrCreateSheet(SHEET_NAMES.CONFIG, ['Key', 'Value']);
  if (cfg.getLastRow() < 2) {
    cfg.appendRow(['AdminEmail', Session.getEffectiveUser().getEmail()]);
    cfg.appendRow(['FormTitle', 'Location Registration Form']);
  }

  seedCountries();
  SpreadsheetApp.flush();
  return 'Setup complete.';
}

function seedCountries() {
  var sh = getOrCreateSheet(SHEET_NAMES.COUNTRIES, ['Country']);
  if (sh.getLastRow() > 1) return; // already seeded, don't overwrite edits
  var countries = [
    'Afghanistan','Albania','Algeria','Andorra','Angola','Antigua and Barbuda','Argentina','Armenia',
    'Australia','Austria','Azerbaijan','Bahamas','Bahrain','Bangladesh','Barbados','Belarus','Belgium',
    'Belize','Benin','Bhutan','Bolivia','Bosnia and Herzegovina','Botswana','Brazil','Brunei','Bulgaria',
    'Burkina Faso','Burundi','Cabo Verde','Cambodia','Cameroon','Canada','Central African Republic','Chad',
    'Chile','China','Colombia','Comoros','Congo (Congo-Brazzaville)','Costa Rica','Croatia','Cuba','Cyprus',
    'Czechia','Democratic Republic of the Congo','Denmark','Djibouti','Dominica','Dominican Republic',
    'Ecuador','Egypt','El Salvador','Equatorial Guinea','Eritrea','Estonia','Eswatini','Ethiopia','Fiji',
    'Finland','France','Gabon','Gambia','Georgia','Germany','Ghana','Greece','Grenada','Guatemala','Guinea',
    'Guinea-Bissau','Guyana','Haiti','Honduras','Hungary','Iceland','India','Indonesia','Iran','Iraq',
    'Ireland','Israel','Italy','Jamaica','Japan','Jordan','Kazakhstan','Kenya','Kiribati','Kuwait',
    'Kyrgyzstan','Laos','Latvia','Lebanon','Lesotho','Liberia','Libya','Liechtenstein','Lithuania',
    'Luxembourg','Madagascar','Malawi','Malaysia','Maldives','Mali','Malta','Marshall Islands','Mauritania',
    'Mauritius','Mexico','Micronesia','Moldova','Monaco','Mongolia','Montenegro','Morocco','Mozambique',
    'Myanmar','Namibia','Nauru','Nepal','Netherlands','New Zealand','Nicaragua','Niger','Nigeria',
    'North Korea','North Macedonia','Norway','Oman','Pakistan','Palau','Palestine','Panama',
    'Papua New Guinea','Paraguay','Peru','Philippines','Poland','Portugal','Qatar','Romania','Russia',
    'Rwanda','Saint Kitts and Nevis','Saint Lucia','Saint Vincent and the Grenadines','Samoa','San Marino',
    'Sao Tome and Principe','Saudi Arabia','Senegal','Serbia','Seychelles','Sierra Leone','Singapore',
    'Slovakia','Slovenia','Solomon Islands','Somalia','South Africa','South Korea','South Sudan','Spain',
    'Sri Lanka','Sudan','Suriname','Sweden','Switzerland','Syria','Taiwan','Tajikistan','Tanzania',
    'Thailand','Timor-Leste','Togo','Tonga','Trinidad and Tobago','Tunisia','Turkey','Turkmenistan',
    'Tuvalu','Uganda','Ukraine','United Arab Emirates','United Kingdom','United States','Uruguay',
    'Uzbekistan','Vanuatu','Vatican City','Venezuela','Vietnam','Yemen','Zambia','Zimbabwe'
  ];
  var rows = countries.map(function (c) { return [c]; });
  sh.getRange(2, 1, rows.length, 1).setValues(rows);
}

/**
 * Small demo dataset so the cascade is testable immediately.
 * REPLACE with the full official dataset via bulkImportLocationData().
 */
function seedIndiaSampleData() {
  var states = getOrCreateSheet(SHEET_NAMES.STATES, ['State/UT']);
  var districts = getOrCreateSheet(SHEET_NAMES.DISTRICTS, ['State/UT', 'District']);
  var pincodes = getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);

  if (states.getLastRow() < 2) {
    var stateList = [
      'Andhra Pradesh','Arunachal Pradesh','Assam','Bihar','Chhattisgarh','Goa','Gujarat','Haryana',
      'Himachal Pradesh','Jharkhand','Karnataka','Kerala','Madhya Pradesh','Maharashtra','Manipur',
      'Meghalaya','Mizoram','Nagaland','Odisha','Punjab','Rajasthan','Sikkim','Tamil Nadu','Telangana',
      'Tripura','Uttar Pradesh','Uttarakhand','West Bengal',
      'Andaman and Nicobar Islands','Chandigarh','Dadra and Nagar Haveli and Daman and Diu','Delhi',
      'Jammu and Kashmir','Ladakh','Lakshadweep','Puducherry'
    ];
    states.getRange(2, 1, stateList.length, 1).setValues(stateList.map(function (s) { return [s]; }));
  }

  if (districts.getLastRow() < 2) {
    var districtRows = [
      ['Gujarat', 'Ahmedabad'], ['Gujarat', 'Surat'], ['Gujarat', 'Vadodara'],
      ['Maharashtra', 'Mumbai City'], ['Maharashtra', 'Pune'], ['Maharashtra', 'Nagpur'],
      ['Delhi', 'New Delhi'], ['Delhi', 'North Delhi'],
      ['Karnataka', 'Bengaluru Urban'], ['Karnataka', 'Mysuru'],
      ['Tamil Nadu', 'Chennai'], ['Tamil Nadu', 'Coimbatore']
    ];
    districts.getRange(districts.getLastRow() + 1, 1, districtRows.length, 2).setValues(districtRows);
  }

  if (pincodes.getLastRow() < 2) {
    var pinRows = [
      ['Ahmedabad', '380001', 'Ahmedabad City', 'Gujarat'],
      ['Ahmedabad', '380006', 'Ellisbridge', 'Gujarat'],
      ['Ahmedabad', '382480', 'Sanand', 'Gujarat'],
      ['Surat', '395001', 'Surat City', 'Gujarat'],
      ['Vadodara', '390001', 'Vadodara City', 'Gujarat'],
      ['Mumbai City', '400001', 'Fort', 'Maharashtra'],
      ['Mumbai City', '400050', 'Bandra', 'Maharashtra'],
      ['Pune', '411001', 'Pune City', 'Maharashtra'],
      ['New Delhi', '110001', 'Connaught Place', 'Delhi'],
      ['North Delhi', '110007', 'Civil Lines', 'Delhi'],
      ['Bengaluru Urban', '560001', 'Bengaluru North', 'Karnataka'],
      ['Mysuru', '570001', 'Mysuru City', 'Karnataka'],
      ['Chennai', '600001', 'Chennai City', 'Tamil Nadu'],
      ['Coimbatore', '641001', 'Coimbatore North', 'Tamil Nadu']
    ];
    pincodes.getRange(pincodes.getLastRow() + 1, 1, pinRows.length, 4).setValues(pinRows);
  }

  clearLocationCache();
  return 'Sample India location data loaded.';
}

/**
 * Bulk import full/official location data.
 * 1. Paste State/UT, District, Pincode, Tehsil/Block columns into the
 *    "RawImport" sheet (with header row matching those 4 columns).
 * 2. Run this function once from the Apps Script editor.
 * It dedupes and rebuilds IN_States, IN_Districts, IN_Pincodes.
 */
function bulkImportLocationData() {
  var raw = getOrCreateSheet(SHEET_NAMES.RAW_IMPORT, ['State/UT', 'District', 'Pincode', 'Tehsil/Block']);
  var data = raw.getDataRange().getValues();
  if (data.length < 2) throw new Error('RawImport sheet is empty. Paste data first.');

  var stateSet = {}, districtSet = {}, districtRows = [], pinSet = {}, pinRows = [];

  for (var i = 1; i < data.length; i++) {
    var state = String(data[i][0] || '').trim();
    var district = String(data[i][1] || '').trim();
    var pincode = String(data[i][2] || '').trim();
    var tehsil = String(data[i][3] || '').trim();
    if (!state || !district || !pincode) continue;

    if (!stateSet[state]) stateSet[state] = true;

    var dKey = state + '||' + district;
    if (!districtSet[dKey]) { districtSet[dKey] = true; districtRows.push([state, district]); }

    var pKey = district + '||' + pincode;
    if (!pinSet[pKey]) { pinSet[pKey] = true; pinRows.push([district, pincode, tehsil, state]); }
  }

  var statesSh = getOrCreateSheet(SHEET_NAMES.STATES, ['State/UT']);
  statesSh.clearContents();
  statesSh.appendRow(['State/UT']);
  var stateRows = Object.keys(stateSet).sort().map(function (s) { return [s]; });
  if (stateRows.length) statesSh.getRange(2, 1, stateRows.length, 1).setValues(stateRows);

  var districtsSh = getOrCreateSheet(SHEET_NAMES.DISTRICTS, ['State/UT', 'District']);
  districtsSh.clearContents();
  districtsSh.appendRow(['State/UT', 'District']);
  if (districtRows.length) districtsSh.getRange(2, 1, districtRows.length, 2).setValues(districtRows);

  var pinSh = getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  pinSh.clearContents();
  pinSh.appendRow(['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  if (pinRows.length) pinSh.getRange(2, 1, pinRows.length, 4).setValues(pinRows);

  clearLocationCache();
  return 'Imported ' + stateRows.length + ' states, ' + districtRows.length + ' districts, ' + pinRows.length + ' pincodes.';
}

function clearLocationCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['countries_list', 'in_states_list']);
}

// ------------------------------------------------------------------
// CONFIG
// ------------------------------------------------------------------

function getConfig(key) {
  var sh = getOrCreateSheet(SHEET_NAMES.CONFIG, ['Key', 'Value']);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

// ------------------------------------------------------------------
// DATA FUNCTIONS CALLED FROM THE CLIENT (google.script.run)
// ------------------------------------------------------------------

function uniqueSorted(arr) {
  return Array.from(new Set(arr)).filter(function (v) { return v !== '' && v !== null; }).sort();
}

function getCountries() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('countries_list');
  if (cached) return JSON.parse(cached);

  var sh = getOrCreateSheet(SHEET_NAMES.COUNTRIES, ['Country']);
  var n = Math.max(sh.getLastRow() - 1, 0);
  var values = n ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return r[0]; }) : [];
  values = uniqueSorted(values);
  cache.put('countries_list', JSON.stringify(values), CACHE_SECONDS);
  return values;
}

function getStates(country) {
  if (country !== 'India') return [];
  var cache = CacheService.getScriptCache();
  var cached = cache.get('in_states_list');
  if (cached) return JSON.parse(cached);

  var sh = getOrCreateSheet(SHEET_NAMES.STATES, ['State/UT']);
  var n = Math.max(sh.getLastRow() - 1, 0);
  var values = n ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return r[0]; }) : [];
  values = uniqueSorted(values);
  cache.put('in_states_list', JSON.stringify(values), CACHE_SECONDS);
  return values;
}

function getDistricts(state) {
  if (!state) return [];
  var sh = getOrCreateSheet(SHEET_NAMES.DISTRICTS, ['State/UT', 'District']);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === state) out.push(data[i][1]);
  }
  return uniqueSorted(out);
}

function getPincodes(district) {
  if (!district) return [];
  var sh = getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === district) out.push(String(data[i][1]));
  }
  return uniqueSorted(out);
}

function getTehsil(district, pincode) {
  if (!district || !pincode) return '';
  var sh = getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === district && String(data[i][1]) === String(pincode)) return data[i][2];
  }
  return '';
}

// ------------------------------------------------------------------
// SUBMISSION + VALIDATION
// ------------------------------------------------------------------

function normalizePhone(phone) {
  return String(phone || '').replace(/\D/g, '');
}

function normalizeText(str) {
  return String(str || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function submitForm(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { success: false, error: 'busy', message: 'Server is busy, please try again in a moment.' };
  }

  try {
    var required = ['name', 'phone', 'country'];
    for (var i = 0; i < required.length; i++) {
      if (!payload[required[i]] || String(payload[required[i]]).trim() === '') {
        return { success: false, error: 'validation', message: 'Missing required field: ' + required[i] };
      }
    }

    var phoneNorm = normalizePhone(payload.phone);
    if (phoneNorm.length < 7) {
      return { success: false, error: 'validation', message: 'Please enter a valid phone number.' };
    }

    var sh = getOrCreateSheet(SHEET_NAMES.RESPONSES);
    var data = sh.getDataRange().getValues();
    var header = data[0];
    var col = {};
    header.forEach(function (h, idx) { col[h] = idx; });

    var duplicatePhone = false;
    var duplicateLocation = false;
    var addrNorm = normalizeText(payload.address);
    var thisLocationKey = [
      normalizeText(payload.country), normalizeText(payload.state),
      normalizeText(payload.district), normalizeText(payload.pincode), addrNorm
    ].join('|');

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (normalizePhone(row[col['Phone']]) === phoneNorm) duplicatePhone = true;

      var rowLocationKey = [
        normalizeText(row[col['Country']]), normalizeText(row[col['State/UT']]),
        normalizeText(row[col['District']]), normalizeText(row[col['Pincode']]),
        normalizeText(row[col['Address']])
      ].join('|');
      if (addrNorm && rowLocationKey === thisLocationKey) duplicateLocation = true;
    }

    if (duplicatePhone) {
      return { success: false, error: 'duplicate_phone', message: 'This phone number has already been registered.' };
    }

    var tehsil = payload.tehsil || (payload.district && payload.pincode ? getTehsil(payload.district, payload.pincode) : '');

    var newRow = [
      new Date(), payload.name, payload.phone, payload.email || '', payload.country,
      payload.state || '', payload.district || '', payload.pincode || '', tehsil, payload.address || '',
      duplicatePhone ? 'YES' : 'NO', duplicateLocation ? 'YES' : 'NO', duplicateLocation ? 'Flagged' : 'New'
    ];
    sh.appendRow(newRow);

    if (duplicateLocation) {
      sh.getRange(sh.getLastRow(), 1, 1, newRow.length).setBackground('#FFF3CD');
    }

    notifyAdmin(payload, duplicateLocation);

    return { success: true, duplicateLocation: duplicateLocation };
  } catch (err) {
    return { success: false, error: 'server_error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function notifyAdmin(payload, duplicateLocation) {
  var adminEmail = getConfig('AdminEmail');
  if (!adminEmail) return;

  var subject = 'New Form Submission' + (duplicateLocation ? ' (Duplicate Location Flagged)' : '');
  var lines = [
    'Name: ' + payload.name,
    'Phone: ' + payload.phone,
    'Email: ' + (payload.email || '-'),
    'Country: ' + payload.country,
    'State/UT: ' + (payload.state || '-'),
    'District: ' + (payload.district || '-'),
    'Pincode: ' + (payload.pincode || '-'),
    'Tehsil/Block: ' + (payload.tehsil || '-'),
    'Address: ' + (payload.address || '-')
  ];
  if (duplicateLocation) {
    lines.push('', '⚠ This location (country/state/district/pincode/address) matches an existing submission. The row has been highlighted in the sheet.');
  }

  MailApp.sendEmail(adminEmail, subject, lines.join('\n'));
}
