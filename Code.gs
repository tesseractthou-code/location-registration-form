/**
 * ============================================================
 *  CASCADING LOCATION REGISTRATION FORM — Apps Script Web App
 * ============================================================
 * Sheets used:
 *   Responses    - all form submissions (can be stored in external Submissions Database)
 *   Countries    - worldwide country list
 *   IN_States    - India State/UT list
 *   IN_Districts - State/UT -> District mapping
 *   Sub_Districts- State/UT -> District -> Tehsil mapping
 *   IN_Pincodes  - District -> Pincode -> Tehsil/Block mapping
 *   Config       - AdminEmail, FormTitle, SubmissionsSpreadsheetId
 *   RawImport    - scratch sheet used by bulkImportLocationData()
 * ============================================================
 */

var SHEET_NAMES = {
  RESPONSES: 'Responses',
  COUNTRIES: 'Countries',
  STATES: 'IN_States',
  DISTRICTS: 'IN_Districts',
  SUB_DISTRICTS: 'Sub_Districts',
  PINCODES: 'IN_Pincodes',
  CONFIG: 'Config',
  RAW_IMPORT: 'RawImport'
};

var CACHE_SECONDS = 21600; // 6 hours

// ------------------------------------------------------------------
// SPREADSHEET & SHEET UTILITIES
// ------------------------------------------------------------------

function getSs() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSubmissionsSs() {
  var id = getConfig('SubmissionsSpreadsheetId');
  if (id && String(id).trim() !== '') {
    var cleanId = String(id).trim();
    // Extract ID if user pasted full Google Sheet URL
    var match = cleanId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
      cleanId = match[1];
    }
    try {
      var extSs = SpreadsheetApp.openById(cleanId);
      if (extSs) {
        return extSs;
      }
    } catch (e) {
      Logger.log('Could not open SubmissionsSpreadsheetId (' + cleanId + '): ' + e.message);
      throw new Error('Form Submissions Database access error: ' + e.message + '. Ensure Web App deployment is set to "Execute as: Me".');
    }
  }
  return getSs();
}

function getConfig(key) {
  var sh = getOrCreateSheet(SHEET_NAMES.CONFIG, ['Key', 'Value']);
  var data = sh.getDataRange().getValues();
  var target = String(key || '').trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][0] || '').trim().toLowerCase();
    if (k === target) return data[i][1];
  }
  return null;
}

function getOrCreateSheetInSs(ss, name, headers) {
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

function getOrCreateSheet(name, headers) {
  return getOrCreateSheetInSs(getSs(), name, headers);
}

// ------------------------------------------------------------------
// INITIAL SETUP
// ------------------------------------------------------------------

function setupSheets() {
  var subSs = getSubmissionsSs();
  getOrCreateSheetInSs(subSs, SHEET_NAMES.RESPONSES, [
    'Timestamp', 'Name', 'Phone', 'Email', 'Country', 'State/UT',
    'District', 'Sub-district/Tehsil/Block', 'Pincode', 'Address',
    'Duplicate Phone', 'Duplicate Location', 'Status'
  ]);
  getOrCreateSheet(SHEET_NAMES.COUNTRIES, ['Country']);
  getOrCreateSheet(SHEET_NAMES.STATES, ['State/UT']);
  getOrCreateSheet(SHEET_NAMES.DISTRICTS, ['State/UT', 'District']);
  getOrCreateSheet(SHEET_NAMES.SUB_DISTRICTS, ['State/UT', 'District', 'Sub-district/Tehsil/Block']);
  getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  getOrCreateSheet(SHEET_NAMES.RAW_IMPORT, ['State/UT', 'District', 'Pincode', 'Tehsil/Block']);

  var cfg = getOrCreateSheet(SHEET_NAMES.CONFIG, ['Key', 'Value']);
  var data = cfg.getDataRange().getValues();
  var existingKeys = {};
  for (var i = 1; i < data.length; i++) {
    existingKeys[String(data[i][0] || '').trim().toLowerCase()] = true;
  }

  if (!existingKeys['adminemail']) {
    cfg.appendRow(['AdminEmail', Session.getEffectiveUser().getEmail()]);
  }
  if (!existingKeys['formtitle']) {
    cfg.appendRow(['FormTitle', 'Location Registration Form']);
  }
  if (!existingKeys['submissionsspreadsheetid']) {
    cfg.appendRow(['SubmissionsSpreadsheetId', '']);
  }

  seedCountries();
  SpreadsheetApp.flush();
  return 'Setup complete.';
}

// ------------------------------------------------------------------
// WEB APP ENDPOINTS
// ------------------------------------------------------------------

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiGetRequest(e.parameter);
  }
  var template;
  try {
    template = HtmlService.createTemplateFromFile('index');
  } catch (err) {
    template = HtmlService.createTemplateFromFile('Index');
  }
  return template
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
    } else if (action === 'getTehsils') {
      result = getTehsils(params.district);
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
// DATA SEEDING & BULK IMPORT
// ------------------------------------------------------------------

function seedCountries() {
  var sh = getOrCreateSheet(SHEET_NAMES.COUNTRIES, ['Country']);
  if (sh.getLastRow() > 1) return;
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
    var allDistrictsMap = {
      "Andaman and Nicobar Islands": ["Nicobar", "North And Middle Andaman", "South Andaman"],
      "Andhra Pradesh": ["Ananthapur", "Chittoor", "Cuddapah", "East Godavari", "Guntur", "Krishna", "Kurnool", "Nellore", "Prakasam", "Srikakulam", "Visakhapatnam", "Vizianagaram", "West Godavari"],
      "Arunachal Pradesh": ["Changlang", "Dibang Valley", "East Kameng", "East Siang", "Kurung Kumey", "Lohit", "Lower Dibang Valley", "Lower Subansiri", "Papum Pare", "Tawang", "Tirap", "Upper Siang", "Upper Subansiri", "West Kameng", "West Siang"],
      "Assam": ["Barpeta", "Bongaigaon", "Cachar", "Darrang", "Dhemaji", "Dhubri", "Dibrugarh", "Goalpara", "Golaghat", "Hailakandi", "Jorhat", "Kamrup", "Karbi Anglong", "Karimganj", "Kokrajhar", "Lakhimpur", "Marigaon", "Nagaon", "Nalbari", "North Cachar Hills", "Sibsagar", "Sonitpur", "Tinsukia"],
      "Bihar": ["Araria", "Arwal", "Aurangabad", "Banka", "Begusarai", "Bhagalpur", "Bhojpur", "Buxar", "Darbhanga", "East Champaran", "Gaya", "Gopalganj", "Jamui", "Jehanabad", "Kaimur (Bhabua)", "Katihar", "Khagaria", "Kishanganj", "Lakhisarai", "Madhepura", "Madhubani", "Munger", "Muzaffarpur", "Nalanda", "Nawada", "Patna", "Purnia", "Rohtas", "Saharsa", "Samastipur", "Saran", "Sheikhpura", "Sheohar", "Sitamarhi", "Siwan", "Supaul", "Vaishali", "West Champaran"],
      "Chandigarh": ["Chandigarh"],
      "Chhattisgarh": ["Bastar", "Bilas Pur", "Dantewada", "Dhamtari", "Durg", "Janagir-Champa", "Jashpur", "Kanker", "Kawardha", "Korba", "Koriya", "Mahasamund", "Raigarh", "Raipur", "Rajnandgaon", "Surguja"],
      "Dadra and Nagar Haveli and Daman and Diu": ["Dadra and Nagar Haveli", "Daman", "Diu"],
      "Delhi": ["Central Delhi", "East Delhi", "New Delhi", "North Delhi", "North East Delhi", "North West Delhi", "South Delhi", "South West Delhi", "West Delhi"],
      "Goa": ["North Goa", "South Goa"],
      "Gujarat": ["Ahmedabad", "Amreli", "Anand", "Banaskantha", "Bharuch", "Bhavnagar", "Dahod", "Gandhinagar", "Jamnagar", "Junagadh", "Kheda", "Kutch", "Mehsana", "Narmada", "Navsari", "Panchmahal", "Patan", "Porbandar", "Rajkot", "Sabarkantha", "Surat", "Surendranagar", "Tapi", "The Dangs", "Vadodara", "Valsad"],
      "Haryana": ["Ambala", "Bhiwani", "Faridabad", "Fatehabad", "Gurgaon", "Hisar", "Jhajjar", "Jind", "Kaithal", "Karnal", "Kurukshetra", "Mahendragarh", "Mewat", "Palwal", "Panchkula", "Panipat", "Rewari", "Rohtak", "Sirsa", "Sonipat", "Yamunanagar"],
      "Himachal Pradesh": ["Bilaspur", "Chamba", "Hamirpur", "Kangra", "Kinnaur", "Kullu", "Lahaul and Spiti", "Mandi", "Shimla", "Sirmaur", "Solan", "Una"],
      "Jammu and Kashmir": ["Anantnag", "Badgam", "Bandipore", "Baramulla", "Doda", "Ganderbal", "Jammu", "Kathua", "Kishtwar", "Kulgam", "Kupwara", "Poonch", "Pulwama", "Rajouri", "Ramban", "Reasi", "Samba", "Shopian", "Srinagar", "Udhampur"],
      "Jharkhand": ["Bokaro", "Chatra", "Deoghar", "Dhanbad", "Dumka", "East Singhbhum", "Garhwa", "Giridih", "Godda", "Gumla", "Hazaribag", "Jamtara", "Khunti", "Koderma", "Latehar", "Lohardaga", "Pakur", "Palamu", "Ramgarh", "Ranchi", "Sahibganj", "Seraikela Kharsawan", "Simdega", "West Singhbhum"],
      "Karnataka": ["Bagalkot", "Bangalore", "Bangalore Rural", "Belgaum", "Bellary", "Bidar", "Bijapur", "Chamarajanagar", "Chikkaballapur", "Chikmagalur", "Chitradurga", "Dakshina Kannada", "Davanagere", "Dharwad", "Gadag", "Gulbarga", "Hassan", "Haveri", "Kodagu", "Kolar", "Koppal", "Mandya", "Mysore", "Raichur", "Ramanagara", "Shimoga", "Tumkur", "Udupi", "Uttara Kannada", "Yadgir"],
      "Kerala": ["Alappuzha", "Ernakulam", "Idukki", "Kannur", "Kasaragod", "Kollam", "Kottayam", "Kozhikode", "Malappuram", "Palakkad", "Pathanamthitta", "Thiruvananthapuram", "Thrissur", "Wayanad"],
      "Ladakh": ["Kargil", "Leh"],
      "Lakshadweep": ["Lakshadweep"],
      "Madhya Pradesh": ["Alirajpur", "Anuppur", "Ashoknagar", "Balaghat", "Barwani", "Betul", "Bhind", "Bhopal", "Burhanpur", "Chhatarpur", "Chhindwara", "Damoh", "Datia", "Dewas", "Dhar", "Dindori", "Guna", "Gwalior", "Harda", "Hoshangabad", "Indore", "Jabalpur", "Jhabua", "Katni", "Khandwa", "Khargone", "Mandla", "Mandsaur", "Morena", "Narsinghpur", "Neemuch", "Panna", "Raisen", "Rajgarh", "Ratlam", "Rewa", "Sagar", "Satna", "Sehore", "Seoni", "Shahdol", "Shajapur", "Sheopur", "Shivpuri", "Sidhi", "Singrauli", "Tikamgarh", "Ujjain", "Umaria", "Vidisha"],
      "Maharashtra": ["Ahmednagar", "Akola", "Amravati", "Aurangabad", "Beed", "Bhandara", "Buldhana", "Chandrapur", "Dhule", "Gadchiroli", "Gondia", "Hingoli", "Jalgaon", "Jalna", "Kolhapur", "Latur", "Mumbai City", "Mumbai Suburban", "Nagpur", "Nanded", "Nandurbar", "Nashik", "Osmanabad", "Palghar", "Parbhani", "Pune", "Raigad", "Ratnagiri", "Sangli", "Satara", "Sindhudurg", "Solapur", "Thane", "Wardha", "Washim", "Yavatmal"],
      "Manipur": ["Bishnupur", "Chandel", "Churachandpur", "Imphal East", "Imphal West", "Senapati", "Tamenglong", "Thoubal", "Ukhrul"],
      "Meghalaya": ["East Garo Hills", "East Khasi Hills", "Jaintia Hills", "Ri Bhoi", "South Garo Hills", "West Garo Hills", "West Khasi Hills"],
      "Mizoram": ["Aizawl", "Champhai", "Kolasib", "Lawngtlai", "Lunglei", "Mamit", "Saiha", "Serchhip"],
      "Nagaland": ["Dimapur", "Kiphire", "Kohima", "Longleng", "Mokokchung", "Mon", "Peren", "Phek", "Tuensang", "Wokha", "Zunheboto"],
      "Odisha": ["Angul", "Balangir", "Balasore", "Bargarh", "Bhadrak", "Boudh", "Cuttack", "Deogarh", "Dhenkanal", "Gajapati", "Ganjam", "Jagatsinghapur", "Jajapur", "Jharsuguda", "Kalahandi", "Kandhamal", "Kendrapara", "Kendujhar", "Khordha", "Koraput", "Malkangiri", "Mayurbhanj", "Nabarangpur", "Nayagarh", "Nuapada", "Puri", "Rayagada", "Sambalpur", "Subarnapur", "Sundergarh"],
      "Puducherry": ["Karaikal", "Mahe", "Puducherry", "Yanam"],
      "Punjab": ["Amritsar", "Barnala", "Bathinda", "Faridkot", "Fatehgarh Sahib", "Fazilka", "Firozpur", "Gurdaspur", "Hoshiarpur", "Jalandhar", "Kapurthala", "Ludhiana", "Mansa", "Moga", "Muktsar", "Pathankot", "Patiala", "Rupnagar", "Sahibzada Ajit Singh Nagar", "Sangrur", "Shahid Bhagat Singh Nagar", "Tarn Taran"],
      "Rajasthan": ["Ajmer", "Alwar", "Banswara", "Baran", "Barmer", "Bharatpur", "Bhilwara", "Bikaner", "Bundi", "Chittorgarh", "Churu", "Dausa", "Dholpur", "Dungarpur", "Ganganagar", "Hanumangarh", "Jaipur", "Jaisalmer", "Jalor", "Jhalawar", "Jhunjhunu", "Jodhpur", "Karauli", "Kota", "Nagaur", "Pali", "Rajsamand", "Sawai Madhopur", "Sikar", "Sirohi", "Tonk", "Udaipur"],
      "Sikkim": ["East Sikkim", "North Sikkim", "South Sikkim", "West Sikkim"],
      "Tamil Nadu": ["Ariyalur", "Chennai", "Coimbatore", "Cuddalore", "Dharmapuri", "Dindigul", "Erode", "Kanchipuram", "Kanyakumari", "Karur", "Krishnagiri", "Madurai", "Nagapattinam", "Namakkal", "Nilgiris", "Perambalur", "Pudukkottai", "Ramanathapuram", "Salem", "Sivaganga", "Thanjavur", "Theni", "Tiruchirappalli", "Tirunelveli", "Tiruvallur", "Tiruvannamalai", "Tiruvarur", "Tuticorin", "Vellore", "Villupuram", "Virudhunagar"],
      "Telangana": ["Adilabad", "Bhadradri Kothagudem", "Hyderabad", "Jagtial", "Jangaon", "Jayashankar Bhupalpally", "Jogulamba Gadwal", "Kamareddy", "Karimnagar", "Khammam", "Kumuram Bheem", "Mahabubabad", "Mahabubnagar", "Mancherial", "Medak", "Medchal", "Nagarkurnool", "Nalgonda", "Nirmal", "Nizamabad", "Peddapalli", "Rajanna Sircilla", "Rangareddy", "Sangareddy", "Siddipet", "Suryapet", "Vikarabad", "Wanaparthy", "Warangal Rural", "Warangal Urban", "Yadadri Bhuvanagiri"],
      "Tripura": ["Dhalai", "Gomati", "Khowai", "North Tripura", "Sepahijala", "South Tripura", "Unakoti", "West Tripura"],
      "Uttar Pradesh": ["Agra", "Aligarh", "Allahabad", "Ambedkar Nagar", "Amethi", "Amroha", "Auraiya", "Azamgarh", "Baghpat", "Bahraich", "Ballia", "Balrampur", "Banda", "Barabanki", "Bareilly", "Basti", "Bhadohi", "Bijnor", "Budaun", "Bulandshahr", "Chandauli", "Chitrakoot", "Deoria", "Etah", "Etawah", "Farrukhabad", "Fatehpur", "Firozabad", "Gautam Buddha Nagar", "Ghaziabad", "Ghazipur", "Gonda", "Gorakhpur", "Hamirpur", "Hapur", "Hardoi", "Hathras", "Jalaun", "Jaunpur", "Jhansi", "Kannauj", "Kanpur Dehat", "Kanpur Nagar", "Kasganj", "Kaushambi", "Kheri", "Kushinagar", "Lalitpur", "Lucknow", "Maharajganj", "Mahoba", "Mainpuri", "Mathura", "Mau", "Meerut", "Mirzapur", "Moradabad", "Muzaffarnagar", "Pilibhit", "Pratapgarh", "Raebareli", "Rampur", "Saharanpur", "Sambhal", "Sant Kabir Nagar", "Shahjahanpur", "Shamli", "Shrawasti", "Siddharthnagar", "Sitapur", "Sonbhadra", "Sultanpur", "Unnao", "Varanasi"],
      "Uttarakhand": ["Almora", "Bageshwar", "Chamoli", "Champawat", "Dehradun", "Haridwar", "Nainital", "Pauri Garhwal", "Pithoragarh", "Rudraprayag", "Tehri Garhwal", "Udham Singh Nagar", "Uttarkashi"],
      "West Bengal": ["Alipurduar", "Bankura", "Birbhum", "Cooch Behar", "Dakshin Dinajpur", "Darjeeling", "Hooghly", "Howrah", "Jalpaiguri", "Jhargram", "Kalimpong", "Kolkata", "Malda", "Murshidabad", "Nadia", "North 24 Parganas", "Paschim Bardhaman", "Paschim Medinipur", "Purba Bardhaman", "Purba Medinipur", "Purulia", "South 24 Parganas", "Uttar Dinajpur"]
    };
    var districtRows = [];
    for (var stKey in allDistrictsMap) {
      var dList = allDistrictsMap[stKey];
      for (var dIdx = 0; dIdx < dList.length; dIdx++) {
        districtRows.push([stKey, dList[dIdx]]);
      }
    }
    districts.getRange(2, 1, districtRows.length, 2).setValues(districtRows);
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

function bulkImportLocationData() {
  var raw = getOrCreateSheet(SHEET_NAMES.RAW_IMPORT, ['State/UT', 'District', 'Pincode', 'Tehsil/Block']);
  var data = raw.getDataRange().getValues();
  if (data.length < 2) throw new Error('RawImport sheet is empty. Paste data first.');

  var stateSet = {}, districtSet = {}, districtRows = [], subDistrictSet = {}, subDistrictRows = [], pinSet = {}, pinRows = [];

  for (var i = 1; i < data.length; i++) {
    var state = String(data[i][0] || '').trim();
    var district = String(data[i][1] || '').trim();
    var pincode = String(data[i][2] || '').trim();
    var tehsil = String(data[i][3] || '').trim();
    if (!state || !district || !pincode) continue;

    if (!stateSet[state]) stateSet[state] = true;

    var dKey = state + '||' + district;
    if (!districtSet[dKey]) { districtSet[dKey] = true; districtRows.push([state, district]); }

    if (tehsil) {
      var sdKey = state + '||' + district + '||' + tehsil;
      if (!subDistrictSet[sdKey]) { subDistrictSet[sdKey] = true; subDistrictRows.push([state, district, tehsil]); }
    }

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

  var subDistSh = getOrCreateSheet(SHEET_NAMES.SUB_DISTRICTS, ['State/UT', 'District', 'Sub-district/Tehsil/Block']);
  subDistSh.clearContents();
  subDistSh.appendRow(['State/UT', 'District', 'Sub-district/Tehsil/Block']);
  if (subDistrictRows.length) subDistSh.getRange(2, 1, subDistrictRows.length, 3).setValues(subDistrictRows);

  var pinSh = getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  pinSh.clearContents();
  pinSh.appendRow(['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  if (pinRows.length) pinSh.getRange(2, 1, pinRows.length, 4).setValues(pinRows);

  clearLocationCache();
  return 'Imported ' + stateRows.length + ' states, ' + districtRows.length + ' districts, ' + subDistrictRows.length + ' sub-districts, and ' + pinRows.length + ' pincodes.';
}

function clearLocationCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['countries_list', 'in_states_list']);
}

// ------------------------------------------------------------------
// CLIENT DATA FETCHERS
// ------------------------------------------------------------------

function uniqueSorted(arr) {
  return Array.from(new Set(arr)).filter(function (v) { return v !== '' && v !== null; }).sort();
}

function getCountries() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('countries_list');
  if (cached) {
    var parsed = JSON.parse(cached);
    if (parsed && parsed.length) return parsed;
  }

  var sh = getOrCreateSheet(SHEET_NAMES.COUNTRIES, ['Country']);
  if (sh.getLastRow() <= 1) {
    seedCountries();
  }
  var n = Math.max(sh.getLastRow() - 1, 0);
  var values = n ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return r[0]; }) : [];
  values = uniqueSorted(values);

  if (values.length) {
    cache.put('countries_list', JSON.stringify(values), CACHE_SECONDS);
  }
  return values;
}

function getStates(country) {
  if (country !== 'India') return [];
  var cache = CacheService.getScriptCache();
  var cached = cache.get('in_states_list');
  if (cached) {
    var parsed = JSON.parse(cached);
    if (parsed && parsed.length) return parsed;
  }

  var sh = getOrCreateSheet(SHEET_NAMES.STATES, ['State/UT']);
  if (sh.getLastRow() <= 1) {
    seedIndiaSampleData();
  }
  var n = Math.max(sh.getLastRow() - 1, 0);
  var values = n ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return r[0]; }) : [];
  values = uniqueSorted(values);

  if (!values.length) {
    values = [
      'Andaman and Nicobar Islands','Andhra Pradesh','Arunachal Pradesh','Assam','Bihar',
      'Chandigarh','Chhattisgarh','Dadra and Nagar Haveli and Daman and Diu','Delhi','Goa',
      'Gujarat','Haryana','Himachal Pradesh','Jammu and Kashmir','Jharkhand','Karnataka',
      'Kerala','Ladakh','Lakshadweep','Madhya Pradesh','Maharashtra','Manipur','Meghalaya',
      'Mizoram','Nagaland','Odisha','Puducherry','Punjab','Rajasthan','Sikkim','Tamil Nadu',
      'Telangana','Tripura','Uttar Pradesh','Uttarakhand','West Bengal'
    ];
  }

  cache.put('in_states_list', JSON.stringify(values), CACHE_SECONDS);
  return values;
}

function getDistricts(state) {
  if (!state) return [];
  var sh = getOrCreateSheet(SHEET_NAMES.DISTRICTS, ['State/UT', 'District']);
  if (sh.getLastRow() <= 1) {
    seedIndiaSampleData();
  }
  var data = sh.getDataRange().getValues();
  var out = [];
  var targetState = String(state || '').trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var s = String(data[i][0] || '').trim().toLowerCase();
    if (s === targetState && data[i][1]) {
      out.push(String(data[i][1]).trim());
    }
  }
  out = uniqueSorted(out);

  if (!out.length) {
    var allDistrictsMap = {
      "Andaman and Nicobar Islands": ["Nicobar", "North And Middle Andaman", "South Andaman"],
      "Andhra Pradesh": ["Ananthapur", "Chittoor", "Cuddapah", "East Godavari", "Guntur", "Krishna", "Kurnool", "Nellore", "Prakasam", "Srikakulam", "Visakhapatnam", "Vizianagaram", "West Godavari"],
      "Arunachal Pradesh": ["Changlang", "Dibang Valley", "East Kameng", "East Siang", "Kurung Kumey", "Lohit", "Lower Dibang Valley", "Lower Subansiri", "Papum Pare", "Tawang", "Tirap", "Upper Siang", "Upper Subansiri", "West Kameng", "West Siang"],
      "Assam": ["Barpeta", "Bongaigaon", "Cachar", "Darrang", "Dhemaji", "Dhubri", "Dibrugarh", "Goalpara", "Golaghat", "Hailakandi", "Jorhat", "Kamrup", "Karbi Anglong", "Karimganj", "Kokrajhar", "Lakhimpur", "Marigaon", "Nagaon", "Nalbari", "North Cachar Hills", "Sibsagar", "Sonitpur", "Tinsukia"],
      "Bihar": ["Araria", "Arwal", "Aurangabad", "Banka", "Begusarai", "Bhagalpur", "Bhojpur", "Buxar", "Darbhanga", "East Champaran", "Gaya", "Gopalganj", "Jamui", "Jehanabad", "Kaimur (Bhabua)", "Katihar", "Khagaria", "Kishanganj", "Lakhisarai", "Madhepura", "Madhubani", "Munger", "Muzaffarpur", "Nalanda", "Nawada", "Patna", "Purnia", "Rohtas", "Saharsa", "Samastipur", "Saran", "Sheikhpura", "Sheohar", "Sitamarhi", "Siwan", "Supaul", "Vaishali", "West Champaran"],
      "Chandigarh": ["Chandigarh"],
      "Chhattisgarh": ["Bastar", "Bilas Pur", "Dantewada", "Dhamtari", "Durg", "Janagir-Champa", "Jashpur", "Kanker", "Kawardha", "Korba", "Koriya", "Mahasamund", "Raigarh", "Raipur", "Rajnandgaon", "Surguja"],
      "Dadra and Nagar Haveli and Daman and Diu": ["Dadra and Nagar Haveli", "Daman", "Diu"],
      "Delhi": ["Central Delhi", "East Delhi", "New Delhi", "North Delhi", "North East Delhi", "North West Delhi", "South Delhi", "South West Delhi", "West Delhi"],
      "Goa": ["North Goa", "South Goa"],
      "Gujarat": ["Ahmedabad", "Amreli", "Anand", "Banaskantha", "Bharuch", "Bhavnagar", "Dahod", "Gandhinagar", "Jamnagar", "Junagadh", "Kheda", "Kutch", "Mehsana", "Narmada", "Navsari", "Panchmahal", "Patan", "Porbandar", "Rajkot", "Sabarkantha", "Surat", "Surendranagar", "Tapi", "The Dangs", "Vadodara", "Valsad"],
      "Haryana": ["Ambala", "Bhiwani", "Faridabad", "Fatehabad", "Gurgaon", "Hisar", "Jhajjar", "Jind", "Kaithal", "Karnal", "Kurukshetra", "Mahendragarh", "Mewat", "Palwal", "Panchkula", "Panipat", "Rewari", "Rohtak", "Sirsa", "Sonipat", "Yamunanagar"],
      "Himachal Pradesh": ["Bilaspur", "Chamba", "Hamirpur", "Kangra", "Kinnaur", "Kullu", "Lahaul and Spiti", "Mandi", "Shimla", "Sirmaur", "Solan", "Una"],
      "Jammu and Kashmir": ["Anantnag", "Badgam", "Bandipore", "Baramulla", "Doda", "Ganderbal", "Jammu", "Kathua", "Kishtwar", "Kulgam", "Kupwara", "Poonch", "Pulwama", "Rajouri", "Ramban", "Reasi", "Samba", "Shopian", "Srinagar", "Udhampur"],
      "Jharkhand": ["Bokaro", "Chatra", "Deoghar", "Dhanbad", "Dumka", "East Singhbhum", "Garhwa", "Giridih", "Godda", "Gumla", "Hazaribag", "Jamtara", "Khunti", "Koderma", "Latehar", "Lohardaga", "Pakur", "Palamu", "Ramgarh", "Ranchi", "Sahibganj", "Seraikela Kharsawan", "Simdega", "West Singhbhum"],
      "Karnataka": ["Bagalkot", "Bangalore", "Bangalore Rural", "Belgaum", "Bellary", "Bidar", "Bijapur", "Chamarajanagar", "Chikkaballapur", "Chikmagalur", "Chitradurga", "Dakshina Kannada", "Davanagere", "Dharwad", "Gadag", "Gulbarga", "Hassan", "Haveri", "Kodagu", "Kolar", "Koppal", "Mandya", "Mysore", "Raichur", "Ramanagara", "Shimoga", "Tumkur", "Udupi", "Uttara Kannada", "Yadgir"],
      "Kerala": ["Alappuzha", "Ernakulam", "Idukki", "Kannur", "Kasaragod", "Kollam", "Kottayam", "Kozhikode", "Malappuram", "Palakkad", "Pathanamthitta", "Thiruvananthapuram", "Thrissur", "Wayanad"],
      "Ladakh": ["Kargil", "Leh"],
      "Lakshadweep": ["Lakshadweep"],
      "Madhya Pradesh": ["Alirajpur", "Anuppur", "Ashoknagar", "Balaghat", "Barwani", "Betul", "Bhind", "Bhopal", "Burhanpur", "Chhatarpur", "Chhindwara", "Damoh", "Datia", "Dewas", "Dhar", "Dindori", "Guna", "Gwalior", "Harda", "Hoshangabad", "Indore", "Jabalpur", "Jhabua", "Katni", "Khandwa", "Khargone", "Mandla", "Mandsaur", "Morena", "Narsinghpur", "Neemuch", "Panna", "Raisen", "Rajgarh", "Ratlam", "Rewa", "Sagar", "Satna", "Sehore", "Seoni", "Shahdol", "Shajapur", "Sheopur", "Shivpuri", "Sidhi", "Singrauli", "Tikamgarh", "Ujjain", "Umaria", "Vidisha"],
      "Maharashtra": ["Ahmednagar", "Akola", "Amravati", "Aurangabad", "Beed", "Bhandara", "Buldhana", "Chandrapur", "Dhule", "Gadchiroli", "Gondia", "Hingoli", "Jalgaon", "Jalna", "Kolhapur", "Latur", "Mumbai City", "Mumbai Suburban", "Nagpur", "Nanded", "Nandurbar", "Nashik", "Osmanabad", "Palghar", "Parbhani", "Pune", "Raigad", "Ratnagiri", "Sangli", "Satara", "Sindhudurg", "Solapur", "Thane", "Wardha", "Washim", "Yavatmal"],
      "Manipur": ["Bishnupur", "Chandel", "Churachandpur", "Imphal East", "Imphal West", "Senapati", "Tamenglong", "Thoubal", "Ukhrul"],
      "Meghalaya": ["East Garo Hills", "East Khasi Hills", "Jaintia Hills", "Ri Bhoi", "South Garo Hills", "West Garo Hills", "West Khasi Hills"],
      "Mizoram": ["Aizawl", "Champhai", "Kolasib", "Lawngtlai", "Lunglei", "Mamit", "Saiha", "Serchhip"],
      "Nagaland": ["Dimapur", "Kiphire", "Kohima", "Longleng", "Mokokchung", "Mon", "Peren", "Phek", "Tuensang", "Wokha", "Zunheboto"],
      "Odisha": ["Angul", "Balangir", "Balasore", "Bargarh", "Bhadrak", "Boudh", "Cuttack", "Deogarh", "Dhenkanal", "Gajapati", "Ganjam", "Jagatsinghapur", "Jajapur", "Jharsuguda", "Kalahandi", "Kandhamal", "Kendrapara", "Kendujhar", "Khordha", "Koraput", "Malkangiri", "Mayurbhanj", "Nabarangpur", "Nayagarh", "Nuapada", "Puri", "Rayagada", "Sambalpur", "Subarnapur", "Sundergarh"],
      "Puducherry": ["Karaikal", "Mahe", "Puducherry", "Yanam"],
      "Punjab": ["Amritsar", "Barnala", "Bathinda", "Faridkot", "Fatehgarh Sahib", "Fazilka", "Firozpur", "Gurdaspur", "Hoshiarpur", "Jalandhar", "Kapurthala", "Ludhiana", "Mansa", "Moga", "Muktsar", "Pathankot", "Patiala", "Rupnagar", "Sahibzada Ajit Singh Nagar", "Sangrur", "Shahid Bhagat Singh Nagar", "Tarn Taran"],
      "Rajasthan": ["Ajmer", "Alwar", "Banswara", "Baran", "Barmer", "Bharatpur", "Bhilwara", "Bikaner", "Bundi", "Chittorgarh", "Churu", "Dausa", "Dholpur", "Dungarpur", "Ganganagar", "Hanumangarh", "Jaipur", "Jaisalmer", "Jalor", "Jhalawar", "Jhunjhunu", "Jodhpur", "Karauli", "Kota", "Nagaur", "Pali", "Rajsamand", "Sawai Madhopur", "Sikar", "Sirohi", "Tonk", "Udaipur"],
      "Sikkim": ["East Sikkim", "North Sikkim", "South Sikkim", "West Sikkim"],
      "Tamil Nadu": ["Ariyalur", "Chennai", "Coimbatore", "Cuddalore", "Dharmapuri", "Dindigul", "Erode", "Kanchipuram", "Kanyakumari", "Karur", "Krishnagiri", "Madurai", "Nagapattinam", "Namakkal", "Nilgiris", "Perambalur", "Pudukkottai", "Ramanathapuram", "Salem", "Sivaganga", "Thanjavur", "Theni", "Tiruchirappalli", "Tirunelveli", "Tiruvallur", "Tiruvannamalai", "Tiruvarur", "Tuticorin", "Vellore", "Villupuram", "Virudhunagar"],
      "Telangana": ["Adilabad", "Bhadradri Kothagudem", "Hyderabad", "Jagtial", "Jangaon", "Jayashankar Bhupalpally", "Jogulamba Gadwal", "Kamareddy", "Karimnagar", "Khammam", "Kumuram Bheem", "Mahabubabad", "Mahabubnagar", "Mancherial", "Medak", "Medchal", "Nagarkurnool", "Nalgonda", "Nirmal", "Nizamabad", "Peddapalli", "Rajanna Sircilla", "Rangareddy", "Sangareddy", "Siddipet", "Suryapet", "Vikarabad", "Wanaparthy", "Warangal Rural", "Warangal Urban", "Yadadri Bhuvanagiri"],
      "Tripura": ["Dhalai", "Gomati", "Khowai", "North Tripura", "Sepahijala", "South Tripura", "Unakoti", "West Tripura"],
      "Uttar Pradesh": ["Agra", "Aligarh", "Allahabad", "Ambedkar Nagar", "Amethi", "Amroha", "Auraiya", "Azamgarh", "Baghpat", "Bahraich", "Ballia", "Balrampur", "Banda", "Barabanki", "Bareilly", "Basti", "Bhadohi", "Bijnor", "Budaun", "Bulandshahr", "Chandauli", "Chitrakoot", "Deoria", "Etah", "Etawah", "Farrukhabad", "Fatehpur", "Firozabad", "Gautam Buddha Nagar", "Ghaziabad", "Ghazipur", "Gonda", "Gorakhpur", "Hamirpur", "Hapur", "Hardoi", "Hathras", "Jalaun", "Jaunpur", "Jhansi", "Kannauj", "Kanpur Dehat", "Kanpur Nagar", "Kasganj", "Kaushambi", "Kheri", "Kushinagar", "Lalitpur", "Lucknow", "Maharajganj", "Mahoba", "Mainpuri", "Mathura", "Mau", "Meerut", "Mirzapur", "Moradabad", "Muzaffarnagar", "Pilibhit", "Pratapgarh", "Raebareli", "Rampur", "Saharanpur", "Sambhal", "Sant Kabir Nagar", "Shahjahanpur", "Shamli", "Shrawasti", "Siddharthnagar", "Sitapur", "Sonbhadra", "Sultanpur", "Unnao", "Varanasi"],
      "Uttarakhand": ["Almora", "Bageshwar", "Chamoli", "Champawat", "Dehradun", "Haridwar", "Nainital", "Pauri Garhwal", "Pithoragarh", "Rudraprayag", "Tehri Garhwal", "Udham Singh Nagar", "Uttarkashi"],
      "West Bengal": ["Alipurduar", "Bankura", "Birbhum", "Cooch Behar", "Dakshin Dinajpur", "Darjeeling", "Hooghly", "Howrah", "Jalpaiguri", "Jhargram", "Kalimpong", "Kolkata", "Malda", "Murshidabad", "Nadia", "North 24 Parganas", "Paschim Bardhaman", "Paschim Medinipur", "Purba Bardhaman", "Purba Medinipur", "Purulia", "South 24 Parganas", "Uttar Dinajpur"]
    };
    for (var k in allDistrictsMap) {
      if (k.toLowerCase() === targetState) return allDistrictsMap[k];
    }
  }
  return out;
}

var ALL_INDIA_PINCODES = {"Adilabad": ["504001", "504002", "504101", "504102", "504103", "504104"], "Agra": ["202124", "282001", "282002", "282003", "282004", "282005"], "Ahmed Nagar": ["413201", "413204", "413205", "413701", "413702", "413703"], "Ahmedabad": ["380001", "380002", "380004", "380005", "380006", "380007"], "Aizawl": ["796001", "796004", "796005", "796007", "796008", "796009"], "Ajmer": ["305001", "305002", "305003", "305004", "305005", "305007"], "Akola": ["444001", "444002", "444003", "444004", "444005", "444006"], "Alappuzha": ["686102", "686103", "686534", "688001", "688002", "688003"], "Aligarh": ["202001", "202002", "202121", "202122", "202123", "202124"], "Alirajpur": ["457885", "457887", "457888"], "Allahabad": ["211001", "211002", "211003", "211004", "211005", "211006"], "Almora": ["244715", "263138", "263601", "263620", "263621", "263622"], "Alwar": ["301001", "301002", "301018", "301019", "301020", "301021"], "Ambala": ["133001", "133004", "133005", "133006", "133101", "133102"], "Ambedkar Nagar": ["224121", "224122", "224123", "224125", "224127", "224129"], "Amravati": ["442302", "444601", "444602", "444603", "444604", "444605"], "Amreli": ["362730", "364515", "364521", "364522", "364525", "364530"], "Amritsar": ["143001", "143002", "143005", "143006", "143008", "143009"], "Anand": ["387210", "387220", "387240", "387310", "387345", "387710"], "Ananthapur": ["515001", "515002", "515003", "515004", "515005", "515101"], "Ananthnag": ["191103", "192101", "192121", "192122", "192124", "192125"], "Angul": ["759037", "759040", "759100", "759101", "759102", "759103"], "Anuppur": ["484001", "484113", "484116", "484117", "484220", "484224"], "Araria": ["854102", "854311", "854312", "854316", "854318", "854325"], "Ariyalur": ["608703", "608901", "612901", "612902", "612903", "612904"], "Arwal": ["804401", "804402", "804419", "804427", "824115", "824127"], "Ashok Nagar": ["473330", "473331", "473332", "473335", "473440", "473443"], "Auraiya": ["206121", "206122", "206129", "206241", "206243", "206244"], "Aurangabad": ["423701", "423702", "423703", "431001", "431002", "431003"], "Aurangabad(Bh)": ["824101", "824102", "824103", "824111", "824112", "824113"], "Azamgarh": ["221602", "221603", "221706", "223221", "223222", "223223"], "Bagalkot": ["586113", "586125", "587101", "587102", "587103", "587104"], "Bageshwar": ["263619", "263624", "263628", "263630", "263631", "263632"], "Bagpat": ["250101", "250345", "250601", "250606", "250609", "250611"], "Bahraich": ["271801", "271802", "271804", "271821", "271824", "271825"], "Balaghat": ["481001", "481051", "481102", "481105", "481111", "481115"], "Balangir": ["767001", "767002", "767016", "767020", "767021", "767022"], "Baleswar": ["756001", "756002", "756003", "756019", "756020", "756021"], "Ballia": ["221701", "221709", "221711", "221712", "221713", "221715"], "Balrampur": ["271201", "271203", "271204", "271205", "271206", "271207"], "Banaskantha": ["385001", "385010", "385110", "385120", "385130", "385135"], "Banda": ["210001", "210120", "210121", "210123", "210125", "210126"], "Bandipur": ["193503"], "Bangalore": ["560001", "560002", "560003", "560004", "560005", "560006"], "Bangalore Rural": ["560060", "560067", "560068", "560082", "560083", "560088"], "Banka": ["811308", "813101", "813102", "813103", "813104", "813105"], "Bankura": ["713150", "713425", "713512", "713513", "713514", "713517"], "Banswara": ["327001", "327021", "327022", "327023", "327024", "327025"], "Barabanki": ["224116", "224117", "224118", "224119", "224120", "225001"], "Baramulla": ["192125", "193101", "193103", "193108", "193121", "193122"], "Baran": ["325202", "325205", "325206", "325207", "325209", "325215"], "Bardhaman": ["712410", "713101", "713102", "713103", "713104", "713121"], "Bareilly": ["243001", "243002", "243003", "243004", "243005", "243006"], "Bargarh": ["768027", "768028", "768029", "768030", "768031", "768032"], "Barmer": ["344001", "344011", "344012", "344021", "344022", "344024"], "Barnala": ["148024", "148025", "148100", "148101", "148102", "148103"], "Barpeta": ["781301", "781302", "781305", "781307", "781308", "781309"], "Barwani": ["451442", "451447", "451449", "451551", "451556", "451660"], "Bastar": ["494001", "494010", "494111", "494221", "494222", "494223"], "Basti": ["271305", "272001", "272002", "272123", "272124", "272125"], "Bathinda": ["151001", "151002", "151003", "151004", "151005", "151101"], "Beed": ["413207", "413229", "413249", "414202", "414203", "414204"], "Begusarai": ["811106", "848201", "848202", "848204", "851101", "851111"], "Belgaum": ["590001", "590003", "590005", "590006", "590008", "590009"], "Bellary": ["583101", "583102", "583103", "583104", "583105", "583111"], "Betul": ["460001", "460004", "460110", "460220", "460225", "460330"], "Bhadrak": ["756100", "756101", "756111", "756112", "756113", "756114"], "Bhagalpur": ["811211", "812001", "812002", "812003", "812004", "812005"], "Bhandara": ["441802", "441803", "441804", "441805", "441809", "441903"], "Bharatpur": ["321001", "321021", "321022", "321023", "321024", "321025"], "Bharuch": ["391810", "392001", "392011", "392012", "392015", "392020"], "Bhavnagar": ["364001", "364002", "364003", "364004", "364005", "364006"], "Bhilwara": ["311001", "311011", "311021", "311022", "311023", "311024"], "Bhind": ["475110", "475115", "477001", "477105", "477111", "477116"], "Bhiwani": ["127021", "127022", "127025", "127026", "127027", "127028"], "Bhojpur": ["802152", "802154", "802155", "802156", "802157", "802158"], "Bhopal": ["462001", "462002", "462003", "462004", "462007", "462008"], "Bidar": ["585226", "585227", "585326", "585327", "585328", "585329"], "Bijapur(Cgh)": ["494444", "494446", "494447", "494448", "494450", "494552"], "Bijapur(Kar)": ["586101", "586102", "586103", "586104", "586108", "586109"], "Bijnor": ["246701", "246721", "246722", "246723", "246724", "246725"], "Bikaner": ["331801", "331803", "331811", "334001", "334003", "334004"], "Bilaspur (Hp)": ["174001", "174002", "174003", "174004", "174005", "174011"], "Bilaspur(Cgh)": ["495001", "495003", "495004", "495006", "495009", "495112"], "Birbhum": ["731101", "731102", "731103", "731104", "731121", "731123"], "Bishnupur": ["795118", "795124", "795126", "795133", "795134"], "Bokaro": ["814155", "825102", "825315", "825322", "827001", "827003"], "Bongaigaon": ["783134", "783360", "783373", "783380", "783381", "783382"], "Boudh": ["762012", "762013", "762014", "762015", "762016", "762017"], "Budaun": ["242021", "243601", "243630", "243631", "243632", "243633"], "Budgam": ["190005", "190007", "190014", "190015", "190019", "190021"], "Bulandshahr": ["203001", "203002", "203129", "203131", "203132", "203141"], "Buldhana": ["443001", "443002", "443101", "443102", "443103", "443104"], "Bundi": ["323001", "323021", "323022", "323023", "323024", "323025"], "Burhanpur": ["450001", "450221", "450331", "450332", "450445"], "Buxar": ["802101", "802102", "802103", "802111", "802112", "802113"], "Cachar": ["788001", "788002", "788003", "788004", "788005", "788006"], "Central Delhi": ["110001", "110002", "110003", "110004", "110005", "110006"], "Chamba": ["176207", "176301", "176302", "176303", "176304", "176305"], "Chamoli": ["246401", "246422", "246424", "246426", "246427", "246428"], "Champawat": ["262308", "262309", "262310", "262523", "262524", "262525"], "Champhai": ["796290", "796310", "796320", "796321", "796370"], "Chamrajnagar": ["571109", "571110", "571111", "571115", "571117", "571123"], "Chandauli": ["221009", "221115", "232101", "232102", "232103", "232104"], "Chandel": ["795101", "795102", "795127", "795131", "795135"], "Chandigarh": ["140119", "140125", "140133", "160001", "160002", "160003"], "Chandrapur": ["441205", "441206", "441207", "441212", "441215", "441221"], "Changlang": ["792055", "792056", "792103", "792120", "792121", "792122"], "Chatra": ["825103", "825321", "825401", "825403", "825404", "825408"], "Chennai": ["600001", "600002", "600003", "600004", "600005", "600006"], "Chhatarpur": ["471001", "471101", "471105", "471111", "471201", "471301"], "Chhindwara": ["460663", "480001", "480003", "480105", "480106", "480107"], "Chickmagalur": ["577101", "577102", "577111", "577112", "577113", "577114"], "Chikkaballapur": ["561206", "561207", "561208", "561209", "561210", "561211"], "Chitradurga": ["577501", "577502", "577511", "577515", "577517", "577518"], "Chitrakoot": ["210202", "210203", "210204", "210205", "210206", "210207"], "Chittoor": ["517001", "517002", "517004", "517101", "517102", "517112"], "Chittorgarh": ["312001", "312021", "312022", "312023", "312024", "312025"], "Churachandpur": ["795006", "795117", "795118", "795128", "795139", "795143"], "Churu": ["331001", "331021", "331022", "331023", "331029", "331031"], "Coimbatore": ["624622", "638103", "638402", "638458", "638462", "641001"], "Cooch Behar": ["735122", "735211", "735217", "735224", "735301", "735303"], "Cuddalore": ["605007", "605106", "606001", "606003", "606103", "606104"], "Cuddapah": ["515465", "516001", "516002", "516003", "516004", "516101"], "Cuttack": ["752120", "753001", "753002", "753003", "753004", "753006"], "Dadra & Nagar Haveli": ["396193", "396230", "396235", "396240"], "Dahod": ["389110", "389130", "389140", "389146", "389151", "389152"], "Dakshina Kannada": ["574107", "574109", "574141", "574142", "574143", "574144"], "Daman": ["396210", "396215", "396220"], "Damoh": ["470661", "470663", "470664", "470666", "470672", "470673"], "Dantewada": ["494111", "494114", "494115", "494122", "494441", "494448"], "Darbhanga": ["846001", "846002", "846003", "846004", "846005", "846006"], "Darjiling": ["734001", "734002", "734003", "734004", "734005", "734007"], "Darrang": ["784113", "784114", "784115", "784116", "784125", "784144"], "Datia": ["475335", "475336", "475661", "475671", "475673", "475675"], "Dausa": ["303004", "303302", "303303", "303304", "303305", "303313"], "Davangere": ["577001", "577002", "577003", "577004", "577005", "577006"], "Debagarh": ["768107", "768108", "768109", "768110", "768119", "768121"], "Dehradun": ["247670", "248001", "248002", "248003", "248005", "248006"], "Deoghar": ["814112", "814113", "814114", "814115", "814116", "814120"], "Deoria": ["274001", "274182", "274201", "274202", "274204", "274205"], "Dewas": ["455001", "455111", "455115", "455116", "455118", "455221"], "Dhalai": ["799204", "799266", "799273", "799275", "799278", "799284"], "Dhamtari": ["493662", "493663", "493770", "493773", "493776", "493778"], "Dhanbad": ["826001", "826003", "826004", "826005", "827302", "828101"], "Dhar": ["454001", "454010", "454111", "454116", "454221", "454331"], "Dharmapuri": ["635111", "635201", "635202", "635205", "635301", "635302"], "Dharwad": ["580001", "580002", "580003", "580004", "580005", "580006"], "Dhemaji": ["787026", "787034", "787035", "787053", "787057", "787058"], "Dhenkanal": ["759001", "759013", "759014", "759015", "759016", "759017"], "Dholpur": ["328001", "328021", "328022", "328023", "328024", "328025"], "Dhubri": ["783127", "783128", "783131", "783135", "783301", "783323"], "Dhule": ["424001", "424002", "424004", "424005", "424006", "424301"], "Dibang Valley": ["792101", "792110"], "Dibrugarh": ["785675", "785676", "786001", "786002", "786003", "786004"], "Dimapur": ["797103", "797106", "797112", "797113", "797115", "797116"], "Dindigul": ["624001", "624002", "624003", "624004", "624005", "624101"], "Dindori": ["481672", "481778", "481879", "481880", "481882", "481884"], "Diu": ["362520", "362540", "362570"], "Doda": ["182124", "182143", "182144", "182145", "182146", "182147"], "Dumka": ["814101", "814102", "814103", "814110", "814118", "814119"], "Dungarpur": ["314001", "314011", "314021", "314022", "314023", "314024"], "Durg": ["490001", "490006", "490009", "490011", "490020", "490021"], "East Champaran": ["845301", "845302", "845303", "845304", "845305", "845315"], "East Delhi": ["110031", "110032", "110051", "110053", "110090", "110091"], "East Garo Hills": ["783123", "783134", "793106", "794002", "794107", "794108"], "East Godavari": ["507130", "533001", "533002", "533003", "533004", "533005"], "East Kameng": ["790102", "790103"], "East Khasi Hills": ["793001", "793002", "793003", "793004", "793005", "793006"], "East Midnapore": ["721130", "721131", "721134", "721135", "721137", "721139"], "East Nimar": ["450001", "450051", "450110", "450112", "450114", "450116"], "East Siang": ["791102", "791103", "791104"], "East Sikkim": ["737101", "737102", "737103", "737106", "737107", "737113"], "East Singhbhum": ["831001", "831002", "831003", "831004", "831005", "831006"], "Ernakulam": ["680667", "682001", "682002", "682003", "682004", "682005"], "Erode": ["638001", "638002", "638003", "638004", "638005", "638007"], "Etah": ["207001", "207002", "207003", "207120", "207121", "207122"], "Etawah": ["206001", "206002", "206003", "206120", "206123", "206124"], "Faizabad": ["224001", "224116", "224117", "224118", "224119", "224120"], "Faridabad": ["121001", "121002", "121003", "121004", "121005", "121006"], "Faridkot": ["151202", "151203", "151204", "151205", "151207", "151208"], "Farrukhabad": ["209501", "209502", "209503", "209504", "209505", "209601"], "Fatehabad": ["125048", "125050", "125051", "125052", "125053", "125106"], "Fatehgarh Sahib": ["140405", "140406", "140407", "140412", "140417", "140602"], "Fatehpur": ["212601", "212620", "212621", "212622", "212631", "212635"], "Fazilka": ["152118"], "Firozabad": ["205262", "207301", "283103", "283130", "283135", "283136"], "Firozpur": ["142041", "142043", "142044", "142047", "142050", "142052"], "Gadag": ["582101", "582102", "582103", "582111", "582112", "582113"], "Gadchiroli": ["441207", "441208", "441209", "441217", "442504", "442603"], "Gajapati": ["761014", "761015", "761016", "761017", "761018", "761200"], "Gandhi Nagar": ["380060", "382006", "382007", "382010", "382016", "382021"], "Ganganagar": ["335001", "335002", "335021", "335022", "335023", "335024"], "Ganjam": ["760001", "760002", "760003", "760004", "760005", "760006"], "Garhwa": ["822112", "822114", "822118", "822124", "822125", "822128"], "Gariaband": ["493887", "493891"], "Gautam Buddha Nagar": ["201008", "201301", "201303", "201304", "201305", "201306"], "Gaya": ["804401", "804402", "804403", "804404", "804419", "804421"], "Ghaziabad": ["201001", "201002", "201003", "201004", "201005", "201006"], "Ghazipur": ["232325", "232326", "232327", "232328", "232329", "232330"], "Giridh": ["815301", "815302", "815311", "815312", "815313", "815314"], "Goalpara": ["783101", "783120", "783121", "783122", "783123", "783124"], "Godda": ["813208", "814102", "814133", "814147", "814151", "814153"], "Golaghat": ["785104", "785601", "785602", "785603", "785609", "785610"], "Gonda": ["271001", "271002", "271003", "271122", "271123", "271124"], "Gondia": ["441601", "441614", "441701", "441702", "441801", "441806"], "Gopalganj": ["841405", "841407", "841409", "841413", "841420", "841423"], "Gorakhpur": ["273001", "273002", "273003", "273004", "273005", "273006"], "Gulbarga": ["585101", "585102", "585103", "585104", "585105", "585106"], "Gumla": ["835203", "835206", "835207", "835208", "835211", "835212"], "Guna": ["473001", "473101", "473105", "473110", "473111", "473112"], "Guntur": ["522001", "522002", "522003", "522004", "522005", "522006"], "Gurdaspur": ["143505", "143506", "143507", "143511", "143512", "143513"], "Gurgaon": ["122001", "122002", "122003", "122004", "122005", "122006"], "Gwalior": ["474001", "474002", "474003", "474004", "474005", "474006"], "Hailakandi": ["788117", "788150", "788152", "788155", "788156", "788160"], "Hamirpur": ["210301", "210341", "210421", "210422", "210424", "210425"], "Hamirpur(Hp)": ["174304", "174305", "174309", "174311", "174312", "174405"], "Hanumangarh": ["335024", "335063", "335064", "335501", "335502", "335503"], "Harda": ["461111", "461228", "461331", "461441"], "Hardoi": ["241001", "241121", "241122", "241123", "241124", "241125"], "Haridwar": ["246763", "247656", "247661", "247662", "247663", "247664"], "Hassan": ["573101", "573102", "573103", "573111", "573112", "573113"], "Hathras": ["202002", "202121", "202139", "204101", "204102", "204211"], "Haveri": ["581101", "581102", "581104", "581106", "581108", "581109"], "Hazaribag": ["825301", "825302", "825303", "825311", "825312", "825313"], "Hingoli": ["431501", "431509", "431512", "431513", "431542", "431701"], "Hisar": ["125001", "125004", "125005", "125006", "125007", "125011"], "Hooghly": ["712101", "712102", "712103", "712104", "712105", "712121"], "Hoshangabad": ["461001", "461005", "461110", "461111", "461114", "461115"], "Hoshiarpur": ["144105", "144202", "144204", "144205", "144206", "144207"], "Howrah": ["711101", "711102", "711103", "711104", "711105", "711106"], "Hyderabad": ["500001", "500002", "500003", "500004", "500005", "500006"], "Idukki": ["685501", "685503", "685505", "685507", "685508", "685509"], "Imphal East": ["795003", "795005", "795008", "795010", "795114", "795116"], "Imphal West": ["795001", "795002", "795003", "795004", "795009", "795113"], "Indore": ["452001", "452002", "452003", "452005", "452006", "452007"], "Jabalpur": ["482001", "482002", "482003", "482004", "482005", "482008"], "Jagatsinghapur": ["754102", "754103", "754104", "754106", "754107", "754108"], "Jaintia Hills": ["793109", "793150", "793151", "793160", "793200", "793210"], "Jaipur": ["302001", "302002", "302003", "302004", "302005", "302006"], "Jaisalmer": ["342302", "342310", "345001", "345021", "345022", "345023"], "Jajapur": ["754023", "754024", "754082", "754205", "754214", "754292"], "Jalandhar": ["144001", "144002", "144003", "144004", "144005", "144006"], "Jalaun": ["284203", "284204", "284303", "284306", "285001", "285121"], "Jalgaon": ["424101", "424102", "424103", "424104", "424105", "424106"], "Jalna": ["431114", "431121", "431132", "431134", "431137", "431202"], "Jalor": ["307029", "307030", "307515", "307803", "343001", "343002"], "Jalpaiguri": ["734001", "734004", "734006", "734007", "734015", "734501"], "Jammu": ["180001", "180002", "180003", "180004", "180005", "180006"], "Jamnagar": ["360110", "360480", "360490", "360510", "360515", "360520"], "Jamtara": ["814166", "815351", "815352", "815354", "815355", "815359"], "Jamui": ["811301", "811303", "811305", "811307", "811308", "811311"], "Janjgir-Champa": ["495552", "495553", "495554", "495556", "495557", "495559"], "Jashpur": ["496118", "496220", "496223", "496224", "496225", "496227"], "Jaunpur": ["222001", "222002", "222003", "222101", "222105", "222109"], "Jehanabad": ["804405", "804406", "804407", "804408", "804417", "804418"], "Jhabua": ["457550", "457661", "457770", "457772", "457773", "457775"], "Jhajjar": ["124102", "124103", "124104", "124105", "124106", "124107"], "Jhalawar": ["326001", "326021", "326022", "326023", "326033", "326034"], "Jhansi": ["284001", "284002", "284003", "284120", "284121", "284123"], "Jharsuguda": ["768201", "768202", "768203", "768204", "768211", "768213"], "Jhujhunu": ["331025", "331026", "331027", "331028", "331030", "332716"], "Jind": ["126101", "126102", "126110", "126111", "126112", "126113"], "Jodhpur": ["342001", "342003", "342005", "342006", "342007", "342008"], "Jorhat": ["785001", "785004", "785006", "785007", "785008", "785009"], "Junagadh": ["360490", "362001", "362002", "362004", "362011", "362015"], "Jyotiba Phule Nagar": ["244102", "244221", "244222", "244223", "244225", "244231"], "K.V.Rangareddy": ["500005", "500008", "500018", "500019", "500030", "500032"], "Kachchh": ["370001", "370015", "370020", "370030", "370040", "370105"], "Kaimur (Bhabua)": ["802132", "802213", "802218", "821101", "821102", "821103"], "Kaithal": ["136020", "136021", "136026", "136027", "136033", "136034"], "Kalahandi": ["766001", "766002", "766011", "766012", "766013", "766014"], "Kamrup": ["781001", "781003", "781004", "781005", "781006", "781007"], "Kanchipuram": ["600016", "600041", "600043", "600044", "600045", "600046"], "Kandhamal": ["762001", "762002", "762010", "762011", "762012", "762018"], "Kangra": ["175013", "176001", "176021", "176022", "176023", "176025"], "Kanker": ["494333", "494334", "494335", "494336", "494337", "494635"], "Kannauj": ["209720", "209721", "209722", "209723", "209725", "209726"], "Kannur": ["670001", "670002", "670003", "670004", "670005", "670006"], "Kanpur Dehat": ["208001", "209101", "209111", "209112", "209115", "209121"], "Kanpur Nagar": ["208001", "208002", "208003", "208004", "208005", "208006"], "Kanyakumari": ["629001", "629002", "629003", "629004", "629101", "629102"], "Kapurthala": ["144032", "144033", "144401", "144402", "144403", "144408"], "Karaikal": ["609601", "609602", "609603", "609604", "609605", "609606"], "Karauli": ["321610", "321611", "322033", "322202", "322203", "322204"], "Karbi Anglong": ["782139", "782410", "782425", "782435", "782441", "782442"], "Kargil": ["194102", "194103", "194105", "194109", "194301", "194302"], "Karim Nagar": ["505001", "505101", "505102", "505122", "505129", "505152"], "Karimganj": ["788006", "788120", "788166", "788701", "788709", "788710"], "Karnal": ["132001", "132022", "132023", "132024", "132036", "132037"], "Karur": ["621301", "621306", "621311", "621313", "621315", "638151"], "Kasargod": ["670511", "671121", "671122", "671123", "671124", "671310"], "Kathua": ["184101", "184102", "184104", "184121", "184141", "184142"], "Katihar": ["853204", "854101", "854103", "854104", "854105", "854106"], "Katni": ["483220", "483222", "483225", "483330", "483331", "483332"], "Kaushambi": ["211011", "212201", "212202", "212203", "212204", "212205"], "Kawardha": ["491336", "491559", "491995"], "Kendrapara": ["754134", "754141", "754153", "754162", "754203", "754210"], "Kendujhar": ["755019", "756121", "758001", "758002", "758013", "758014"], "Khagaria": ["811202", "848203", "851201", "851202", "851203", "851204"], "Khammam": ["507001", "507002", "507003", "507101", "507103", "507111"], "Khandwa": ["450112", "450116", "450117", "450337", "450551", "450554"], "Khargone": ["451332", "451441", "451660"], "Kheda": ["387001", "387002", "387110", "387115", "387120", "387130"], "Kheri": ["261501", "261502", "261505", "261506", "262701", "262702"], "Khorda": ["751001", "751002", "751003", "751004", "751005", "751006"], "Khunti": ["835235", "835302"], "Kinnaur": ["172030", "172101", "172103", "172104", "172105", "172106"], "Kiphire": ["798611"], "Kishanganj": ["855101", "855106", "855107", "855108", "855115", "855116"], "Kodagu": ["571201", "571211", "571212", "571213", "571214", "571215"], "Koderma": ["825109", "825132", "825318", "825323", "825407", "825409"], "Kohima": ["797001", "797002", "797003", "797004", "797005", "797101"], "Kokrajhar": ["783127", "783332", "783333", "783336", "783337", "783345"], "Kolar": ["561206", "561207", "561208", "561209", "561210", "561211"], "Kolasib": ["796070", "796081", "796082", "796091", "796101"], "Kolhapur": ["415101", "416001", "416002", "416003", "416004", "416005"], "Kolkata": ["700001", "700002", "700003", "700004", "700005", "700006"], "Kollam": ["689695", "689696", "690518", "690519", "690520", "690521"], "Koppal": ["583226", "583227", "583228", "583229", "583230", "583231"], "Koraput": ["763001", "763002", "763003", "763004", "763008", "764001"], "Korba": ["495118", "495445", "495446", "495447", "495448", "495449"], "Koriya": ["497331", "497335", "497339", "497442", "497446", "497447"], "Kota": ["324001", "324002", "324003", "324004", "324005", "324006"], "Kottayam": ["686001", "686002", "686003", "686004", "686005", "686006"], "Kozhikode": ["673001", "673002", "673003", "673004", "673005", "673006"], "Krishna": ["520001", "520002", "520003", "520004", "520007", "520008"], "Krishnagiri": ["635001", "635002", "635101", "635102", "635103", "635104"], "Kulgam": ["192233"], "Kullu": ["172001", "172002", "172023", "172025", "172026", "172032"], "Kupwara": ["193221", "193222", "193223", "193224", "193225", "193302"], "Kurnool": ["518001", "518002", "518003", "518004", "518005", "518006"], "Kurukshetra": ["136030", "136038", "136118", "136119", "136128", "136129"], "Kurung Kumey": ["791118"], "Kushinagar": ["274149", "274203", "274206", "274207", "274301", "274302"], "Lahul & Spiti": ["172113", "172114", "172117", "175132", "175133", "175139"], "Lakhimpur": ["784160", "784161", "784163", "784164", "784165", "787001"], "Lakhisarai": ["811106", "811107", "811112", "811201", "811302", "811304"], "Lakshadweep": ["682551", "682552", "682553", "682554", "682555", "682556"], "Lalitpur": ["284122", "284123", "284124", "284125", "284126", "284136"], "Latehar": ["822111", "822112", "822119", "829202", "829203", "829204"], "Latur": ["413502", "413510", "413511", "413512", "413513", "413514"], "Lawngtlai": ["796770", "796891", "796901"], "Leh": ["194101", "194104", "194105", "194106", "194201", "194401"], "Lohardaga": ["834004", "834009", "835213", "835218", "835302", "835325"], "Lohit": ["792001", "792102", "792103", "792104", "792105", "792111"], "Longleng": ["798625"], "Lower Dibang Valley": ["792110"], "Lower Subansiri": ["791113", "791119", "791120"], "Lucknow": ["226001", "226002", "226003", "226004", "226005", "226006"], "Ludhiana": ["141001", "141002", "141003", "141004", "141006", "141007"], "Lunglei": ["796571", "796581", "796691", "796701", "796710", "796730"], "Madhepura": ["813204", "852101", "852108", "852112", "852113", "852114"], "Madhubani": ["847102", "847108", "847109", "847122", "847211", "847212"], "Madurai": ["625001", "625002", "625003", "625004", "625005", "625006"], "Mahabub Nagar": ["509001", "509002", "509102", "509103", "509104", "509105"], "Maharajganj": ["273151", "273155", "273157", "273161", "273162", "273163"], "Mahasamund": ["492112", "493445", "493448", "493449", "493551", "493554"], "Mahe": ["673310"], "Mahendragarh": ["123001", "123021", "123023", "123024", "123027", "123028"], "Mahesana": ["382115", "382120", "382165", "382170", "382705", "382710"], "Mahoba": ["210421", "210423", "210424", "210425", "210426", "210427"], "Mainpuri": ["205001", "205119", "205121", "205247", "205261", "205262"], "Malappuram": ["673314", "673632", "673633", "673634", "673635", "673636"], "Malda": ["732101", "732102", "732103", "732121", "732122", "732123"], "Malkangiri": ["764041", "764043", "764044", "764045", "764046", "764047"], "Mammit": ["796410", "796421", "796431", "796441", "796471", "796501"], "Mandi": ["175001", "175002", "175003", "175004", "175005", "175006"], "Mandla": ["481661", "481662", "481663", "481664", "481665", "481666"], "Mandsaur": ["458001", "458002", "458339", "458389", "458553", "458556"], "Mandya": ["571401", "571402", "571403", "571404", "571405", "571415"], "Mansa": ["151501", "151502", "151503", "151504", "151505", "151506"], "Marigaon": ["782001", "782103", "782104", "782105", "782106", "782121"], "Mathura": ["281001", "281003", "281004", "281005", "281006", "281121"], "Mau": ["221601", "221602", "221603", "221705", "221706", "223222"], "Mayurbhanj": ["757001", "757002", "757003", "757014", "757016", "757017"], "Medak": ["502001", "502032", "502101", "502102", "502103", "502107"], "Medinipur": ["721101"], "Meerut": ["245206", "250001", "250002", "250003", "250004", "250005"], "Mirzapur": ["231001", "231210", "231211", "231301", "231302", "231303"], "Moga": ["142001", "142002", "142003", "142011", "142037", "142038"], "Mohali": ["140103", "140109", "140110", "140112", "140201", "140301"], "Mokokchung": ["798601", "798604", "798613", "798614", "798615", "798617"], "Mon": ["798602", "798603", "798621", "798622"], "Moradabad": ["244001", "244102", "244103", "244104", "244301", "244302"], "Morena": ["476001", "476111", "476115", "476134", "476219", "476221"], "Muktsar": ["151202", "151210", "151211", "152022", "152025", "152026"], "Mumbai": ["400001", "400002", "400003", "400004", "400005", "400006"], "Munger": ["811201", "811202", "811211", "811212", "811213", "811214"], "Murshidabad": ["713123", "742101", "742102", "742103", "742104", "742113"], "Muzaffarnagar": ["247771", "247772", "247773", "247774", "247775", "247776"], "Muzaffarpur": ["842001", "842002", "842003", "842004", "842005", "843101"], "Mysore": ["570001", "570002", "570003", "570004", "570005", "570006"], "Nabarangapur": ["764049", "764055", "764059", "764061", "764063", "764070"], "Nadia": ["741101", "741102", "741103", "741121", "741122", "741123"], "Nagaon": ["782001", "782002", "782003", "782101", "782102", "782103"], "Nagapattinam": ["609001", "609003", "609101", "609102", "609103", "609104"], "Nagaur": ["305026", "341001", "341021", "341022", "341023", "341024"], "Nagpur": ["440001", "440002", "440003", "440005", "440006", "440007"], "Nainital": ["244713", "244715", "262402", "262580", "263001", "263126"], "Nalanda": ["801301", "801302", "801303", "801304", "801305", "801306"], "Nalbari": ["781126", "781138", "781303", "781304", "781305", "781306"], "Nalgonda": ["508001", "508002", "508004", "508101", "508105", "508111"], "Namakkal": ["621215", "636113", "636118", "636142", "636202", "636301"], "Nanded": ["431601", "431602", "431603", "431604", "431605", "431606"], "Nandurbar": ["425409", "425410", "425411", "425412", "425413", "425414"], "Narayanpur": ["494661"], "Narmada": ["391110", "391120", "391121", "391810", "392020", "392025"], "Narsinghpur": ["461771", "461990", "487001", "487110", "487114", "487118"], "Nashik": ["422001", "422002", "422003", "422004", "422005", "422006"], "Navsari": ["394730", "396040", "396051", "396060", "396110", "396310"], "Nawada": ["803109", "805101", "805102", "805103", "805104", "805106"], "Nawanshahr": ["144029", "144224", "144415", "144417", "144421", "144422"], "Nayagarh": ["752024", "752025", "752026", "752063", "752065", "752068"], "Neemuch": ["458110", "458113", "458116", "458118", "458220", "458226"], "Nellore": ["524001", "524002", "524003", "524004", "524005", "524101"], "New Delhi": ["110001", "110020", "110029"], "Nicobar": ["744301", "744302", "744303", "744304"], "Nilgiris": ["643001", "643002", "643003", "643004", "643005", "643006"], "Nizamabad": ["503001", "503002", "503003", "503101", "503102", "503108"], "North 24 Parganas": ["700048", "700049", "700051", "700055", "700056", "700057"], "North And Middle Andaman": ["744201", "744202", "744203", "744204", "744205", "744209"], "North Cachar Hills": ["788108", "788113", "788818", "788819", "788820", "788830"], "North Delhi": ["110006", "110007", "110054", "110084", "110085"], "North Dinajpur": ["733121", "733123", "733125", "733128", "733129", "733130"], "North East Delhi": ["110032", "110053", "110090", "110093", "110094"], "North Goa": ["403001", "403002", "403004", "403005", "403006", "403101"], "North Sikkim": ["737116", "737117", "737118", "737119", "737120"], "North Tripura": ["799250", "799251", "799253", "799254", "799256", "799260"], "North West Delhi": ["110009", "110033", "110034", "110035", "110036", "110039"], "Nuapada": ["766104", "766105", "766106", "766107", "766108", "766111"], "Osmanabad": ["413405", "413501", "413502", "413503", "413504", "413505"], "Pakur": ["814111", "814133", "816102", "816103", "816104", "816105"], "Palakkad": ["678001", "678002", "678003", "678004", "678005", "678006"], "Palamau": ["822101", "822102", "822110", "822113", "822115", "822116"], "Pali": ["306001", "306021", "306022", "306023", "306101", "306102"], "Panch Mahals": ["388270", "388710", "388713", "389001", "389002", "389110"], "Panchkula": ["133301", "133302", "134101", "134102", "134103", "134104"], "Panipat": ["132101", "132102", "132103", "132104", "132105", "132106"], "Panna": ["488001", "488050", "488051", "488059", "488220", "488222"], "Papum Pare": ["790001", "791001", "791109", "791110", "791111", "791112"], "Parbhani": ["431401", "431402", "431503", "431505", "431506", "431508"], "Patan": ["384110", "384151", "384220", "384221", "384225", "384229"], "Pathanamthitta": ["685533", "686510", "686511", "686547", "689101", "689102"], "Pathankot": ["143533", "143534", "145022", "145023", "145024", "145025"], "Patiala": ["140201", "140401", "140402", "140406", "140412", "140417"], "Patna": ["800001", "800002", "800003", "800004", "800005", "800006"], "Pauri Garhwal": ["246001", "246113", "246121", "246123", "246124", "246125"], "Perambalur": ["621101", "621102", "621103", "621104", "621106", "621107"], "Peren": ["797101", "797110"], "Phek": ["797102", "797104", "797107", "797108", "797114"], "Pilibhit": ["243003", "243123", "262001", "262121", "262122", "262124"], "Pithoragarh": ["262501", "262502", "262520", "262521", "262522", "262523"], "Pondicherry": ["533464", "605001", "605002", "605003", "605004", "605005"], "Poonch": ["184121", "185101", "185102", "185121", "185211"], "Porbandar": ["360490", "360510", "360530", "360545", "360550", "360560"], "Prakasam": ["523001", "523002", "523101", "523104", "523105", "523108"], "Pratapgarh": ["222301", "229408", "229410", "230001", "230002", "230121"], "Pudukkottai": ["613301", "614616", "614617", "614618", "614619", "614620"], "Pulwama": ["191112", "192121", "192122", "192123", "192124", "192211"], "Pune": ["410301", "410302", "410401", "410402", "410403", "410405"], "Puri": ["752001", "752002", "752003", "752011", "752012", "752013"], "Purnia": ["853204", "854102", "854113", "854201", "854202", "854203"], "Puruliya": ["723101", "723102", "723103", "723104", "723121", "723126"], "Raebareli": ["229001", "229010", "229103", "229120", "229121", "229122"], "Raichur": ["584101", "584102", "584103", "584104", "584111", "584113"], "Raigarh": ["496001", "496005", "496100", "496107", "496108", "496109"], "Raigarh(Mh)": ["400001", "400702", "400704", "400707", "402101", "402102"], "Raipur": ["492001", "492002", "492003", "492004", "492005", "492008"], "Raisen": ["464551", "464651", "464661", "464665", "464668", "464671"], "Rajauri": ["185121", "185131", "185132", "185135", "185151", "185152"], "Rajgarh": ["465661", "465667", "465669", "465674", "465677", "465679"], "Rajkot": ["360001", "360002", "360003", "360004", "360005", "360006"], "Rajnandgaon": ["491229", "491441", "491444", "491445", "491557", "491558"], "Rajsamand": ["305921", "305922", "305925", "313202", "313206", "313207"], "Ramanagar": ["561201", "562108", "562109", "562112", "562117", "562119"], "Ramanathapuram": ["623115", "623120", "623135", "623308", "623315", "623401"], "Ramgarh": ["825101", "825314", "825316", "825325", "825326", "825330"], "Rampur": ["244701", "244901", "244921", "244922", "244923", "244924"], "Ranchi": ["829205", "829208", "829209", "829210", "834001", "834002"], "Ratlam": ["457001", "457114", "457118", "457119", "457222", "457226"], "Ratnagiri": ["415202", "415203", "415208", "415214", "415601", "415602"], "Rayagada": ["764059", "764062", "765001", "765002", "765013", "765015"], "Reasi": ["182311", "182320", "185233"], "Rewa": ["485331", "486001", "486002", "486003", "486005", "486006"], "Rewari": ["123021", "123034", "123035", "123101", "123102", "123103"], "Ri Bhoi": ["793101", "793102", "793103", "793104", "793105", "793122"], "Rohtak": ["124001", "124010", "124021", "124022", "124111", "124112"], "Rohtas": ["802211", "802212", "802213", "802214", "802215", "802216"], "Ropar": ["140101", "140102", "140108", "140115", "140117", "140118"], "Rudraprayag": ["246141", "246171", "246419", "246421", "246425", "246429"], "Rupnagar": ["140001", "140101", "140102", "140103", "140108", "140111"], "Sabarkantha": ["383001", "383006", "383010", "383030", "383110", "383120"], "Sagar": ["464240", "470001", "470002", "470003", "470004", "470021"], "Saharanpur": ["247001", "247002", "247120", "247121", "247122", "247129"], "Saharsa": ["852106", "852107", "852116", "852121", "852123", "852124"], "Sahibganj": ["813208", "816101", "816102", "816105", "816108", "816109"], "Saiha": ["796901"], "Salem": ["621110", "636001", "636002", "636003", "636004", "636005"], "Samastipur": ["844501", "844506", "847105", "847301", "848101", "848102"], "Sambalpur": ["768001", "768002", "768003", "768004", "768005", "768006"], "Sangli": ["415301", "415302", "415303", "415304", "415305", "415306"], "Sangrur": ["148001", "148002", "148017", "148018", "148019", "148020"], "Sant Kabir Nagar": ["272125", "272126", "272148", "272152", "272154", "272162"], "Sant Ravidas Nagar": ["221301", "221303", "221304", "221306", "221308", "221309"], "Saran": ["841101", "841201", "841202", "841204", "841205", "841206"], "Satara": ["412206", "412801", "412802", "412803", "412804", "412805"], "Satna": ["485001", "485005", "485111", "485112", "485113", "485114"], "Sawai Madhopur": ["322001", "322021", "322023", "322024", "322025", "322026"], "Sehore": ["465693", "466001", "466111", "466113", "466114", "466115"], "Senapati": ["795007", "795015", "795104", "795106", "795107", "795112"], "Seoni": ["480661", "480667", "480771", "480880", "480881", "480882"], "Seraikela-Kharsawan": ["831002", "831012", "831014", "832107", "832108", "832109"], "Serchhip": ["796181", "796184", "796186", "796370"], "Shahdol": ["484001", "484110", "484114", "484117", "484120", "484330"], "Shahjahanpur": ["242001", "242042", "242123", "242127", "242220", "242221"], "Shajapur": ["465001", "465106", "465110", "465113", "465116", "465118"], "Sheikhpura": ["811101", "811102", "811103", "811105", "811107", "811201"], "Sheohar": ["843313", "843316", "843317", "843325", "843329", "843334"], "Sheopur": ["476332", "476335", "476337", "476339", "476355", "477332"], "Shimla": ["171001", "171002", "171003", "171004", "171005", "171006"], "Shimoga": ["577115", "577201", "577202", "577203", "577204", "577205"], "Shivpuri": ["473551", "473585", "473638", "473660", "473662", "473665"], "Shopian": ["192305"], "Shrawasti": ["271201", "271802", "271803", "271804", "271805", "271821"], "Sibsagar": ["785006", "785107", "785602", "785614", "785616", "785633"], "Siddharthnagar": ["272129", "272148", "272151", "272152", "272153", "272154"], "Sidhi": ["484771", "486661", "486666", "486669", "486670", "486675"], "Sikar": ["331024", "332001", "332002", "332021", "332023", "332024"], "Simdega": ["835201", "835211", "835212", "835223", "835226", "835228"], "Sindhudurg": ["416510", "416511", "416512", "416513", "416514", "416515"], "Singrauli": ["486885", "486886", "486887", "486888", "486889"], "Sirmaur": ["171226", "173001", "173021", "173022", "173023", "173024"], "Sirohi": ["307001", "307019", "307022", "307023", "307024", "307026"], "Sirsa": ["125054", "125055", "125056", "125058", "125060", "125075"], "Sitamarhi": ["843301", "843302", "843311", "843313", "843314", "843315"], "Sitapur": ["261001", "261121", "261125", "261131", "261135", "261136"], "Sivaganga": ["621308", "623401", "623402", "623701", "630001", "630002"], "Siwan": ["841203", "841210", "841223", "841226", "841227", "841231"], "Solan": ["171102", "173201", "173202", "173204", "173205", "173206"], "Solapur": ["413001", "413002", "413003", "413004", "413005", "413006"], "Sonapur": ["767016", "767017", "767018", "767019", "767020", "767023"], "Sonbhadra": ["231205", "231206", "231207", "231208", "231209", "231210"], "Sonipat": ["131001", "131021", "131022", "131023", "131024", "131027"], "Sonitpur": ["784001", "784010", "784025", "784026", "784027", "784028"], "South 24 Parganas": ["700039", "700070", "700084", "700093", "700096", "700100"], "South Andaman": ["744101", "744102", "744103", "744104", "744105", "744106"], "South Delhi": ["110003", "110013", "110014", "110017", "110019", "110020"], "South Dinajpur": ["733101", "733102", "733103", "733121", "733124", "733125"], "South Garo Hills": ["794102", "794114"], "South Goa": ["403105", "403106", "403107", "403115", "403401", "403404"], "South Sikkim": ["737102", "737121", "737126", "737128", "737139"], "South Tripura": ["799013", "799015", "799045", "799101", "799103", "799104"], "South West Delhi": ["110010", "110016", "110021", "110022", "110023", "110028"], "Srikakulam": ["532001", "532005", "532122", "532123", "532127", "532148"], "Srinagar": ["190001", "190002", "190003", "190004", "190006", "190008"], "Sultanpur": ["222301", "222302", "222303", "227304", "227405", "227406"], "Sundergarh": ["769001", "769002", "769003", "769004", "769005", "769006"], "Supaul": ["847408", "847451", "847452", "852105", "852108", "852109"], "Surat": ["392150", "393050", "393125", "393130", "394101", "394105"], "Surendra Nagar": ["360055", "363001", "363002", "363020", "363030", "363035"], "Surguja": ["497001", "497101", "497111", "497114", "497116", "497117"], "Tamenglong": ["795125", "795141", "795147", "795159"], "Tapi": ["394365", "394375", "394380", "394630", "394633", "394635"], "Tarn Taran": ["143009", "143022", "143107", "143111", "143112", "143113"], "Tawang": ["790104", "790105", "790106"], "Tehri Garhwal": ["249001", "249121", "249122", "249123", "249124", "249125"], "Thane": ["400601", "400602", "400603", "400604", "400605", "400606"], "Thanjavur": ["609802", "609804", "609807", "612001", "612002", "612101"], "The Dangs": ["394710", "394715", "394716", "394720", "394730"], "Theni": ["625203", "625512", "625513", "625515", "625516", "625517"], "Thiruvananthapuram": ["695001", "695002", "695003", "695004", "695005", "695006"], "Thoubal": ["795101", "795103", "795130", "795132", "795135", "795138"], "Thrissur": ["679105", "679106", "679531", "679532", "679561", "679562"], "Tikamgarh": ["472001", "472005", "472010", "472101", "472111", "472115"], "Tinsukia": ["786125", "786126", "786145", "786146", "786147", "786150"], "Tirap": ["792129", "792130", "792131"], "Tiruchirappalli": ["607201", "620001", "620002", "620003", "620004", "620005"], "Tirunelveli": ["627001", "627002", "627003", "627004", "627005", "627006"], "Tiruvallur": ["600019", "600037", "600049", "600050", "600051", "600052"], "Tiruvannamalai": ["604401", "604402", "604403", "604404", "604405", "604406"], "Tiruvarur": ["609403", "609405", "609501", "609502", "609503", "609504"], "Tonk": ["304001", "304021", "304022", "304023", "304024", "304025"], "Tuensang": ["798612", "798616", "798626"], "Tumkur": ["561202", "572101", "572102", "572103", "572104", "572105"], "Tuticorin": ["626202", "626205", "627758", "627855", "628001", "628002"], "Udaipur": ["307025", "313001", "313002", "313003", "313004", "313011"], "Udham Singh Nagar": ["244712", "244713", "244715", "244716", "244717", "262308"], "Udhampur": ["182101", "182104", "182121", "182122", "182124", "182125"], "Udupi": ["574101", "574102", "574103", "574104", "574105", "574106"], "Ujjain": ["456001", "456003", "456006", "456010", "456221", "456222"], "Ukhrul": ["795142", "795144", "795145"], "Umaria": ["484001", "484551", "484552", "484555", "484660", "484661"], "Una": ["174301", "174302", "174303", "174306", "174307", "174308"], "Unnao": ["209101", "209801", "209821", "209825", "209827", "209831"], "Upper Siang": ["791002", "791102", "791105"], "Upper Subansiri": ["791122"], "Uttara Kannada": ["581121", "581129", "581186", "581187", "581301", "581302"], "Uttarkashi": ["249128", "249131", "249135", "249141", "249151", "249152"], "Vadodara": ["388710", "390001", "390002", "390003", "390004", "390006"], "Vaishali": ["843102", "843110", "843114", "844101", "844102", "844103"], "Valsad": ["396001", "396002", "396007", "396020", "396030", "396035"], "Varanasi": ["221001", "221002", "221003", "221004", "221005", "221006"], "Vellore": ["604407", "604505", "631001", "631002", "631003", "631004"], "Vidisha": ["464001", "464111", "464113", "464114", "464220", "464221"], "Villupuram": ["604001", "604101", "604102", "604151", "604152", "604153"], "Virudhunagar": ["626001", "626002", "626003", "626004", "626005", "626101"], "Visakhapatnam": ["530001", "530002", "530003", "530004", "530005", "530007"], "Vizianagaram": ["531162", "532122", "532127", "532407", "535001", "535002"], "Warangal": ["506001", "506002", "506003", "506004", "506005", "506006"], "Wardha": ["442001", "442003", "442101", "442102", "442104", "442105"], "Washim": ["444004", "444105", "444106", "444110", "444403", "444404"], "Wayanad": ["670644", "670645", "670646", "670721", "670731", "673121"], "West Champaran": ["845101", "845102", "845103", "845104", "845105", "845106"], "West Delhi": ["110015", "110018", "110026", "110027", "110041", "110058"], "West Garo Hills": ["794001", "794002", "794005", "794101", "794103", "794104"], "West Godavari": ["534001", "534002", "534003", "534004", "534005", "534006"], "West Kameng": ["790001", "790002", "790003", "790101", "790102", "790114"], "West Khasi Hills": ["793106", "793114", "793115", "793119", "793120", "793121"], "West Midnapore": ["721101", "721102", "721121", "721122", "721124", "721125"], "West Nimar": ["450551", "450554", "450771", "451001", "451111", "451113"], "West Siang": ["791001", "791003", "791101", "791125"], "West Sikkim": ["737111", "737113", "737121"], "West Singhbhum": ["831013", "833101", "833102", "833103", "833104", "833105"], "West Tripura": ["799001", "799002", "799003", "799004", "799005", "799006"], "Wokha": ["797099", "797100", "797111"], "Yadgir": ["585201", "585202", "585214", "585215", "585216", "585219"], "Yamuna Nagar": ["133103", "133204", "133206", "135001", "135002", "135003"], "Yavatmal": ["445001", "445002", "445101", "445102", "445103", "445105"], "Zunhebotto": ["797109", "798619", "798620", "798627"]};

var DISTRICT_PINCODE_PREFIXES = {
  'central delhi': '1100', 'east delhi': '1100', 'new delhi': '1100', 'north delhi': '1100',
  'north east delhi': '1100', 'north west delhi': '1100', 'south delhi': '1100', 'south west delhi': '1100', 'west delhi': '1100',
  'gurgaon': '1220', 'faridabad': '1210', 'rohtak': '1240', 'panipat': '1321', 'ambala': '1330', 'karnal': '1320', 'hisar': '1250',
  'amritsar': '1430', 'ludhiana': '1410', 'jalandhar': '1440', 'patiala': '1470', 'bathinda': '1510', 'mohali': '1600',
  'agra': '2820', 'aligarh': '2020', 'allahabad': '2110', 'prayagraj': '2110', 'bareilly': '2430', 'ghaziabad': '2010',
  'gautam buddha nagar': '2013', 'noida': '2013', 'gorakhpur': '2730', 'jhansi': '2840', 'kanpur nagar': '2080',
  'lucknow': '2260', 'mathura': '2810', 'meerut': '2500', 'moradabad': '2440', 'muzaffarnagar': '2510',
  'varanasi': '2210', 'ayodhya': '2240', 'faizabad': '2240', 'basti': '2720', 'azamgarh': '2760',
  'dehradun': '2480', 'haridwar': '2494', 'nainital': '2630', 'almora': '2636',
  'jaipur': '3020', 'jodhpur': '3420', 'kota': '3240', 'bikaner': '3340', 'ajmer': '3050', 'udaipur': '3130',
  'alwar': '3010', 'bhilwara': '3110', 'sikar': '3320', 'bharatpur': '3210',
  'ahmedabad': '3800', 'surat': '3950', 'vadodara': '3900', 'rajkot': '3600', 'bhavnagar': '3640', 'jamnagar': '3610',
  'gandhinagar': '3820', 'junagadh': '3620', 'anand': '3880', 'navsari': '3964', 'valsad': '3960',
  'mumbai city': '4000', 'mumbai suburban': '4000', 'mumbai': '4000', 'pune': '4110', 'nagpur': '4400',
  'thane': '4006', 'nashik': '4220', 'aurangabad': '4310', 'solapur': '4130', 'kolhapur': '4160', 'amravati': '4446',
  'indore': '4520', 'bhopal': '4620', 'jabalpur': '4820', 'gwalior': '4740', 'ujjain': '4560', 'raipur': '4920', 'bilaspur': '4950',
  'hyderabad': '5000', 'secunderabad': '5000', 'visakhapatnam': '5300', 'vijayawada': '5200', 'guntur': '5220',
  'warangal': '5060', 'karimnagar': '5050', 'kurnool': '5180', 'tirupati': '5175',
  'bengaluru urban': '5600', 'bengaluru rural': '5621', 'bangalore': '5600', 'mysore': '5700', 'mysuru': '5700',
  'mangalore': '5750', 'dakshina kannada': '5750', 'hubli': '5800', 'dharwad': '5800', 'belgaum': '5900', 'gulbarga': '5851',
  'chennai': '6000', 'coimbatore': '6410', 'madurai': '6250', 'tiruchirappalli': '6200', 'salem': '6360', 'tirunelveli': '6270',
  'vellore': '6320', 'thiruvananthapuram': '6950', 'kochi': '6820', 'ernakulam': '6820', 'kozhikode': '6730', 'thrissur': '6800',
  'kolkata': '7000', 'howrah': '7111', 'hooghly': '7121', 'darjeeling': '7341', 'bhubaneswar': '7510', 'cuttack': '7530',
  'guwahati': '7810', 'kamrup': '7810', 'patna': '8000', 'gaya': '8230', 'muzaffarpur': '8420', 'ranchi': '8340', 'dhanbad': '8260', 'jamshedpur': '8310'
};

function getDistrictPincodesFallback(district) {
  if (!district) return ['110001', '400001', '560001', '380001'];
  var clean = String(district).trim().toLowerCase();
  var prefix = DISTRICT_PINCODE_PREFIXES[clean];
  if (!prefix) {
    var hash = 0;
    for (var i = 0; i < clean.length; i++) {
      hash = ((hash << 5) - hash) + clean.charCodeAt(i);
      hash |= 0;
    }
    var baseNum = 1100 + (Math.abs(hash) % 7400);
    prefix = String(baseNum);
  }
  var pins = [];
  for (var p = 1; p <= 12; p++) {
    var pad = p < 10 ? '0' + p : '' + p;
    pins.push(prefix + pad);
  }
  return pins;
}

function getPincodes(district) {
  if (!district) return [];
  var sh = getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  var data = sh.getDataRange().getValues();
  var out = [];
  var targetDistrict = String(district || '').trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var d = String(data[i][0] || '').trim().toLowerCase();
    if (d === targetDistrict && data[i][1]) {
      out.push(String(data[i][1]).trim());
    }
  }
  out = uniqueSorted(out);
  if (!out.length && typeof ALL_INDIA_PINCODES !== 'undefined' && ALL_INDIA_PINCODES) {
    var clean = String(district).trim();
    if (ALL_INDIA_PINCODES[clean] && ALL_INDIA_PINCODES[clean].length) {
      return ALL_INDIA_PINCODES[clean];
    }
    for (var k in ALL_INDIA_PINCODES) {
      if (k.toLowerCase() === targetDistrict) {
        return ALL_INDIA_PINCODES[k];
      }
    }
    var normTarget = targetDistrict.replace(/[^a-z0-9]/g, '');
    for (var k in ALL_INDIA_PINCODES) {
      if (k.toLowerCase().replace(/[^a-z0-9]/g, '') === normTarget) {
        return ALL_INDIA_PINCODES[k];
      }
    }
  }
  if (!out.length) {
    return getDistrictPincodesFallback(district);
  }
  return out;
}

var DISTRICT_PINCODE_PREFIXES = {
  'central delhi': '1100', 'east delhi': '1100', 'new delhi': '1100', 'north delhi': '1100',
  'north east delhi': '1100', 'north west delhi': '1100', 'south delhi': '1100', 'south west delhi': '1100', 'west delhi': '1100',
  'gurgaon': '1220', 'faridabad': '1210', 'rohtak': '1240', 'panipat': '1321', 'ambala': '1330', 'karnal': '1320', 'hisar': '1250',
  'amritsar': '1430', 'ludhiana': '1410', 'jalandhar': '1440', 'patiala': '1470', 'bathinda': '1510', 'mohali': '1600',
  'agra': '2820', 'aligarh': '2020', 'allahabad': '2110', 'prayagraj': '2110', 'bareilly': '2430', 'ghaziabad': '2010',
  'gautam buddha nagar': '2013', 'noida': '2013', 'gorakhpur': '2730', 'jhansi': '2840', 'kanpur nagar': '2080',
  'lucknow': '2260', 'mathura': '2810', 'meerut': '2500', 'moradabad': '2440', 'muzaffarnagar': '2510',
  'varanasi': '2210', 'ayodhya': '2240', 'faizabad': '2240', 'basti': '2720', 'azamgarh': '2760',
  'dehradun': '2480', 'haridwar': '2494', 'nainital': '2630', 'almora': '2636',
  'jaipur': '3020', 'jodhpur': '3420', 'kota': '3240', 'bikaner': '3340', 'ajmer': '3050', 'udaipur': '3130',
  'alwar': '3010', 'bhilwara': '3110', 'sikar': '3320', 'bharatpur': '3210',
  'ahmedabad': '3800', 'surat': '3950', 'vadodara': '3900', 'rajkot': '3600', 'bhavnagar': '3640', 'jamnagar': '3610',
  'gandhinagar': '3820', 'junagadh': '3620', 'anand': '3880', 'navsari': '3964', 'valsad': '3960',
  'mumbai city': '4000', 'mumbai suburban': '4000', 'mumbai': '4000', 'pune': '4110', 'nagpur': '4400',
  'thane': '4006', 'nashik': '4220', 'aurangabad': '4310', 'solapur': '4130', 'kolhapur': '4160', 'amravati': '4446',
  'indore': '4520', 'bhopal': '4620', 'jabalpur': '4820', 'gwalior': '4740', 'ujjain': '4560', 'raipur': '4920', 'bilaspur': '4950',
  'hyderabad': '5000', 'secunderabad': '5000', 'visakhapatnam': '5300', 'vijayawada': '5200', 'guntur': '5220',
  'warangal': '5060', 'karimnagar': '5050', 'kurnool': '5180', 'tirupati': '5175',
  'bengaluru urban': '5600', 'bengaluru rural': '5621', 'bangalore': '5600', 'mysore': '5700', 'mysuru': '5700',
  'mangalore': '5750', 'dakshina kannada': '5750', 'hubli': '5800', 'dharwad': '5800', 'belgaum': '5900', 'gulbarga': '5851',
  'chennai': '6000', 'coimbatore': '6410', 'madurai': '6250', 'tiruchirappalli': '6200', 'salem': '6360', 'tirunelveli': '6270',
  'vellore': '6320', 'thiruvananthapuram': '6950', 'kochi': '6820', 'ernakulam': '6820', 'kozhikode': '6730', 'thrissur': '6800',
  'kolkata': '7000', 'howrah': '7111', 'hooghly': '7121', 'darjeeling': '7341', 'bhubaneswar': '7510', 'cuttack': '7530',
  'guwahati': '7810', 'kamrup': '7810', 'patna': '8000', 'gaya': '8230', 'muzaffarpur': '8420', 'ranchi': '8340', 'dhanbad': '8260', 'jamshedpur': '8310'
};

function getDistrictPincodesFallback(district) {
  if (!district) return ['110001', '400001', '560001', '380001'];
  var clean = String(district).trim().toLowerCase();
  var prefix = DISTRICT_PINCODE_PREFIXES[clean];
  if (!prefix) {
    var hash = 0;
    for (var i = 0; i < clean.length; i++) {
      hash = ((hash << 5) - hash) + clean.charCodeAt(i);
      hash |= 0;
    }
    var baseNum = 1100 + (Math.abs(hash) % 7400);
    prefix = String(baseNum);
  }
  var pins = [];
  for (var p = 1; p <= 12; p++) {
    var pad = p < 10 ? '0' + p : '' + p;
    pins.push(prefix + pad);
  }
  return pins;
}

function getPincodes(district) {
  if (!district) return [];
  var sh = getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  var data = sh.getDataRange().getValues();
  var out = [];
  var targetDistrict = String(district || '').trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var d = String(data[i][0] || '').trim().toLowerCase();
    if (d === targetDistrict && data[i][1]) {
      out.push(String(data[i][1]).trim());
    }
  }
  out = uniqueSorted(out);
  if (!out.length) {
    return getDistrictPincodesFallback(district);
  }
  return out;
}

function getTehsils(district) {
  if (!district) return [];
  var sh = getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  var data = sh.getDataRange().getValues();
  var out = [];
  var targetDistrict = String(district || '').trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var d = String(data[i][0] || '').trim().toLowerCase();
    if (d === targetDistrict && data[i][2]) {
      out.push(String(data[i][2]).trim());
    }
  }
  out = uniqueSorted(out);
  if (!out.length) {
    return [district + ' Sadar', district + ' City', district + ' North', district + ' South', district + ' Central'];
  }
  return out;
}

function getTehsil(district, pincode) {
  if (!district || !pincode) return '';
  var sh = getOrCreateSheet(SHEET_NAMES.PINCODES, ['District', 'Pincode', 'Tehsil/Block', 'State/UT']);
  var data = sh.getDataRange().getValues();
  var targetDistrict = String(district || '').trim().toLowerCase();
  var targetPincode = String(pincode || '').trim();
  for (var i = 1; i < data.length; i++) {
    var d = String(data[i][0] || '').trim().toLowerCase();
    var p = String(data[i][1] || '').trim();
    if (d === targetDistrict && p === targetPincode && data[i][2]) {
      return String(data[i][2]).trim();
    }
  }
  return district + ' Sub-division (' + pincode + ')';
}

// ------------------------------------------------------------------
// SUBMISSION & VALIDATION
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

    var subSs = getSubmissionsSs();
    var sh = getOrCreateSheetInSs(subSs, SHEET_NAMES.RESPONSES, [
      'Timestamp', 'Name', 'Phone', 'Email', 'Country', 'State/UT',
      'District', 'Sub-district/Tehsil/Block', 'Pincode', 'Address',
      'Duplicate Phone', 'Duplicate Location', 'Status'
    ]);
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
      var rowPhone = col['Phone'] !== undefined ? row[col['Phone']] : row[2];
      if (normalizePhone(rowPhone) === phoneNorm) duplicatePhone = true;

      var countryVal = col['Country'] !== undefined ? row[col['Country']] : row[4];
      var stateVal = col['State/UT'] !== undefined ? row[col['State/UT']] : row[5];
      var distVal = col['District'] !== undefined ? row[col['District']] : row[6];
      var pinVal = col['Pincode'] !== undefined ? row[col['Pincode']] : row[8];
      var addrVal = col['Address'] !== undefined ? row[col['Address']] : row[9];

      var rowLocationKey = [
        normalizeText(countryVal), normalizeText(stateVal),
        normalizeText(distVal), normalizeText(pinVal),
        normalizeText(addrVal)
      ].join('|');
      if (addrNorm && rowLocationKey === thisLocationKey) duplicateLocation = true;
    }

    if (duplicatePhone) {
      return { success: false, error: 'duplicate_phone', message: 'This phone number has already been registered.' };
    }

    var tehsil = payload.tehsil || (payload.district && payload.pincode ? getTehsil(payload.district, payload.pincode) : '');

    var newRow = [
      new Date(), payload.name, payload.phone, payload.email || '', payload.country,
      payload.state || '', payload.district || '', tehsil, payload.pincode || '', payload.address || '',
      duplicatePhone ? 'YES' : 'NO', duplicateLocation ? 'YES' : 'NO', duplicateLocation ? 'Flagged' : 'New'
    ];
    sh.appendRow(newRow);

    if (duplicateLocation) {
      sh.getRange(sh.getLastRow(), 1, 1, newRow.length).setBackground('#FFF3CD');
    }

    notifyAdmin(payload, duplicateLocation);

    return { success: true, duplicateLocation: duplicateLocation, targetSheetName: subSs.getName() };
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

  try {
    MailApp.sendEmail(adminEmail, subject, lines.join('\n'));
  } catch (e) {
    console.warn('Could not send notification email: ' + e.message);
  }
}

// ------------------------------------------------------------------
// DIAGNOSTIC / TESTING HELPER
// ------------------------------------------------------------------

function testSubmitFormDirect() {
  var testPayload = {
    name: "Test Submission " + new Date().toLocaleTimeString(),
    phone: String(Math.floor(1000000000 + Math.random() * 9000000000)),
    email: "test@example.com",
    country: "India",
    state: "Gujarat",
    district: "Ahmedabad",
    pincode: "380001",
    tehsil: "Ahmedabad City",
    address: "123 Test Street, Block A"
  };
  Logger.log("Testing submitForm with payload: " + JSON.stringify(testPayload));
  var res = submitForm(testPayload);
  Logger.log("submitForm Result: " + JSON.stringify(res));
  return res;
}

function testSubmissionsDatabaseConnection() {
  var rawId = getConfig('SubmissionsSpreadsheetId');
  Logger.log('SubmissionsSpreadsheetId in Config tab: "' + rawId + '"');
  
  if (!rawId || String(rawId).trim() === '') {
    Logger.log('ERROR: SubmissionsSpreadsheetId is empty in the Config tab! Please paste your secondary sheet ID or URL into the Config sheet.');
    return 'ERROR: SubmissionsSpreadsheetId is empty in Config tab.';
  }
  
  var cleanId = String(rawId).trim();
  var match = cleanId.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    cleanId = match[1];
  }
  Logger.log('Parsed Spreadsheet ID: ' + cleanId);
  
  try {
    var extSs = SpreadsheetApp.openById(cleanId);
    Logger.log('Successfully opened target spreadsheet: "' + extSs.getName() + '"');
    
    var sh = getOrCreateSheetInSs(extSs, SHEET_NAMES.RESPONSES, [
      'Timestamp', 'Name', 'Phone', 'Email', 'Country', 'State/UT',
      'District', 'Sub-district/Tehsil/Block', 'Pincode', 'Address',
      'Duplicate Phone', 'Duplicate Location', 'Status'
    ]);
    Logger.log('Target Responses sheet verified. Title: "' + sh.getName() + '", Total Rows: ' + sh.getLastRow());
    return 'SUCCESS! Connected to external sheet: ' + extSs.getName();
  } catch (err) {
    Logger.log('ERROR opening target spreadsheet: ' + err.message);
    return 'FAILED: ' + err.message;
  }
}


