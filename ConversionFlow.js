function updateConversionFlowWeekly() {
  var PROPERTY_ID   = '418611571';
  var SPREADSHEET_ID = '1VxsGju5iO2WUInUJiy8zyUfdQ5ZyObN3NirxcAvQs94';
  var TAB_NAME       = 'Conversion Flow';
  var DRIVE_FOLDER_ID = '1cDY3s5pK99jHkSuliIifjrI_M3Fa245b';

  var propertyName = 'properties/' + PROPERTY_ID;

  // 1) Periods
  var curr = getLastFullWeek();          // {startDate,endDate} YYYY-MM-DD
  var prev = shiftPeriodDays_(curr, -7);
  var yoy  = shiftPeriodDays_(curr, -364);

  var weekLabel = formatDMY_(curr.startDate) + ' / ' + formatDMY_(curr.endDate);

  // 2) Pull GA4 counts for each stage (curr/prev/yoy)
  var gaCurr = fetchJourneyCountsGA4_(propertyName, curr.startDate, curr.endDate);
  var gaPrev = fetchJourneyCountsGA4_(propertyName, prev.startDate, prev.endDate);
  var gaYoY  = fetchJourneyCountsGA4_(propertyName, yoy.startDate,  yoy.endDate);

  // 3) Device shares (Desktop%, iOS%, Android%) for key GA4 stages (curr week only)
  // For conversion rows we leave device shares blank.
  var shares = {};
  shares.sessions           = getDeviceSharesForStage_(propertyName, 'sessions', null, null, curr);
  shares.listingViews       = getDeviceSharesForStage_(propertyName, 'event', 'listing_page_view', null, curr);
  shares.tripCartInquiry    = getDeviceSharesForStage_(propertyName, 'event', 'trip-cart_price-calculated', 'false', curr);
  shares.inquiryStart       = getDeviceSharesForStage_(propertyName, 'event', 'inquiry_start', null, curr);
  shares.inquirySubmit      = getDeviceSharesForStage_(propertyName, 'event', 'inquiry_submit_success', null, curr);
  shares.tripCartBookNow    = getDeviceSharesForStage_(propertyName, 'event', 'trip-cart_price-calculated', 'true', curr);
  shares.bookNowClicks      = getDeviceSharesForStage_(propertyName, 'event', 'trip-cart_book-now-click', null, curr);

  // 4) CSV system-of-record (curr/prev/yoy)
  var csvCurr = fetchCsvTotals_(DRIVE_FOLDER_ID, curr.startDate, curr.endDate);
  var csvPrev = fetchCsvTotals_(DRIVE_FOLDER_ID, prev.startDate, prev.endDate);
  var csvYoY  = fetchCsvTotals_(DRIVE_FOLDER_ID, yoy.startDate,  yoy.endDate);

  // 5) Build rows in the order of your mock
  var rows = [];

  // Helper to push a row (no Drop-off column)
  function pushRow(stageLabel, valueCurr, valuePrev, valueYoY, sharesKey, isConversion) {
    var wow  = pctChange_(Number(valueCurr||0), Number(valuePrev||0));
    var yoyp = pctChange_(Number(valueCurr||0), Number(valueYoY||0));

    var desktop = '', ios = '', android = '';
    if (!isConversion && sharesKey && shares[sharesKey]) {
      desktop = pctStr_(shares[sharesKey].desktop);
      ios     = pctStr_(shares[sharesKey].ios);
      android = pctStr_(shares[sharesKey].android);
    }

    rows.push([
      weekLabel,
      stageLabel,
      isConversion ? pctStr_(valueCurr) : Number(valueCurr || 0),
      pctStr_(wow),
      pctStr_(yoyp),
      desktop,
      ios,
      android
    ]);
  }

  // TOP (orange)
  pushRow('Sessions',
          gaCurr.sessions, gaPrev.sessions, gaYoY.sessions,
          'sessions', false);

  pushRow('Listing Views',
          gaCurr.listingViews, gaPrev.listingViews, gaYoY.listingViews,
          'listingViews', false);

  // INQUIRY (blue)
  pushRow('Send Inquiry Trip Cart', // trip-cart_price-calculated + p2 = false
          gaCurr.tripCartInquiry, gaPrev.tripCartInquiry, gaYoY.tripCartInquiry,
          'tripCartInquiry', false);

  pushRow('Inquiry Start',
          gaCurr.inquiryStart, gaPrev.inquiryStart, gaYoY.inquiryStart,
          'inquiryStart', false);

  pushRow('Inquiry Submitted',
          gaCurr.inquirySubmit, gaPrev.inquirySubmit, gaYoY.inquirySubmit,
          'inquirySubmit', false);

  pushRow('Bookings', // (Inquiry)
          csvCurr.completedBookings, csvPrev.completedBookings, csvYoY.completedBookings,
          null, false);

  // Conversion % = Bookings / Send Inquiry Trip Cart
  var inquiryConvCurr = safeRatio_(csvCurr.completedBookings, gaCurr.inquirySubmit);
  var inquiryConvPrev = safeRatio_(csvPrev.completedBookings, gaPrev.inquirySubmit);
  var inquiryConvYoY  = safeRatio_(csvYoY.completedBookings,  gaYoY.inquirySubmit);
  pushRow('Inquiry Conversion', // (Bookings / Send Inquiry Trip Cart)
          inquiryConvCurr, inquiryConvPrev, inquiryConvYoY,
          null, true);

  // INSTABOOK (green)
  pushRow('Book Now Trip Cart', // (trip-cart_price-calculated + p2 = true)
          gaCurr.tripCartBookNow, gaPrev.tripCartBookNow, gaYoY.tripCartBookNow,
          'tripCartBookNow', false);

  pushRow('Book Now Clicks',
          gaCurr.bookNowClicks, gaPrev.bookNowClicks, gaYoY.bookNowClicks,
          'bookNowClicks', false);

  pushRow('Instabook Confirmed',
          csvCurr.instabookConfirmed, csvPrev.instabookConfirmed, csvYoY.instabookConfirmed,
          null, false);

  var instaConvCurr = safeRatio_(csvCurr.instabookConfirmed, gaCurr.bookNowClicks);
  var instaConvPrev = safeRatio_(csvPrev.instabookConfirmed, gaPrev.bookNowClicks);
  var instaConvYoY  = safeRatio_(csvYoY.instabookConfirmed,  gaYoY.bookNowClicks); 
  pushRow('Instabook Conversion', // (Instabook Confirmed / Book Now Trip Cart)
          instaConvCurr, instaConvPrev, instaConvYoY,
          null, true);

  // 6) Append to sheet (no clearing), then color the 3 sections
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TAB_NAME) || ss.insertSheet(TAB_NAME);

  // Ensure header
  var headers = ['Week (UTC)','Stage','Users/Sessions','WoW %','YoY %','Desktop %','iOS %','Android %'];
  if (needsHeader_(sh, headers.length, headers)) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#D9D9D9');
  }

  var startRow = sh.getLastRow() + 1;
  if (rows.length) {
    sh.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
  }

  // Updated colors
  var BLUE  = '#C6DEF0'; // Inquiry flow
  var GREEN = '#D3EED3'; // Instabook flow

  // Clear backgrounds for all new rows first
  sh.getRange(startRow, 1, rows.length, headers.length).setBackground(null);

  // Apply colors explicitly by row
  // No color rows
  // Row 1: Sessions
  // Row 2: Listing Views

  // Blue rows
  sh.getRange(startRow + 2, 1, 5, headers.length).setBackground(BLUE); // Send Inquiry Trip Cart, Inquiry Start, Inquiry Submitted, Bookings, Inquiry Conversion

  // Green rows
  sh.getRange(startRow + 7, 1, 4, headers.length).setBackground(GREEN); // Book Now Trip Cart, Book Now Clicks, Instabook Confirmed, Instabook Conversion

  Logger.log('Customer Journey appended for ' + weekLabel);
}

/* ---------------- GA4 pulling ---------------- */

// Fetches all GA4 counts needed for the journey for a given date range.
function fetchJourneyCountsGA4_(propertyName, startDate, endDate) {
  var out = {
    sessions: 0,
    listingViews: 0,
    tripCartInquiry: 0,  // p2=false
    inquiryStart: 0,
    inquirySubmit: 0,
    tripCartBookNow: 0,  // p2=true
    bookNowClicks: 0
  };

  // Sessions (metric)
  out.sessions = readSingleMetric_(AnalyticsData.Properties.runReport({
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    metrics: [{ name: 'sessions' }]
  }, propertyName), 0);

  // listing_page_view (unique users)
  out.listingViews = readUsersForEvent_(propertyName, startDate, endDate, 'listing_page_view', null);

  // trip-cart_price-calculated with p2=false (Inquiry trip cart)
  out.tripCartInquiry = readUsersForEvent_(propertyName, startDate, endDate, 'trip-cart_price-calculated', 'false');

  // inquiry_start
  out.inquiryStart = readUsersForEvent_(propertyName, startDate, endDate, 'inquiry_start', null);

  // inquiry_submit_success
  out.inquirySubmit = readUsersForEvent_(propertyName, startDate, endDate, 'inquiry_submit_success', null);

  // trip-cart_price-calculated with p2=true (Book Now trip cart)
  out.tripCartBookNow = readUsersForEvent_(propertyName, startDate, endDate, 'trip-cart_price-calculated', 'true');

  // trip-cart_book-now-click
  out.bookNowClicks = readUsersForEvent_(propertyName, startDate, endDate, 'trip-cart_book-now-click', null);

  return out;
}

// Reads totalUsers for a specific event, optionally filtered by customEvent:p2 ("true"/"false").
function readUsersForEvent_(propertyName, startDate, endDate, eventName, p2Value) {
  var filter = {
    andGroup: {
      expressions: [{
        filter: { fieldName: 'eventName', stringFilter: { value: eventName, matchType: 'EXACT' } }
      }]
    }
  };

  if (p2Value !== null && p2Value !== undefined) {
    filter.andGroup.expressions.push({
      filter: { fieldName: 'customEvent:p2', stringFilter: { value: p2Value, matchType: 'EXACT' } }
    });
  }

  var resp = AnalyticsData.Properties.runReport({
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    metrics: [{ name: 'totalUsers' }],
    dimensionFilter: filter
  }, propertyName);

  return readSingleMetric_(resp, 0);
}

// Returns Desktop%, iOS%, Android% for a stage (sessions or an event with totalUsers).
// stageType: 'sessions' or 'event'
// eventName: only for stageType='event'
// p2Value: 'true'/'false' or null
function getDeviceSharesForStage_(propertyName, stageType, eventName, p2Value, period) {
  var result = { desktop: 0, ios: 0, android: 0 };

  if (stageType === 'sessions') {
    // Desktop share by deviceCategory
    var devResp = AnalyticsData.Properties.runReport({
      dateRanges: [{ startDate: period.startDate, endDate: period.endDate }],
      dimensions: [{ name: 'deviceCategory' }],
      metrics: [{ name: 'sessions' }],
      limit: 1000
    }, propertyName);
    var devMap = pivotCount_(devResp, 0, 0);
    var devTotal = sumValues_(devMap);
    result.desktop = devTotal ? (Number(devMap['desktop'] || 0) / devTotal) : 0;

    // OS share (iOS/Android)
    var osResp = AnalyticsData.Properties.runReport({
      dateRanges: [{ startDate: period.startDate, endDate: period.endDate }],
      dimensions: [{ name: 'operatingSystem' }],
      metrics: [{ name: 'sessions' }],
      limit: 10000
    }, propertyName);
    var osMap = pivotCount_(osResp, 0, 0);
    var osTotal = sumValues_(osMap);
    result.ios     = osTotal ? (Number(osMap['iOS']     || 0) / osTotal) : 0;
    result.android = osTotal ? (Number(osMap['Android'] || 0) / osTotal) : 0;

    return result;
  }

  // Event stages: use totalUsers and filter by event (+ p2 if provided)
  var baseFilter = {
    andGroup: {
      expressions: [{
        filter: { fieldName: 'eventName', stringFilter: { value: eventName, matchType: 'EXACT' } }
      }]
    }
  };
  if (p2Value !== null && p2Value !== undefined) {
    baseFilter.andGroup.expressions.push({
      filter: { fieldName: 'customEvent:p2', stringFilter: { value: p2Value, matchType: 'EXACT' } }
    });
  }

  // Desktop share
  var devResp2 = AnalyticsData.Properties.runReport({
    dateRanges: [{ startDate: period.startDate, endDate: period.endDate }],
    dimensions: [{ name: 'deviceCategory' }],
    metrics: [{ name: 'totalUsers' }],
    dimensionFilter: baseFilter,
    limit: 1000
  }, propertyName);
  var dMap2 = pivotCount_(devResp2, 0, 0);
  var dTot2 = sumValues_(dMap2);
  result.desktop = dTot2 ? (Number(dMap2['desktop'] || 0) / dTot2) : 0;

  // OS share (iOS/Android)
  var osResp2 = AnalyticsData.Properties.runReport({
    dateRanges: [{ startDate: period.startDate, endDate: period.endDate }],
    dimensions: [{ name: 'operatingSystem' }],
    metrics: [{ name: 'totalUsers' }],
    dimensionFilter: baseFilter,
    limit: 10000
  }, propertyName);
  var osMap2 = pivotCount_(osResp2, 0, 0);
  var osTot2 = sumValues_(osMap2);
  result.ios     = osTot2 ? (Number(osMap2['iOS']     || 0) / osTot2) : 0;
  result.android = osTot2 ? (Number(osMap2['Android'] || 0) / osTot2) : 0;

  return result;
}

function readSingleMetric_(resp, metricIndex) {
  if (!resp || !resp.rows || !resp.rows.length) return 0;
  var v = resp.rows[0].metricValues && resp.rows[0].metricValues[metricIndex]
        ? resp.rows[0].metricValues[metricIndex].value : 0;
  return Number(v || 0);
}

/* ---------------- CSV (Drive) ---------------- */

function fetchCsvTotals_(folderId, startDate, endDate) {
  // Sums n across all rows in all files matching date range.
  // Filenames supported: completed-booking-YYYY-MM-DD.csv, instabook-YYYY-MM-DD.csv
  var out = {
    completedBookings: 0,
    instabookConfirmed: 0
    // You can add "inquiries", "offers" here later (wi-*.csv, offers-*.csv)
  };

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var name = f.getName();
    var lower = name.toLowerCase();
    var blob = f.getBlob();
    var text = blob.getDataAsString();
    var sum = sumCsvRange_(text, startDate, endDate); // sums 'n' within range

    if (lower.indexOf('weekly-bookings-non-ib-') === 0) {
      out.completedBookings += sum;
    } else if (lower.indexOf('instabook-') === 0) {
      out.instabookConfirmed += sum;
    }
    // else if (lower.indexOf('wi-') === 0) { out.inquiries += sum; }
    // else if (lower.indexOf('offers-') === 0) { out.offers += sum; }
  }
  return out;
}

function sumCsvRange_(csvText, startDate, endDate) {
  // CSV schema: date_created,n
  var lines = csvText.split(/\r?\n/);
  var total = 0;
  for (var i = 1; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    var parts = line.split(',');
    if (parts.length < 2) continue;
    var d = parts[0];
    var n = Number(parts[1] || 0);
    if (d >= startDate && d <= endDate) total += n;
  }
  return total;
}

/* ---------------- Utilities ---------------- */

function getLastFullWeek() {
  var today = new Date();
  var utcToday = new Date(today.getTime() + today.getTimezoneOffset() * 60000);
  var day = utcToday.getUTCDay(); // 0=Sun
  var startOfThisWeek = new Date(Date.UTC(
    utcToday.getUTCFullYear(),
    utcToday.getUTCMonth(),
    utcToday.getUTCDate() - day
  ));
  var start = new Date(startOfThisWeek);
  start.setUTCDate(start.getUTCDate() - 7);
  var end = new Date(startOfThisWeek);
  end.setUTCDate(end.getUTCDate() - 1);
  return { startDate: ymd_(start), endDate: ymd_(end) };
}

function shiftPeriodDays_(period, deltaDays) {
  var s = parseYMD_(period.startDate);
  var e = parseYMD_(period.endDate);
  s.setUTCDate(s.getUTCDate() + deltaDays);
  e.setUTCDate(e.getUTCDate() + deltaDays);
  return { startDate: ymd_(s), endDate: ymd_(e) };
}

function ymd_(d) {
  var y = d.getUTCFullYear();
  var m = ('0' + (d.getUTCMonth() + 1)).slice(-2);
  var day = ('0' + d.getUTCDate()).slice(-2);
  return y + '-' + m + '-' + day;
}

function parseYMD_(s) {
  var p = s.split('-');
  return new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2])));
}

function formatDMY_(ymd) {
  var p = ymd.split('-'); // yyyy-mm-dd
  return p[2] + '-' + p[1] + '-' + p[0];
}

function pctChange_(curr, prev) {
  curr = Number(curr || 0);
  prev = Number(prev || 0);
  if (!prev) return null;
  return (curr - prev) / prev;
}

function pctStr_(ratio) {
  if (ratio === null || typeof ratio === 'undefined') return '—';
  return (Number(ratio) * 100).toFixed(1) + '%';
}

function safeRatio_(num, den) {
  num = Number(num || 0);
  den = Number(den || 0);
  if (!den) return 0;
  return num / den;
}

function pivotCount_(resp, dimIndex, metricIndex) {
  var map = {};
  var rows = (resp && resp.rows) ? resp.rows : [];
  for (var i = 0; i < rows.length; i++) {
    var key = rows[i].dimensionValues[dimIndex].value;
    var cnt = Number(rows[i].metricValues[metricIndex].value || 0);
    map[key] = (map[key] || 0) + cnt;
  }
  return map;
}

function sumValues_(obj) {
  var sum = 0;
  for (var k in obj) sum += Number(obj[k] || 0);
  return sum;
}

function needsHeader_(sheet, width, headers) {
  var first = sheet.getRange(1, 1, 1, width).getValues()[0];
  if (!first || first.join('') === '') return true;
  if (first.length !== headers.length) return true;
  for (var i = 0; i < headers.length; i++) {
    if (first[i] !== headers[i]) return true;
  }
  return false;
}