// ===== Debug toggle =====

var DP_DEBUG = true; // set to false to silence logs

// Optional hard override: if you KNOW the apiName, set it here (e.g., 'customEvent:is_webview').
// Leave empty string to auto-resolve.
var DP_WV_DIM_OVERRIDE = 'customEvent:is_webview';

// Returns start/end date (UTC) of last full week (Sunday–Saturday)
function DP_getLastFullWeek() {
  var today = new Date();
  var utc = new Date(today.getTime() + today.getTimezoneOffset() * 60000);
  var day = utc.getUTCDay(); // Sunday = 0
  var startOfThisWeek = new Date(Date.UTC(
    utc.getUTCFullYear(),
    utc.getUTCMonth(),
    utc.getUTCDate() - day
  ));
  var start = new Date(startOfThisWeek);
  start.setUTCDate(start.getUTCDate() - 7);
  var end = new Date(startOfThisWeek);
  end.setUTCDate(end.getUTCDate() - 1);
  return { startDate: DP_ymd_(start), endDate: DP_ymd_(end) };
}

// Formats a Date as 'YYYY-MM-DD' in UTC
function DP_ymd_(d) {
  var y = d.getUTCFullYear();
  var m = ('0' + (d.getUTCMonth() + 1)).slice(-2);
  var day = ('0' + d.getUTCDate()).slice(-2);
  return y + '-' + m + '-' + day;
}

/**
 * Converts a date string 'YYYY-MM-DD' to 'DD-MM-YYYY'
 */
function DP_formatDMY_(ymd) {
  var parts = ymd.split('-'); // yyyy-mm-dd
  return parts[2] + '-' + parts[1] + '-' + parts[0];
}

function updateDevicePlatformPerformanceWeekly() {
  var PROPERTY_ID     = '418611571';
  var SPREADSHEET_ID  = '1VxsGju5iO2WUInUJiy8zyUfdQ5ZyObN3NirxcAvQs94';
  var TAB_NAME        = 'Device & Platform Performance';

  var propertyName = 'properties/' + PROPERTY_ID;

  // Period windows
  var curr = DP_getLastFullWeek();
  var weekLabel = DP_formatDMY_(curr.startDate) + ' / ' + DP_formatDMY_(curr.endDate);

  // Define segments (keys are used in filter builder)
  var SEGMENTS = [
    { key: 'desktop-web',   label: 'Desktop' },
    { key: 'ios-web',       label: 'iOS - Mobweb' },
    { key: 'ios-app',       label: 'iOS - App' },
    { key: 'android-web',   label: 'Android - Mobweb' },
    { key: 'android-app',   label: 'Android - App' }
  ];

  // Stages (eventName + optional p2)
  var STAGES = [
    { label: 'Send Inquiry Trip Cart', event: 'trip-cart_price-calculated', p2: 'false' },
    { label: 'Inquiry Start',          event: 'inquiry_start' },
    { label: 'Inquiry Submitted',      event: 'inquiry_submit_success' },
    { label: 'Book Now Trip Cart',     event: 'trip-cart_price-calculated', p2: 'true' },
    { label: 'Book Now Clicks',        event: 'trip-cart_book-now-click' },
    { label: 'Proceed to Payment',     event: 'trip-cart_book-now-proceed-to-payment-cl' }
  ];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TAB_NAME) || ss.insertSheet(TAB_NAME);

  // ---------- Header (2 rows, grouped bands) ----------
  var header1 = [
    'Week (UTC)', 'Stage',
    'Desktop %', 'iOS Mobweb %', 'iOS App %', 'Android Mobweb %', 'Android App %',
    'Desktop', '', 'iOS - Mobweb', '', 'iOS - App', '', 'Android - Mobweb', '', 'Android - App', ''
  ];
  var header2 = [
    '', '',
    '','', '', '', '',
    'Users','Drop-off',
    'Users','Drop-off',
    'Users','Drop-off',
    'Users','Drop-off',
    'Users','Drop-off'
  ];

  if (sh.getLastRow() < 2 || sh.getRange(1,1,1,header1.length).getValues()[0].join('') === '') {
    sh.getRange(1,1,1,header1.length).setValues([header1]).setFontWeight('bold');
    sh.getRange(2,1,1,header2.length).setValues([header2]).setFontWeight('bold');

    // Merge device group headers (users/drop-off bands)
    sh.getRange(1,8,1,2).merge();   // Desktop
    sh.getRange(1,10,1,2).merge();  // iOS - Mobweb
    sh.getRange(1,12,1,2).merge();  // iOS - App
    sh.getRange(1,14,1,2).merge();  // Android - Mobweb
    sh.getRange(1,16,1,2).merge();  // Android - App

    // Band colors (users/drop-off bands)
    var desktopColor = '#e8e0f5'; // desktop
    var iosWebColor  = '#f9efc3'; // iOS web
    var iosAppColor  = '#f6e69c'; // iOS app
    var andWebColor  = '#c6def0'; // Android web
    var andAppColor  = '#b9d7ee'; // Android app

    sh.getRange(1,8,2,2).setBackground(desktopColor);
    sh.getRange(1,10,2,2).setBackground(iosWebColor);
    sh.getRange(1,12,2,2).setBackground(iosAppColor);
    sh.getRange(1,14,2,2).setBackground(andWebColor);
    sh.getRange(1,16,2,2).setBackground(andAppColor);
  }

  // ---------- Pull data: totalUsers for each stage × segment (current week only) ----------
  var outRows = []; // 6 rows

  // Pre-read counts per segment so we can compute drop-offs vs previous stage per flow
  var countsBySeg = {}; // countsBySeg[segKey] = [values per stage]
  for (var g = 0; g < SEGMENTS.length; g++) {
    var seg = SEGMENTS[g];
    countsBySeg[seg.key] = [];
    for (var s = 0; s < STAGES.length; s++) {
      var st = STAGES[s];
      var val = DP_readUsersForEventBySegment_('properties/' + PROPERTY_ID, curr.startDate, curr.endDate, st.event, st.p2, seg.key);
      countsBySeg[seg.key].push(Number(val || 0));
    }
  }

  function dropPct(curr, prev) {
    curr = Number(curr || 0); prev = Number(prev || 0);
    if (!prev) return '—';
    var r = (prev - curr) / prev;
    if (r < 0) r = 0; // guard against increases
    return (r * 100).toFixed(1) + '%';
  }

  for (var s2 = 0; s2 < STAGES.length; s2++) {
    var stage = STAGES[s2];
    var row = [weekLabel, stage.label];

    // Compute total users for this stage across all 5 segments
    var usersPerDevice = [];
    for (var g2 = 0; g2 < SEGMENTS.length; g2++) {
      var seg2 = SEGMENTS[g2];
      var series = countsBySeg[seg2.key];
      usersPerDevice.push(series[s2] || 0);
    }
    var totalUsers = usersPerDevice.reduce(function(a,b){return a+b;}, 0);

    // Compute device % (as decimals, e.g. 0.1234)
    for (var i = 0; i < usersPerDevice.length; i++) {
      var pct = (totalUsers > 0) ? usersPerDevice[i] / totalUsers : 0;
      row.push(pct);
    }

    // For each segment, compute users + drop-off vs previous stage in same flow
    for (var g2 = 0; g2 < SEGMENTS.length; g2++) {
      var seg2 = SEGMENTS[g2];
      var series = countsBySeg[seg2.key];
      var users = series[s2] || 0;
      var drop = '—';
      // Inquiry flow indices 0,1,2 ; Instabook flow indices 3,4,5
      if (s2 === 1) drop = dropPct(series[1], series[0]);
      if (s2 === 2) drop = dropPct(series[2], series[1]);
      if (s2 === 4) drop = dropPct(series[4], series[3]);
      if (s2 === 5) drop = dropPct(series[5], series[4]);
      row.push(users, drop);
    }

    outRows.push(row);
  }

  // ---------- Append rows ----------
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, outRows.length, 17).setValues(outRows);

  Logger.log('Device & Platform Pivot (Drop-off) appended for ' + weekLabel + ' (' + outRows.length + ' rows)');
}

/**
 * Backfill Device & Platform Performance for a range of weeks.
 * @param {string} startDateStr - Start date in 'yyyy-mm-dd'
 * @param {string} endDateStr - End date in 'yyyy-mm-dd'
 */
function backfillDevicePlatformPerformance(startDateStr, endDateStr) {
  var PROPERTY_ID     = '418611571';
  var SPREADSHEET_ID  = '1VxsGju5iO2WUInUJiy8zyUfdQ5ZyObN3NirxcAvQs94';
  var TAB_NAME        = 'Device & Platform Performance';
  var propertyName = 'properties/' + PROPERTY_ID;

  // Helper: Week-aligned: find first Sunday on/after startDate, then each Sunday–Saturday window until endDate
  function toUTCDate(str) {
    if (typeof str === "undefined" || str === null) {
      throw new Error("toUTCDate: date string is undefined. Argument required in format 'YYYY-MM-DD'.");
    }
    var p = str.split('-');
    return new Date(Date.UTC(Number(p[0]), Number(p[1])-1, Number(p[2])));
  }

  // Handle missing arguments: default to 2025-08-17 as startDate, yesterday as endDate
  var defaultStart = "2025-08-17";
  var defaultEnd = (function() {
    var now = new Date();
    now.setUTCDate(now.getUTCDate() - 1);
    var y = now.getUTCFullYear();
    var m = ('0' + (now.getUTCMonth() + 1)).slice(-2);
    var d = ('0' + now.getUTCDate()).slice(-2);
    return y + '-' + m + '-' + d;
  })();
  var actualStartDateStr = startDateStr || defaultStart;
  var actualEndDateStr = endDateStr || defaultEnd;
  if (DP_DEBUG) {
    Logger.log('[Backfill] Using date range: ' + actualStartDateStr + ' to ' + actualEndDateStr);
  }
  var startDate = toUTCDate(actualStartDateStr);
  var endDate = toUTCDate(actualEndDateStr);

  // Find first Sunday on/after startDate
  var currStart = new Date(startDate);
  var day = currStart.getUTCDay();
  if (day !== 0) { // not Sunday
    currStart.setUTCDate(currStart.getUTCDate() + (7 - day));
  }

  // The header, segments, stages, etc, match updateDevicePlatformPerformanceWeekly
  var SEGMENTS = [
    { key: 'desktop-web',   label: 'Desktop' },
    { key: 'ios-web',       label: 'iOS - Mobweb' },
    { key: 'ios-app',       label: 'iOS - App' },
    { key: 'android-web',   label: 'Android - Mobweb' },
    { key: 'android-app',   label: 'Android - App' }
  ];
  var STAGES = [
    { label: 'Send Inquiry Trip Cart', event: 'trip-cart_price-calculated', p2: 'false' },
    { label: 'Inquiry Start',          event: 'inquiry_start' },
    { label: 'Inquiry Submitted',      event: 'inquiry_submit_success' },
    { label: 'Book Now Trip Cart',     event: 'trip-cart_price-calculated', p2: 'true' },
    { label: 'Book Now Clicks',        event: 'trip-cart_book-now-click' },
    { label: 'Proceed to Payment',     event: 'trip-cart_book-now-proceed-to-payment-cl' }
  ];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TAB_NAME) || ss.insertSheet(TAB_NAME);

  // Ensure header
  var header1 = [
    'Week (UTC)', 'Stage',
    'Desktop %', 'iOS Mobweb %', 'iOS App %', 'Android Mobweb %', 'Android App %',
    'Desktop', '', 'iOS - Mobweb', '', 'iOS - App', '', 'Android - Mobweb', '', 'Android - App', ''
  ];
  var header2 = [
    '', '',
    '','', '', '', '',
    'Users','Drop-off',
    'Users','Drop-off',
    'Users','Drop-off',
    'Users','Drop-off',
    'Users','Drop-off'
  ];
  if (sh.getLastRow() < 2 || sh.getRange(1,1,1,header1.length).getValues()[0].join('') === '') {
    sh.getRange(1,1,1,header1.length).setValues([header1]).setFontWeight('bold');
    sh.getRange(2,1,1,header2.length).setValues([header2]).setFontWeight('bold');
    // Merge device group headers (users/drop-off bands)
    sh.getRange(1,8,1,2).merge();   // Desktop
    sh.getRange(1,10,1,2).merge();  // iOS - Mobweb
    sh.getRange(1,12,1,2).merge();  // iOS - App
    sh.getRange(1,14,1,2).merge();  // Android - Mobweb
    sh.getRange(1,16,1,2).merge();  // Android - App
    // Band colors
    var desktopColor = '#e8e0f5'; // desktop
    var iosWebColor  = '#f9efc3'; // iOS web
    var iosAppColor  = '#f6e69c'; // iOS app
    var andWebColor  = '#c6def0'; // Android web
    var andAppColor  = '#b9d7ee'; // Android app
    sh.getRange(1,8,2,2).setBackground(desktopColor);
    sh.getRange(1,10,2,2).setBackground(iosWebColor);
    sh.getRange(1,12,2,2).setBackground(iosAppColor);
    sh.getRange(1,14,2,2).setBackground(andWebColor);
    sh.getRange(1,16,2,2).setBackground(andAppColor);
  }

  // Read existing rows to check for weekLabel existence
  var data = sh.getDataRange().getValues();
  var weekStageToRowIdx = {};
  for (var i = 2; i < data.length; i++) {
    var row = data[i];
    var week = row[0];
    var stage = row[1];
    if (week && stage) {
      weekStageToRowIdx[week + '|' + stage] = i + 1; // 1-based row index for setValues
    }
  }

  // Helper for drop-off %
  function dropPct(curr, prev) {
    curr = Number(curr || 0); prev = Number(prev || 0);
    if (!prev) return '—';
    var r = (prev - curr) / prev;
    if (r < 0) r = 0;
    return (r * 100).toFixed(1) + '%';
  }

  // Loop through each week
  while (currStart <= endDate) {
    var weekStart = new Date(currStart);
    var weekEnd = new Date(currStart); weekEnd.setUTCDate(weekEnd.getUTCDate() + 6);
    if (weekEnd > endDate) weekEnd = new Date(endDate); // Clamp last week to endDate
    var weekLabel = DP_formatDMY_(DP_ymd_(weekStart)) + ' / ' + DP_formatDMY_(DP_ymd_(weekEnd));

    // Pre-read counts per segment for this week
    var countsBySeg = {};
    for (var g = 0; g < SEGMENTS.length; g++) {
      var seg = SEGMENTS[g];
      countsBySeg[seg.key] = [];
      for (var s = 0; s < STAGES.length; s++) {
        var st = STAGES[s];
        var val = DP_readUsersForEventBySegment_(propertyName, DP_ymd_(weekStart), DP_ymd_(weekEnd), st.event, st.p2, seg.key);
        countsBySeg[seg.key].push(Number(val || 0));
      }
    }

    // Prepare rows for this week
    var weekRows = [];
    for (var s2 = 0; s2 < STAGES.length; s2++) {
      var stage = STAGES[s2];
      var row = [weekLabel, stage.label];
      var usersPerDevice = [];
      for (var g2 = 0; g2 < SEGMENTS.length; g2++) {
        var seg2 = SEGMENTS[g2];
        var series = countsBySeg[seg2.key];
        usersPerDevice.push(series[s2] || 0);
      }
      var totalUsers = usersPerDevice.reduce(function(a,b){return a+b;}, 0);
      for (var i2 = 0; i2 < usersPerDevice.length; i2++) {
        var pct = (totalUsers > 0) ? usersPerDevice[i2] / totalUsers : 0;
        row.push(pct);
      }
      for (var g2 = 0; g2 < SEGMENTS.length; g2++) {
        var seg2 = SEGMENTS[g2];
        var series = countsBySeg[seg2.key];
        var users = series[s2] || 0;
        var drop = '—';
        if (s2 === 1) drop = dropPct(series[1], series[0]);
        if (s2 === 2) drop = dropPct(series[2], series[1]);
        if (s2 === 4) drop = dropPct(series[4], series[3]);
        if (s2 === 5) drop = dropPct(series[5], series[4]);
        row.push(users, drop);
      }
      weekRows.push(row);
    }

    // Write/update rows in sheet
    for (var r = 0; r < weekRows.length; r++) {
      var row = weekRows[r];
      var key = row[0] + '|' + row[1];
      if (weekStageToRowIdx[key]) {
        // Update existing row
        sh.getRange(weekStageToRowIdx[key], 1, 1, row.length).setValues([row]);
      } else {
        // Append new row
        sh.appendRow(row);
        // Update mapping for future checks in this run
        weekStageToRowIdx[key] = sh.getLastRow();
      }
    }
    Logger.log('Backfilled Device & Platform Pivot for week: ' + weekLabel + ' (' + weekRows.length + ' rows)');
    // Advance to next week
    currStart.setUTCDate(currStart.getUTCDate() + 7);
  }
}

/************ GA4 metadata helpers (cache) ************/
var __DP_META_CACHE = null;
function DP_getMetadata_(propertyName) {
  if (__DP_META_CACHE) return __DP_META_CACHE;
  try {
    __DP_META_CACHE = AnalyticsData.Properties.getMetadata(propertyName);
  } catch (e) {
    __DP_META_CACHE = { dimensions: [], metrics: [] };
  }
  return __DP_META_CACHE;
}
function DP_hasDimension_(propertyName, dimName) {
  var meta = DP_getMetadata_(propertyName);
  if (!meta || !meta.dimensions) return false;
  for (var i = 0; i < meta.dimensions.length; i++) {
    if (meta.dimensions[i].apiName === dimName) return true;
  }
  return false;
}

/**
 * Resolve WebView custom dimension via Analytics Admin API (when Data API metadata hasn’t propagated).
 * Returns apiName like 'customEvent:is_webview' or null.
 */
function DP_resolveWVDim_viaAdmin_(propertyName) {
  try {
    var pid = (propertyName || '').replace('properties/','');
    if (!pid) return null;
    var parent = 'properties/' + pid;
    var list = AnalyticsAdmin.Properties.CustomDimensions.list(parent);
    var cds = (list && list.customDimensions) ? list.customDimensions : [];
    // Prefer event-scoped `is_webview` by parameterName; else any displayName containing 'webview'
    var pick = null;
    for (var i=0;i<cds.length;i++) {
      var cd = cds[i];
      if ((cd.scope === 'EVENT') && (cd.parameterName === 'is_webview')) { pick = {scope:'EVENT', name:cd.parameterName, display:cd.displayName||''}; break; }
    }
    if (!pick) {
      for (var j=0;j<cds.length;j++) {
        var cd2 = cds[j];
        var disp = (cd2.displayName||'').toLowerCase();
        if (disp.indexOf('webview') !== -1) { pick = {scope:cd2.scope, name:cd2.parameterName, display:cd2.displayName||''}; break; }
      }
    }
    if (!pick) return null;
    var apiName = (pick.scope === 'EVENT') ? ('customEvent:' + pick.name) : ('customUser:' + pick.name);
    if (DP_DEBUG) Logger.log('[WVDim/Admin] resolved → ' + apiName + ' (' + pick.display + ')');
    return apiName;
  } catch (e) {
    if (DP_DEBUG) Logger.log('[WVDim/Admin] error: ' + e);
    return null;
  }
}

function DP_resolveWVDim_(propertyName) {
  // 0) Manual override wins
  if (DP_WV_DIM_OVERRIDE && typeof DP_WV_DIM_OVERRIDE === 'string') {
    if (DP_DEBUG) Logger.log('[WVDim] using manual override → ' + DP_WV_DIM_OVERRIDE);
    return DP_WV_DIM_OVERRIDE;
  }

  // 1) Try Data API metadata (fast path)
  var meta = DP_getMetadata_(propertyName);
  if (meta && meta.dimensions) {
    var dims = meta.dimensions;
    // match by displayName
    var hits = [];
    for (var i = 0; i < dims.length; i++) {
      var d = dims[i];
      var disp = (d.displayName || '').toLowerCase();
      if (disp.indexOf('webview') !== -1) hits.push(d);
    }
    if (hits.length) {
      var pick = null;
      for (var h=0; h<hits.length; h++) { if (hits[h].apiName.indexOf('customEvent:') === 0) { pick = hits[h]; break; } }
      if (!pick) { for (var u=0; u<hits.length; u++) { if (hits[u].apiName.indexOf('customUser:') === 0) { pick = hits[u]; break; } } }
      if (!pick) pick = hits[0];
      if (DP_DEBUG) Logger.log('[WVDim] matched by displayName → ' + pick.apiName + ' (' + pick.displayName + ')');
      return pick.apiName;
    }
    // preferred apiNames
    var preferred = ['customEvent:is_webview','customEvent:mobile_webview','customUser:is_webview','customUser:mobile_webview'];
    for (var p=0;p<preferred.length;p++) {
      for (var q=0;q<dims.length;q++) {
        if (dims[q].apiName === preferred[p]) {
          if (DP_DEBUG) Logger.log('[WVDim] matched by apiName (preferred) → ' + preferred[p] + ' (' + (dims[q].displayName||'') + ')');
          return preferred[p];
        }
      }
    }
    for (var r=0;r<dims.length;r++) {
      var api = (dims[r].apiName||'').toLowerCase();
      if (api.indexOf('webview') !== -1) {
        if (DP_DEBUG) Logger.log('[WVDim] matched by apiName (contains webview) → ' + dims[r].apiName + ' (' + (dims[r].displayName||'') + ')');
        return dims[r].apiName;
      }
    }
  }

  // 2) Try Analytics Admin API (often updates sooner than Data API metadata)
  var viaAdmin = DP_resolveWVDim_viaAdmin_(propertyName);
  if (viaAdmin) return viaAdmin;

  if (DP_DEBUG) Logger.log('[WVDim] no custom Webview dimension found in metadata or Admin API');
  return null;
}

function DP_clearMetaCache_() {
  __DP_META_CACHE = null;
}

/************ App share via `is_webview` event (fallback when no is_app param) ************/
var __DP_APP_SHARE_CACHE = {};
function DP_getAppShare_(propertyName, startDate, endDate, osName) {
  var key = [startDate, endDate, osName].join('|');
  if (__DP_APP_SHARE_CACHE[key] !== undefined) return __DP_APP_SHARE_CACHE[key];

  var WV_DIM = DP_resolveWVDim_(propertyName);
  if (!WV_DIM) {
    if (DP_DEBUG) Logger.log('[AppShare] ABORT: Webview dimension not resolvable yet. Register is_webview (Event scope) or wait for API propagation.');
    __DP_APP_SHARE_CACHE[key] = 0; // explicit zero to keep sheet deterministic
    return 0;
  }

  var code = (osName === 'iOS') ? '1' : '2';
  if (DP_DEBUG) Logger.log('[AppShare] Using ' + WV_DIM + ' (code=' + code + ') for ' + osName);

  var numResp = AnalyticsData.Properties.runReport({
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    metrics: [{ name: 'totalUsers' }],
    dimensionFilter: { andGroup: { expressions: [
      { filter: { fieldName: WV_DIM, stringFilter: { value: code, matchType: 'EXACT' } } }
    ] } }
  }, propertyName);
  var num = DP_readSingleMetric_(numResp, 0);

  var denResp = AnalyticsData.Properties.runReport({
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    metrics: [{ name: 'totalUsers' }],
    dimensionFilter: { andGroup: { expressions: [
      { filter: { fieldName: 'eventName', stringFilter: { value: 'session_start', matchType: 'EXACT' } } },
      { filter: { fieldName: 'platform', stringFilter: { value: 'web', matchType: 'EXACT' } } },
      { filter: { fieldName: 'operatingSystem', stringFilter: { value: osName, matchType: 'EXACT' } } }
    ] } }
  }, propertyName);
  var den = DP_readSingleMetric_(denResp, 0);

  var share = (den > 0) ? Math.max(0, Math.min(1, num/den)) : 0;
  if (DP_DEBUG) Logger.log('[AppShare] OS=' + osName + ' ' + startDate + '…' + endDate + ' num=' + num + ' den=' + den + ' share=' + share.toFixed(4));
  __DP_APP_SHARE_CACHE[key] = share;
  return share;
}

/************ Minimal metric reader ************/

function DP_readSingleMetric_(resp, metricIndex) {
  if (!resp || !resp.rows || !resp.rows.length) return 0;
  var mv = resp.rows[0].metricValues;
  if (!mv || mv.length <= metricIndex) return 0;
  return Number(mv[metricIndex].value || 0);
}

// Builds a NOT string-equals filter for GA4 Data API (wraps in notExpression)
function DP_notStringEq_(fieldName, value) {
  return { notExpression: { filter: { fieldName: fieldName, stringFilter: { value: value, matchType: 'EXACT' } } } };
}

/************ Segment-aware event reader ************/
/**
 * Reads GA4 totalUsers for an event (and optional p2) within a device/platform segment.
 * Segments:
 *   'desktop-web'  => platform='web' AND deviceCategory='desktop'
 *   'ios-web'      => platform='web' AND operatingSystem='iOS'
 *   'ios-app'      => platform='iOS' OR (platform='web' AND operatingSystem='iOS' AND is_app=true*)
 *   'android-web'  => platform='web' AND operatingSystem='Android'
 *   'android-app'  => platform='Android' OR (platform='web' AND operatingSystem='Android' AND is_app=true*)
 * The is_app flag may exist as customUser:is_app or customEvent:is_app. If neither is registered, we omit it.
 * If no is_app dimension is available, estimates app counts for ios-app and android-app by multiplying
 * web counts by share computed from `is_webview` event.
 */
function DP_readUsersForEventBySegment_(propertyName, startDate, endDate, eventName, p2Value, segmentKey) {
  var base = [{ filter: { fieldName: 'eventName', stringFilter: { value: eventName, matchType: 'EXACT' } } }];
  if (p2Value !== null && p2Value !== undefined) {
    base.push({ filter: { fieldName: 'customEvent:p2', stringFilter: { value: p2Value, matchType: 'EXACT' } } });
  }
  function andExpr(exprs) { return { andGroup: { expressions: exprs } }; }
  function eq(field, value) { return { filter: { fieldName: field, stringFilter: { value: value, matchType: 'EXACT' } } }; }

  // Prefer explicit is_app dimensions if they exist
  var hasUserIsApp  = DP_hasDimension_(propertyName, 'customUser:is_app');
  var hasEventIsApp = DP_hasDimension_(propertyName, 'customEvent:is_app');
  var IS_APP_DIM = hasUserIsApp ? 'customUser:is_app' : (hasEventIsApp ? 'customEvent:is_app' : null);

  // Also try to resolve the WebView custom dimension (is_webview) if available
  var WV_DIM = DP_resolveWVDim_(propertyName); // e.g., customEvent:is_webview

  if (DP_DEBUG) {
    Logger.log('[SegRead] event=' + eventName + (p2Value!=null?(' p2='+p2Value):'') + ' seg=' + segmentKey + ' isAppDim=' + IS_APP_DIM);
  }

  // Desktop Web
  if (segmentKey === 'desktop-web') {
    var respD = AnalyticsData.Properties.runReport({
      dateRanges: [{ startDate: startDate, endDate: endDate }],
      metrics: [{ name: 'totalUsers' }],
      dimensionFilter: { andGroup: { expressions: base.concat([ eq('platform','web'), eq('deviceCategory','desktop') ]) } }
    }, propertyName);
    var vD = DP_readSingleMetric_(respD, 0);
    if (DP_DEBUG) Logger.log('[SegRead] seg=desktop-web value=' + vD);
    return vD;
  }

  // iOS Web (exclude iOS WebView if WV_DIM available)
  if (segmentKey === 'ios-web') {
    var exprIW = base.concat([ eq('platform','web'), eq('operatingSystem','iOS') ]);
    if (WV_DIM) exprIW.push(DP_notStringEq_(WV_DIM, '1'));
    var respIW = AnalyticsData.Properties.runReport({
      dateRanges: [{ startDate: startDate, endDate: endDate }],
      metrics: [{ name: 'totalUsers' }],
      dimensionFilter: { andGroup: { expressions: exprIW } }
    }, propertyName);
    var vIW = DP_readSingleMetric_(respIW, 0);
    if (DP_DEBUG) Logger.log('[SegRead] seg=ios-web value=' + vIW + (WV_DIM? ' (excluding WV=1 via '+WV_DIM+')' : ''));
    return vIW;
  }

  // Android Web (exclude Android WebView if WV_DIM available)
  if (segmentKey === 'android-web') {
    var exprAW = base.concat([ eq('platform','web'), eq('operatingSystem','Android') ]);
    if (WV_DIM) exprAW.push(DP_notStringEq_(WV_DIM, '2'));
    var respAW = AnalyticsData.Properties.runReport({
      dateRanges: [{ startDate: startDate, endDate: endDate }],
      metrics: [{ name: 'totalUsers' }],
      dimensionFilter: { andGroup: { expressions: exprAW } }
    }, propertyName);
    var vAW = DP_readSingleMetric_(respAW, 0);
    if (DP_DEBUG) Logger.log('[SegRead] seg=android-web value=' + vAW + (WV_DIM? ' (excluding WV=2 via '+WV_DIM+')' : ''));
    return vAW;
  }

  // iOS App
  if (segmentKey === 'ios-app') {
    // Highest priority: if GA4 exposes is_webview as a custom dimension, count iOS WebView directly
    if (WV_DIM) {
      var respIAwv = AnalyticsData.Properties.runReport({
        dateRanges: [{ startDate: startDate, endDate: endDate }],
        metrics: [{ name: 'totalUsers' }],
        dimensionFilter: { andGroup: { expressions: base.concat([
          eq('platform','web'),
          eq('operatingSystem','iOS'),
          { filter: { fieldName: WV_DIM, stringFilter: { value: '1', matchType: 'EXACT' } } }
        ]) } }
      }, propertyName);
      var vIAwv = DP_readSingleMetric_(respIAwv, 0);

      // Include any native iOS app users (if present)
      var iosNativeResp0 = AnalyticsData.Properties.runReport({
        dateRanges: [{ startDate: startDate, endDate: endDate }],
        metrics: [{ name: 'totalUsers' }],
        dimensionFilter: { andGroup: { expressions: base.concat([ eq('platform','iOS') ]) } }
      }, propertyName);
      var iosNative0 = DP_readSingleMetric_(iosNativeResp0, 0);

      var totalIA = vIAwv + iosNative0;
      if (DP_DEBUG) Logger.log('[SegRead] seg=ios-app(webview+native via ' + WV_DIM + ') webview=' + vIAwv + ' native=' + iosNative0 + ' total=' + totalIA);
      return totalIA;
    }
    if (IS_APP_DIM) {
      // Native iOS OR WebView flagged as is_app
      var respIA = AnalyticsData.Properties.runReport({
        dateRanges: [{ startDate: startDate, endDate: endDate }],
        metrics: [{ name: 'totalUsers' }],
        dimensionFilter: { orGroup: { expressions: [
          andExpr(base.concat([ eq('platform','iOS') ])),
          andExpr(base.concat([ eq('platform','web'), eq('operatingSystem','iOS'), eq(IS_APP_DIM,'true') ]))
        ] } }
      }, propertyName);
      var vIA = DP_readSingleMetric_(respIA, 0);
      if (DP_DEBUG) Logger.log('[SegRead] seg=ios-app(native+flag) value=' + vIA);
      return vIA;
    }
    // Fallback: include native iOS app users (platform='iOS')
    // plus an estimate of iOS WebView = (iOS-web users) × appShare(iOS)
    var iosNativeResp = AnalyticsData.Properties.runReport({
      dateRanges: [{ startDate: startDate, endDate: endDate }],
      metrics: [{ name: 'totalUsers' }],
      dimensionFilter: { andGroup: { expressions: base.concat([ eq('platform','iOS') ]) } }
    }, propertyName);
    var iosNativeUsers = DP_readSingleMetric_(iosNativeResp, 0);

    var iosWebResp = AnalyticsData.Properties.runReport({
      dateRanges: [{ startDate: startDate, endDate: endDate }],
      metrics: [{ name: 'totalUsers' }],
      dimensionFilter: { andGroup: { expressions: base.concat([ eq('platform','web'), eq('operatingSystem','iOS') ]) } }
    }, propertyName);
    var iosWebUsers = DP_readSingleMetric_(iosWebResp, 0);

    var shareIOS = DP_getAppShare_(propertyName, startDate, endDate, 'iOS');
    var estWV = Math.round(iosWebUsers * shareIOS);
    var estTotalIOSApp = iosNativeUsers + estWV;
    if (DP_DEBUG) Logger.log('[SegRead] seg=ios-app(estimate native+iOS webview) native=' + iosNativeUsers + ' web=' + iosWebUsers + ' share=' + shareIOS.toFixed(4) + ' total=' + estTotalIOSApp);
    return estTotalIOSApp;
  }

  // Android App
  if (segmentKey === 'android-app') {
    // Highest priority: if GA4 exposes is_webview as a custom dimension, count Android WebView directly
    if (WV_DIM) {
      var respAAwv = AnalyticsData.Properties.runReport({
        dateRanges: [{ startDate: startDate, endDate: endDate }],
        metrics: [{ name: 'totalUsers' }],
        dimensionFilter: { andGroup: { expressions: base.concat([
          eq('platform','web'),
          eq('operatingSystem','Android'),
          { filter: { fieldName: WV_DIM, stringFilter: { value: '2', matchType: 'EXACT' } } }
        ]) } }
      }, propertyName);
      var vAAwv = DP_readSingleMetric_(respAAwv, 0);

      // Include any native Android app users (if present)
      var andNativeResp0 = AnalyticsData.Properties.runReport({
        dateRanges: [{ startDate: startDate, endDate: endDate }],
        metrics: [{ name: 'totalUsers' }],
        dimensionFilter: { andGroup: { expressions: base.concat([ eq('platform','Android') ]) } }
      }, propertyName);
      var andNative0 = DP_readSingleMetric_(andNativeResp0, 0);

      var totalAA = vAAwv + andNative0;
      if (DP_DEBUG) Logger.log('[SegRead] seg=android-app(webview+native via ' + WV_DIM + ') webview=' + vAAwv + ' native=' + andNative0 + ' total=' + totalAA);
      return totalAA;
    }
    if (IS_APP_DIM) {
      var respAA = AnalyticsData.Properties.runReport({
        dateRanges: [{ startDate: startDate, endDate: endDate }],
        metrics: [{ name: 'totalUsers' }],
        dimensionFilter: { orGroup: { expressions: [
          andExpr(base.concat([ eq('platform','Android') ])),
          andExpr(base.concat([ eq('platform','web'), eq('operatingSystem','Android'), eq(IS_APP_DIM,'true') ]))
        ] } }
      }, propertyName);
      var vAA = DP_readSingleMetric_(respAA, 0);
      if (DP_DEBUG) Logger.log('[SegRead] seg=android-app(native+flag) value=' + vAA);
      return vAA;
    }
    // Fallback: include native Android app users (platform='Android')
    // plus an estimate of Android WebView = (Android-web users) × appShare(Android)
    var andNativeResp = AnalyticsData.Properties.runReport({
      dateRanges: [{ startDate: startDate, endDate: endDate }],
      metrics: [{ name: 'totalUsers' }],
      dimensionFilter: { andGroup: { expressions: base.concat([ eq('platform','Android') ]) } }
    }, propertyName);
    var andNativeUsers = DP_readSingleMetric_(andNativeResp, 0);

    var andWebResp = AnalyticsData.Properties.runReport({
      dateRanges: [{ startDate: startDate, endDate: endDate }],
      metrics: [{ name: 'totalUsers' }],
      dimensionFilter: { andGroup: { expressions: base.concat([ eq('platform','web'), eq('operatingSystem','Android') ]) } }
    }, propertyName);
    var andWebUsers = DP_readSingleMetric_(andWebResp, 0);

    var shareAND = DP_getAppShare_(propertyName, startDate, endDate, 'Android');
    var estWVand = Math.round(andWebUsers * shareAND);
    var estTotalANDApp = andNativeUsers + estWVand;
    if (DP_DEBUG) Logger.log('[SegRead] seg=android-app(estimate native+Android webview) native=' + andNativeUsers + ' web=' + andWebUsers + ' share=' + shareAND.toFixed(4) + ' total=' + estTotalANDApp);
    return estTotalANDApp;
  }

  // Default guard
  return 0;
}

function DP_debugDevicePlatformBlock() {
  var PROPERTY_ID = '418611571';
  var propertyName = 'properties/' + PROPERTY_ID;
  var curr = DP_getLastFullWeek();
  Logger.log('=== Device & Platform Debug for ' + curr.startDate + '…' + curr.endDate + ' ===');

  var segs = ['desktop-web','ios-web','ios-app','android-web','android-app'];
  var stages = [
    { k:'inqTC',  label:'Send Inquiry Trip Cart',  e:'trip-cart_price-calculated', p2:'false' },
    { k:'inqS',   label:'Inquiry Start',           e:'inquiry_start' },
    { k:'inqSub', label:'Inquiry Submitted',       e:'inquiry_submit_success' },
    { k:'bnTC',   label:'Book Now Trip Cart',      e:'trip-cart_price-calculated', p2:'true' },
    { k:'bnClk',  label:'Book Now Clicks',         e:'trip-cart_book-now-click' },
    { k:'bnPay',  label:'Proceed to Payment',      e:'trip-cart_book-now-proceed-to-payment-cl' }
  ];

  for (var si=0; si<segs.length; si++) {
    var seg = segs[si];
    var line = [seg];
    for (var st=0; st<stages.length; st++) {
      var stg = stages[st];
      var val = DP_readUsersForEventBySegment_(propertyName, curr.startDate, curr.endDate, stg.e, stg.p2, seg);
      line.push(val);
    }
    Logger.log('[Block] ' + line.join(' | '));
  }

  // Also print app shares used
  try {
    var sIOS = DP_getAppShare_(propertyName, curr.startDate, curr.endDate, 'iOS');
    var sAND = DP_getAppShare_(propertyName, curr.startDate, curr.endDate, 'Android');
    Logger.log('[AppShareSummary] iOS=' + (sIOS*100).toFixed(2) + '% Android=' + (sAND*100).toFixed(2) + '%');
  } catch (e) {
    Logger.log('[AppShareSummary] error: ' + e);
  }
}

/**
 * Debug ONLY the `inquiry_start` stage across segments for the last full week.
 * Prints per-segment users and app shares used.
 */
function DP_debugInquiryStartOnly() {
  var PROPERTY_ID = '418611571';
  var propertyName = 'properties/' + PROPERTY_ID;
  var curr = DP_getLastFullWeek();
  Logger.log('=== Inquiry Start Debug for ' + curr.startDate + '…' + curr.endDate + ' ===');

  // Show which WV custom dimension the API exposes, if any
  var wvDim = DP_resolveWVDim_(propertyName);
  Logger.log('[WVDim] resolved = ' + wvDim);

  // Compute app shares that will be used in fallback
  var sIOS = 0, sAND = 0;
  try { sIOS = DP_getAppShare_(propertyName, curr.startDate, curr.endDate, 'iOS'); } catch (e) { Logger.log('[AppShare] iOS error: ' + e); }
  try { sAND = DP_getAppShare_(propertyName, curr.startDate, curr.endDate, 'Android'); } catch (e) { Logger.log('[AppShare] Android error: ' + e); }
  Logger.log('[AppShareSummary] iOS=' + (sIOS*100).toFixed(2) + '% Android=' + (sAND*100).toFixed(2) + '%');

  var segs = ['desktop-web','ios-web','ios-app','android-web','android-app'];
  for (var i=0; i<segs.length; i++) {
    var seg = segs[i];
    var val = DP_readUsersForEventBySegment_(propertyName, curr.startDate, curr.endDate, 'inquiry_start', null, seg);
    Logger.log('[InquiryStart] ' + seg + ' = ' + val);
  }
}

/**
 * List GA4 custom dimensions visible to the API (quick sanity check).
 * Helpful to verify if 'mobile_webview' is registered and exposed.
 */
function DP_debugListCustomDims() {
  var PROPERTY_ID = '418611571';
  var meta = DP_getMetadata_('properties/' + PROPERTY_ID); // use cached helper (avoids transient 404s)
  if (!meta || !meta.dimensions) {
    Logger.log('[Meta] No dimensions returned by API (check permissions/property ID or try again).');
    return;
  }
  Logger.log('--- Event-scoped custom dimensions ---');
  (meta.dimensions || []).forEach(function(d) {
    if (d.apiName && d.apiName.indexOf('customEvent:') === 0) {
      Logger.log(d.apiName + ' — ' + (d.displayName || ''));
    }
  });
  Logger.log('--- User-scoped custom dimensions ---');
  (meta.dimensions || []).forEach(function(d) {
    if (d.apiName && d.apiName.indexOf('customUser:') === 0) {
      Logger.log(d.apiName + ' — ' + (d.displayName || ''));
    }
  });
}

function DP_debugFindWebviewDim() {
  var PROPERTY_ID = '418611571';
  DP_clearMetaCache_(); // force refresh
  var meta = DP_getMetadata_('properties/' + PROPERTY_ID);
  if (!meta || !meta.dimensions) { Logger.log('[Meta] no dimensions'); return; }
  var any = false;
  meta.dimensions.forEach(function(d){
    var disp = (d.displayName || '');
    var api  = (d.apiName || '');
    if (disp.toLowerCase().indexOf('webview') !== -1 || api.toLowerCase().indexOf('webview') !== -1) {
      Logger.log('[WV] ' + api + ' — ' + disp);
      any = true;
    }
  });
  if (!any) Logger.log('[WV] none found in metadata');
}