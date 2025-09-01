/***** GA4 Event ↔ Parameter Inventory (7 days, Rhino-safe, with counts)
 * Produces two tabs:
 * 1) "Event Inventory"  : Event Name | Event Count (7 days) | Parameters Used (registered, last 7 days)
 * 2) "Param Coverage"   : Parameter | Events Using It (last 7 days) | #Events | Total Uses (7 days)
 ********************************************************/

function buildEventParameterInventory7d() {
  var PROPERTY_ID = '418611571';
  var SPREADSHEET_ID = '1VxsGju5iO2WUInUJiy8zyUfdQ5ZyObN3NirxcAvQs94';
  var TAB_EVENTS = 'Event Inventory';
  var TAB_PARAMS = 'Param Coverage';

  // Date range: last 7 days
  var today = new Date();
  var endDate = today.toISOString().slice(0, 10);
  var start = new Date();
  start.setDate(start.getDate() - 7);
  var startDate = start.toISOString().slice(0, 10);

  var propertyName = 'properties/' + PROPERTY_ID;

  // 1) Get all active events + their eventCount (7d)
  var eventsResp = AnalyticsData.Properties.runReport({
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    dimensions: [{ name: 'eventName' }],
    metrics: [{ name: 'eventCount' }],
    limit: 10000
  }, propertyName);

  var activeEvents = [];                 // array of event names (for stable ordering)
  var eventCounts = {};                  // eventName -> total eventCount (number)
  var rows1 = (eventsResp.rows || []);
  for (var i = 0; i < rows1.length; i++) {
    var evName = rows1[i].dimensionValues[0].value;
    var cnt = Number(rows1[i].metricValues[0].value || 0);
    activeEvents.push(evName);
    eventCounts[evName] = cnt;
  }

  // 2) Get registered custom event parameters (customEvent:*)
  var meta = AnalyticsData.Properties.getMetadata(propertyName + '/metadata');
  var customEventDims = (meta.dimensions || []).filter(function (d) {
    var name = String(d.apiName || '');
    return name.indexOf('customEvent:') === 0;
  });

  Logger.log('Found ' + customEventDims.length + ' registered custom event parameters.');

  // 3) Build:
  //    - paramToEvents: param -> [unique events using it]
  //    - paramTotalUses: param -> sum of eventCount where param was set
  //    - eventToParams: event -> [unique params seen on that event]
  var paramToEvents = {};     // { param: [event, ...] }
  var paramTotalUses = {};    // { param: number }
  var eventToParams = {};     // { event: [param, ...] }

  // Initialize eventToParams so every event appears
  for (var e = 0; e < activeEvents.length; e++) {
    eventToParams[activeEvents[e]] = [];
  }

  for (var d = 0; d < customEventDims.length; d++) {
    var apiName = customEventDims[d].apiName;       // e.g. "customEvent:trip_id"
    var shortName = apiName.split(':')[1];          // "trip_id"
    Logger.log('Scanning parameter: ' + shortName);

    try {
      var resp = AnalyticsData.Properties.runReport({
        dateRanges: [{ startDate: startDate, endDate: endDate }],
        dimensions: [
          { name: 'eventName' },
          { name: apiName } // customEvent:<param>
        ],
        metrics: [{ name: 'eventCount' }],
        limit: 10000
      }, propertyName);

      var rows = resp.rows || [];
      for (var r = 0; r < rows.length; r++) {
        var ev = rows[r].dimensionValues[0].value;
        var paramVal = rows[r].dimensionValues[1].value;
        var cnt2 = Number(rows[r].metricValues[0].value || 0);

        // Only count when parameter is actually set (exclude "(not set)")
        if (paramVal && paramVal !== '(not set)') {
          // param -> events (unique)
          if (!paramToEvents[shortName]) paramToEvents[shortName] = [];
          if (paramToEvents[shortName].indexOf(ev) === -1) {
            paramToEvents[shortName].push(ev);
          }

          // param -> total uses
          if (!paramTotalUses[shortName]) paramTotalUses[shortName] = 0;
          paramTotalUses[shortName] += cnt2;

          // event -> params (unique)
          if (!eventToParams[ev]) eventToParams[ev] = [];
          if (eventToParams[ev].indexOf(shortName) === -1) {
            eventToParams[ev].push(shortName);
          }
        }
      }
    } catch (err) {
      Logger.log('Error fetching param ' + shortName + ': ' + err);
    }
  }

  // 4) Write Event Inventory: Event | Event Count (7d) | Parameters Used
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetEvents = ss.getSheetByName(TAB_EVENTS) || ss.insertSheet(TAB_EVENTS);
  sheetEvents.clear();

  var eventsHeader = ['Event Name', 'Event Count (7 days)', 'Parameters Used (registered, last 7 days)'];
  sheetEvents.getRange(1, 1, 1, eventsHeader.length).setValues([eventsHeader]);

  var eventRows = [];
  for (var i2 = 0; i2 < activeEvents.length; i2++) {
    var evn = activeEvents[i2];
    var params = (eventToParams[evn] || []).slice().sort();
    eventRows.push([evn, eventCounts[evn] || 0, params.length ? params.join(', ') : '—']);
  }
  if (eventRows.length) sheetEvents.getRange(2, 1, eventRows.length, eventsHeader.length).setValues(eventRows);
  sheetEvents.autoResizeColumns(1, eventsHeader.length);

  // 5) Write Param Coverage: Parameter | Events Using It | #Events | Total Uses (7d)
  var sheetParams = ss.getSheetByName(TAB_PARAMS) || ss.insertSheet(TAB_PARAMS);
  sheetParams.clear();

  var paramsHeader = ['Parameter (registered)', 'Events Using It (last 7 days)', '#Events', 'Total Uses (7 days)'];
  sheetParams.getRange(1, 1, 1, paramsHeader.length).setValues([paramsHeader]);

  var allParams = Object.keys(paramToEvents).sort();
  var paramRows = [];
  for (var j = 0; j < allParams.length; j++) {
    var param = allParams[j];
    var evList = (paramToEvents[param] || []).slice().sort();
    var totalUses = Number(paramTotalUses[param] || 0);
    paramRows.push([param, evList.join(', '), evList.length, totalUses]);
  }
  if (paramRows.length) sheetParams.getRange(2, 1, paramRows.length, paramsHeader.length).setValues(paramRows);
  sheetParams.autoResizeColumns(1, paramsHeader.length);

  Logger.log('Event Inventory updated for ' + activeEvents.length + ' events (with counts).');
}

/***** Drop-Off Analysis — Weekly (Sun–Sat UTC)
 * Tabs created/appended:
 *  1) "Drop-off Pages (Weekly)"
 *     Columns: Week | Page Path | Views | Entrances (Sessions) | Next Page Views | Drop-off %
 *     - Views: screenPageViews by pagePath
 *     - Entrances: sessions by landingPage (GA4-supported proxy)
 *     - Next Page Views: sum of screenPageViews whose pageReferrer equals this Page Path (internal only)
 *       -> Drop-off % = 1 - (Next Page Views / Views)
 *  2) "Path Pairs (Weekly)"
 *     Columns: Week | Referrer | Page Path | Views
 *     - Views: screenPageViews by pageReferrer + pagePath
 *
 * Notes:
 * - Old data is preserved (rows are appended). No column auto-resize.
 * - Week labeled as dd-mm-yyyy / dd-mm-yyyy (UTC).
 * - Advanced Service required: Google Analytics Data API (v1beta).
 * - GA4 API does NOT expose Exits/Exit Rate or isEntrance/isExit in Core Reporting (available in UI/Explorations only).
 *   This script computes a robust drop-off proxy using internal next-page transitions.
 **************************************************************/

function updateDropoffAnalysisWeekly() {
  var PROPERTY_ID     = '418611571';
  var SPREADSHEET_ID  = '1VxsGju5iO2WUInUJiy8zyUfdQ5ZyObN3NirxcAvQs94';

  var TAB_PAGES = 'Drop-off Pages (Weekly)';
  var TAB_PAIRS = 'Path Pairs (Weekly)';

  var INTERNAL_HOST = 'getmyboat.com'; // <-- adjust if needed

  var propertyName = 'properties/' + PROPERTY_ID;
  var period = getLastFullWeek(); // { startDate, endDate } in YYYY-MM-DD
  var weekLabel = formatDMY_(period.startDate) + ' / ' + formatDMY_(period.endDate);

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  /* ---------- 1) PAGE-LEVEL: Views + Entrances + Next Page Views (internal) ---------- */

  // 1a) Views by pagePath
  var pagesResp = AnalyticsData.Properties.runReport({
    dateRanges: [{ startDate: period.startDate, endDate: period.endDate }],
    dimensions: [{ name: 'pagePath' }],
    metrics: [{ name: 'screenPageViews' }],
    limit: 10000
  }, propertyName);

  var viewsMap = {}; // pagePath -> views
  if (pagesResp && pagesResp.rows) {
    for (var i = 0; i < pagesResp.rows.length; i++) {
      var r = pagesResp.rows[i];
      var path = r.dimensionValues[0].value || '/';
      var views = Number(r.metricValues[0].value || 0);
      viewsMap[path] = (viewsMap[path] || 0) + views;
    }
  }

  // 1b) Entrances proxy: sessions by landingPage (GA4-supported)
  var landResp = AnalyticsData.Properties.runReport({
    dateRanges: [{ startDate: period.startDate, endDate: period.endDate }],
    dimensions: [{ name: 'landingPage' }],
    metrics: [{ name: 'sessions' }],
    limit: 10000
  }, propertyName);

  var entrancesMap = {}; // pagePath -> entrance sessions
  if (landResp && landResp.rows) {
    for (var j = 0; j < landResp.rows.length; j++) {
      var lr = landResp.rows[j];
      var lp = lr.dimensionValues[0].value || '/';
      var sess = Number(lr.metricValues[0].value || 0);
      entrancesMap[lp] = (entrancesMap[lp] || 0) + sess;
    }
  }

  /* ---------- 2) PATH PAIRS: pageReferrer -> pagePath ---------- */

  var pairsResp = AnalyticsData.Properties.runReport({
    dateRanges: [{ startDate: period.startDate, endDate: period.endDate }],
    dimensions: [
      { name: 'pageReferrer' },
      { name: 'pagePath' }
    ],
    metrics: [{ name: 'screenPageViews' }],
    limit: 10000
  }, propertyName);

  // Build Next Page Views by referrer path (internal only)
  var nextViewsMap = {}; // pagePath(referrer) -> views to next pages
  if (pairsResp && pairsResp.rows) {
    for (var m = 0; m < pairsResp.rows.length; m++) {
      var rr = pairsResp.rows[m];
      var ref  = rr.dimensionValues[0].value || '';
      var v    = Number(rr.metricValues[0].value || 0);

      // Keep only internal referrers
      var refPath = extractInternalPath_(ref, INTERNAL_HOST);
      if (!refPath) continue; // skip external or empty

      nextViewsMap[refPath] = (nextViewsMap[refPath] || 0) + v;
    }
  }

  // Optional: also write the ref->page pairs for exploration
  var pairsRows = [];
  if (pairsResp && pairsResp.rows) {
    for (var n = 0; n < pairsResp.rows.length; n++) {
      var rr2 = pairsResp.rows[n];
      var ref2  = rr2.dimensionValues[0].value || '(direct/none)';
      var pth2  = rr2.dimensionValues[1].value || '/';
      var vv2   = Number(rr2.metricValues[0].value || 0);
      pairsRows.push([weekLabel, ref2, pth2, vv2]);
    }
    // sort by views desc for readability
    pairsRows.sort(function(a, b){ return b[3] - a[3]; });
  }

  // 3) Merge to Page rows
  var pagesRows = [];
  var allPaths = {};
  for (var p in viewsMap) allPaths[p] = true;
  for (var p2 in entrancesMap) allPaths[p2] = true;
  for (var p3 in nextViewsMap) allPaths[p3] = true;

  for (var pathKey in allPaths) {
    var vTot   = Number(viewsMap[pathKey] || 0);
    var entSes = Number(entrancesMap[pathKey] || 0);
    var nextV  = Number(nextViewsMap[pathKey] || 0);
    var drop   = vTot ? (1 - (nextV / vTot)) : null; // fraction
    pagesRows.push([weekLabel, pathKey, vTot, entSes, nextV, pctStr_(drop)]);
  }
  pagesRows.sort(function(a, b){ return b[2] - a[2]; });

  // 4) Write to Sheets (append, no auto-resize)
  var sheetPages = ss.getSheetByName(TAB_PAGES) || ss.insertSheet(TAB_PAGES);
  var pagesHeaders = ['Week', 'Page Path', 'Views', 'Entrances (Sessions)', 'Next Page Views', 'Drop-off %'];
  if (needsHeader_(sheetPages, pagesHeaders.length, pagesHeaders)) {
    sheetPages.getRange(1,1,1,pagesHeaders.length).setValues([pagesHeaders]).setFontWeight('bold');
  }
  if (pagesRows.length) {
    sheetPages.getRange(sheetPages.getLastRow() + 1, 1, pagesRows.length, pagesHeaders.length).setValues(pagesRows);
  }

  var sheetPairs = ss.getSheetByName(TAB_PAIRS) || ss.insertSheet(TAB_PAIRS);
  var pairsHeaders = ['Week', 'Referrer', 'Page Path', 'Views'];
  if (needsHeader_(sheetPairs, pairsHeaders.length, pairsHeaders)) {
    sheetPairs.getRange(1,1,1,pairsHeaders.length).setValues([pairsHeaders]).setFontWeight('bold');
  }
  if (pairsRows.length) {
    sheetPairs.getRange(sheetPairs.getLastRow() + 1, 1, pairsRows.length, pairsHeaders.length).setValues(pairsRows);
  }

  Logger.log('Drop-off Analysis appended for ' + weekLabel + ' (pages: ' + pagesRows.length + ', pairs: ' + pairsRows.length + ')');
}

/* ------------ Helpers ------------ */

function extractInternalPath_(ref, host) {
  if (!ref) return null;
  // If it's already a path like "/..."
  if (ref.charAt(0) === '/') return ref;
  // Try to match URLs of our own host
  // Examples: https://getmyboat.com/abc, http://www.getmyboat.com/xyz?...
  var lower = String(ref).toLowerCase();
  var h = String(host || '').toLowerCase();
  if (h && lower.indexOf(h) !== -1) {
    var m = lower.match(/^[a-z]+:\/\/[^\/]+(\/.*)$/);
    if (m && m[1]) return m[1];
    // If no path captured but host matches, fallback to '/'
    return '/';
  }
  return null; // external or empty
}

function getLastFullWeek() {
  var today = new Date();
  var utcToday = new Date(today.getTime() + today.getTimezoneOffset() * 60000);
  var day = utcToday.getUTCDay(); // Sunday=0
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

function ymd_(d) {
  var y = d.getUTCFullYear();
  var m = ('0' + (d.getUTCMonth() + 1)).slice(-2);
  var day = ('0' + d.getUTCDate()).slice(-2);
  return y + '-' + m + '-' + day;
}

function formatDMY_(ymd) {
  var p = ymd.split('-'); // yyyy-mm-dd
  return p[2] + '-' + p[1] + '-' + p[0];
}

function pctStr_(ratio) {
  if (ratio === null || typeof ratio === 'undefined') return '—';
  return (Number(ratio) * 100).toFixed(1) + '%';
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
