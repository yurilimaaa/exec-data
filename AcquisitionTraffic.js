/* 
test change 
upload change on feature-branch on git 
*/
function updateAcqTrafficChannelsWoWYoY() {
  var PROPERTY_ID = '418611571'; // GA4 property
  var SPREADSHEET_ID = '1VxsGju5iO2WUInUJiy8zyUfdQ5ZyObN3NirxcAvQs94';
  var TAB_NAME = 'Acquisition & Traffic';
  var propertyName = 'properties/' + PROPERTY_ID;

  // Executive roll-up groups
  var channelMap = {
    'Organic':        ['Organic Search', 'Direct', 'Organic Shopping', 'Organic Video'],
    'Paid':           ['Paid Search', 'Paid Social', 'Paid Other', 'Cross-network'],
    'Organic Social': ['Organic Social'],
    'Email':          ['Email'],
    'Referral/Other': ['Referral', 'Unassigned']
  };

  // Periods
  var curr = getLastFullWeek();          // last full Sun–Sat
  var prev = shiftPeriodDays_(curr, -7); // week -1
  var yoy  = shiftPeriodDays_(curr, -364); // ~same week last year

  // Pull raw channel reports
  var currMap = runChannelReport_(propertyName, curr.startDate, curr.endDate); // {channel: {sessions,activeUsers,engagedSessions,engagementRate}}
  var prevMap = runChannelReport_(propertyName, prev.startDate, prev.endDate);
  var yoyMap  = runChannelReport_(propertyName, yoy.startDate,  yoy.endDate);

  // Sum to groups
  var groupCurr = rollupToGroups_(channelMap, currMap);
  var groupPrev = rollupToGroups_(channelMap, prevMap);
  var groupYoY  = rollupToGroups_(channelMap, yoyMap);

  // Totals for share calc
  var totals = totalsFromMap_(groupCurr); // uses same structure: {sessions,activeUsers,engagedSessions}
  var totalSessionsCurr = Number(totals.sessions || 0);

  // Prepare week label
  var weekStr = formatDMY_(curr.startDate) + ' / ' + formatDMY_(curr.endDate);

  // Output header
  var headers = [
    'Week',
    'Group',
    'Sessions',
    'Sessions WoW %',
    'Sessions YoY %',
    'New Users',
    'New Users WoW %',
    'New Users YoY %',
    'Engaged Sessions',
    'Engagement Rate',
    'Group % of Sessions'
  ];

  // Stable group order for execs
  var groupOrder = ['Organic', 'Paid', 'Organic Social', 'Email', 'Referral/Other'];

  // Build rows
  var rows = [];
  for (var gi = 0; gi < groupOrder.length; gi++) {
    var g = groupOrder[gi];
    var c = groupCurr[g] || { sessions:0, activeUsers:0, engagedSessions:0, engagementRate:0 };
    var p = groupPrev[g] || { sessions:0, activeUsers:0 };
    var y = groupYoY[g]  || { sessions:0, activeUsers:0 };

    var sessions = num_(c.sessions);
    var activeUsers = num_(c.activeUsers);
    var engaged  = num_(c.engagedSessions);

    // Recompute engagement rate from counts (more accurate than averaging rates)
    var erate = sessions ? (engaged / sessions) : 0;

    var sessionsWoW = pctChange_(sessions, num_(p.sessions));
    var sessionsYoY = pctChange_(sessions, num_(y.sessions));
    var activeUsersWoW = pctChange_(activeUsers, num_(p.activeUsers));
    var activeUsersYoY = pctChange_(activeUsers, num_(y.activeUsers));

    var share = ratio_(sessions, totalSessionsCurr);

    rows.push([
      weekStr,
      g,
      sessions,
      pctStr_(sessionsWoW),
      pctStr_(sessionsYoY),
      activeUsers,
      pctStr_(activeUsersWoW),
      pctStr_(activeUsersYoY),
      engaged,
      pctStr_(erate),
      pctStr_(share)
    ]);
  }

  // Write to sheet (replace any existing rows for this Week label)
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME) || ss.insertSheet(TAB_NAME);

  var needHeader = needsHeader_(sheet, headers.length, headers);
  if (needHeader) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Preserve history: if this week already exists, do NOT overwrite or delete — just skip writing.
  if (weekExists_(sheet, weekStr)) {
    Logger.log('Week ' + weekStr + ' already present; skipping write to preserve history.');
    return;
  }

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }

  Logger.log('Grouped Acquisition written for ' + weekStr + ' (' + rows.length + ' groups)');
}


function rollupToGroups_(channelMap, inMap) {
  var out = {}; // group -> metrics
  // Initialize groups
  for (var g in channelMap) {
    out[g] = { sessions:0, activeUsers:0, engagedSessions:0, engagementRate:0 };
  }
  // Accumulate
  for (var g2 in channelMap) {
    var list = channelMap[g2];
    for (var i = 0; i < list.length; i++) {
      var ch = list[i];
      var src = inMap[ch];
      if (!src) continue;
      out[g2].sessions        += Number(src.sessions || 0);
      out[g2].activeUsers        += Number(src.activeUsers || 0);
      out[g2].engagedSessions += Number(src.engagedSessions || 0);
    }
    // compute engagementRate for the group from counts
    var s = out[g2].sessions;
    out[g2].engagementRate = s ? (out[g2].engagedSessions / s) : 0;
  }
  return out;
}

/* ---------------- GA4 + Transform Helpers ---------------- */

function runChannelReport_(propertyName, startDate, endDate) {
  var req = {
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    dimensions: [{ name: 'sessionDefaultChannelGroup' }],
    metrics: [
      { name: 'sessions' },
      { name: 'activeUsers' },
      { name: 'engagedSessions' },
      { name: 'engagementRate' }
    ],
    limit: 1000
  };
  var resp = AnalyticsData.Properties.runReport(req, propertyName);
  var map = {}; // channel -> metrics
  var rows = (resp && resp.rows) ? resp.rows : [];
  for (var i = 0; i < rows.length; i++) {
    var ch = rows[i].dimensionValues[0].value || '(not set)';
    map[ch] = {
      sessions: Number(rows[i].metricValues[0].value || 0),
      activeUsers: Number(rows[i].metricValues[1].value || 0),
      engagedSessions: Number(rows[i].metricValues[2].value || 0),
      engagementRate: Number(rows[i].metricValues[3].value || 0) // 0..1
    };
  }
  return map;
}

function totalsFromMap_(map) {
  var t = { sessions:0, activeUsers:0, engagedSessions:0 };
  for (var k in map) {
    t.sessions += Number(map[k].sessions || 0);
    t.activeUsers += Number(map[k].activeUsers || 0);
    t.engagedSessions += Number(map[k].engagedSessions || 0);
  }
  return t;
}

/* ---------------- Generic Helpers (Rhino-safe) ---------------- */

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

function shiftPeriodDays_(period, deltaDays) {
  // Returns a new {startDate,endDate} by shifting both dates by deltaDays
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
  var parts = s.split('-');
  return new Date(Date.UTC(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2])));
}

function formatDMY_(ymd) {
  var parts = ymd.split('-'); // yyyy-mm-dd
  return parts[2] + '-' + parts[1] + '-' + parts[0];
}

function num_(v) {
  return Number(v || 0);
}

function ratio_(part, total) {
  part = Number(part || 0);
  total = Number(total || 0);
  if (!total) return 0;
  return part / total;
}

function pctChange_(curr, prev) {
  curr = Number(curr || 0);
  prev = Number(prev || 0);
  if (!prev) return null; // undefined % change (avoid div by zero)
  return (curr - prev) / prev; // ratio, e.g., 0.12 => +12%
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

function deleteRowsForWeekKey_(sheet, weekKey, width) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var values = sheet.getRange(2, 1, last - 1, width).getValues();
  var rowsToDelete = [];
  for (var r = 0; r < values.length; r++) {
    if (String(values[r][0]) === weekKey) {
      rowsToDelete.push(r + 2); // actual row index
    }
  }
  // Delete from bottom to top
  for (var i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
}

function weekExists_(sheet, weekKey) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var values = sheet.getRange(2, 1, last - 1, 1).getValues(); // column A only
  for (var r = 0; r < values.length; r++) {
    if (String(values[r][0]) === weekKey) return true;
  }
  return false;
}
