/**
 * Google Analytics 4 to Google Sheets Automation (AUTO-FETCH VERSION)
 * Automatically fetches property names from GA4
 * Pulls monthly GA4 data for multiple properties into separate sheets
 * Only updates data when a full month is complete
 * Months are displayed as columns
 */

// Configuration - Add your property IDs
const PROPERTY_IDS = [
  /** '326567794',
  '326566302',
  '321438080',
  '499674438',
  '261288477',
  '414341373',
  '417823247',
  '321428119',
  '378485873',
  '350449967',
  '321726839',
  '504309352',
  '366090887',
  '321057392',
  '458349362',
  '493908197',
  '361212512',
  '494174283',
  '321069855',
  '504008380', */
  '309532396'
];

/**
 * Get property display name from GA4
 */
function getPropertyName(propertyId) {
  try {
    const property = AnalyticsAdmin.Properties.get('properties/' + propertyId);
    return property.displayName || propertyId;
  } catch (error) {
    Logger.log(`Error fetching property name for ${propertyId}: ${error.message}`);
    return propertyId;  // Fallback to property ID if name can't be fetched
  }
}

/**
 * Main function to update all properties
 * Run this manually or set up a trigger to run monthly
 */
function updateAllProperties() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  PROPERTY_IDS.forEach(propertyId => {
    try {
      const propertyName = getPropertyName(propertyId);
      updatePropertyData(spreadsheet, propertyId, propertyName);
    } catch (error) {
      Logger.log(`Error updating ${propertyId}: ${error.message}`);
    }
  });
}

/**
 * Update data for a single property
 */
function updatePropertyData(spreadsheet, propertyId, propertyName) {
  Logger.log(`\n========================================`);
  Logger.log(`Processing: ${propertyName} (${propertyId})`);
  Logger.log(`========================================`);
  
  let sheet = spreadsheet.getSheetByName(propertyName);
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    Logger.log(`Creating new sheet: ${propertyName}`);
    sheet = spreadsheet.insertSheet(propertyName);
    initializeSheet(sheet);
  } else {
    Logger.log(`Sheet "${propertyName}" already exists`);
  }
  
  // Get all complete months for this property
  const allMonths = getAllCompleteMonths(propertyId);
  Logger.log(`Found ${allMonths.length} complete month(s) to check`);
  
  let updatedCount = 0;
  let skippedCount = 0;
  
  // Process each month
  allMonths.forEach(monthString => {
    // Check if we already have this month's data
    if (hasMonthData(sheet, monthString)) {
      skippedCount++;
      return;
    }
    
    // Fetch and write data
    try {
      Logger.log(`Fetching data for ${formatMonthLabel(monthString)}...`);
      const monthData = fetchMonthlyData('properties/' + propertyId, monthString);
      writeMonthData(sheet, monthString, monthData);
      updatedCount++;
      Logger.log(`✓ Added ${formatMonthLabel(monthString)} data`);
      
      // Add a small delay to avoid rate limits
      Utilities.sleep(500);
    } catch (error) {
      Logger.log(`✗ Error fetching data for ${monthString}: ${error.message}`);
    }
  });
  
  Logger.log(`Summary: ${updatedCount} month(s) added, ${skippedCount} month(s) skipped (already exist)`);
  
  if (updatedCount === 0 && skippedCount === 0) {
    Logger.log(`No data available for ${propertyName}`);
  }
}

/**
 * Initialize sheet with headers (metrics in rows, months in columns)
 */
function initializeSheet(sheet) {
  const metrics = [
    ['Metric'],  // Header row will have months added dynamically
    // Pageviews
    ['Pageviews - Total'],
    ['Pageviews - Mobile'],
    ['Pageviews - Desktop'],
    ['Pageviews - Tablet'],
    ['Pageviews per Session'],
    ['Pageviews per User'],
    // Users
    ['Users - Total'],
    ['Users - Mobile'],
    ['Users - Desktop'],
    ['Users - Tablet'],
    ['Returning Users'],
    ['New Users'],
    ['Active Users'],
    // Sessions
    ['Sessions - Total'],
    ['Sessions - Mobile'],
    ['Sessions - Desktop'],
    ['Sessions - Tablet'],
    ['Sessions per User'],
    // Other metrics
    ['Bounce Rate'],
    ['Video Plays']
  ];
  
  sheet.getRange(1, 1, metrics.length, 1).setValues(metrics);
  sheet.getRange(1, 1, metrics.length, 1).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.autoResizeColumns(1, 1);
}

/**
 * Get the last complete month in YYYY-MM format
 */
function getLastCompleteMonth() {
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const year = lastMonth.getFullYear();
  const month = String(lastMonth.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

/**
 * Convert YYYY-MM to readable format like "October 2025"
 */
function formatMonthLabel(monthString) {
  const [year, month] = monthString.split('-');
  const date = new Date(parseInt(year), parseInt(month) - 1, 1);
  const monthName = date.toLocaleString('en-US', { month: 'long' });
  return `${monthName} ${year}`;
}

/**
 * Get all complete months from GA4 property creation to last complete month
 */
function getAllCompleteMonths(propertyId) {
  const months = [];
  const now = new Date();
  const lastCompleteMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  
  // Try to get property creation date
  let startDate;
  try {
    const property = AnalyticsAdmin.Properties.get('properties/' + propertyId);
    if (property.createTime) {
      startDate = new Date(property.createTime);
    } else {
      // Default to 2 years ago if creation date not available
      startDate = new Date(now.getFullYear() - 2, now.getMonth(), 1);
    }
  } catch (error) {
    Logger.log(`Could not fetch property creation date: ${error.message}`);
    // Default to 2 years ago
    startDate = new Date(now.getFullYear() - 2, now.getMonth(), 1);
  }
  
  // Start from the first full month after creation
  let currentMonth = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 1);
  
  // Generate all complete months up to last complete month
  while (currentMonth <= lastCompleteMonth) {
    const year = currentMonth.getFullYear();
    const month = String(currentMonth.getMonth() + 1).padStart(2, '0');
    months.push(`${year}-${month}`);
    currentMonth.setMonth(currentMonth.getMonth() + 1);
  }
  
  return months;
}

/**
 * Check if month data already exists in sheet
 */
function hasMonthData(sheet, monthString) {
  const lastCol = sheet.getLastColumn();
  
  // If only 1 column exists (just the metric names), no data yet
  if (lastCol < 2) {
    return false;
  }
  
  // Get all month headers (row 1, starting from column 2)
  const monthRow = sheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
  const formattedMonth = formatMonthLabel(monthString);
  
  // Check if this month already exists
  const exists = monthRow.some(cell => {
    const cellValue = String(cell).trim();
    return cellValue === monthString || cellValue === formattedMonth;
  });
  
  if (exists) {
    Logger.log(`✓ Month "${formattedMonth}" already exists in sheet - skipping`);
  }
  
  return exists;
}

/**
 * Fetch all metrics for a given month
 */
function fetchMonthlyData(propertyId, monthString) {
  const [year, month] = monthString.split('-');
  const startDate = `${year}-${month}-01`;
  const lastDay = new Date(parseInt(year), parseInt(month), 0).getDate();
  const endDate = `${year}-${month}-${lastDay}`;
  
  const data = {
    month: monthString,
    pageviews: {},
    users: {},
    sessions: {},
    bounceRate: 0,
    videoPlays: 0
  };
  
  // Fetch overall metrics (without video plays for now)
  const overallMetrics = fetchGA4Data(propertyId, startDate, endDate, [
    'screenPageViews',
    'totalUsers',
    'newUsers',
    'activeUsers',
    'sessions',
    'bounceRate'
  ], []);
  
  if (overallMetrics && overallMetrics.rows && overallMetrics.rows.length > 0) {
    const row = overallMetrics.rows[0].metricValues;
    data.pageviews.total = parseFloat(row[0].value || 0);
    data.users.total = parseFloat(row[1].value || 0);
    data.users.new = parseFloat(row[2].value || 0);
    data.users.active = parseFloat(row[3].value || 0);
    data.sessions.total = parseFloat(row[4].value || 0);
    data.bounceRate = parseFloat(row[5].value || 0);
    
    // Calculate returning users
    data.users.returning = data.users.total - data.users.new;
    
    // Calculate per session and per user metrics
    data.pageviews.perSession = data.sessions.total > 0 ? 
      data.pageviews.total / data.sessions.total : 0;
    data.pageviews.perUser = data.users.total > 0 ? 
      data.pageviews.total / data.users.total : 0;
    data.sessions.perUser = data.users.total > 0 ? 
      data.sessions.total / data.users.total : 0;
  }
  
  // Fetch video plays separately with correct filter syntax
  const videoMetrics = fetchGA4DataWithEventFilter(propertyId, startDate, endDate, 'video_start');
  if (videoMetrics && videoMetrics.rows && videoMetrics.rows.length > 0) {
    data.videoPlays = parseFloat(videoMetrics.rows[0].metricValues[0].value || 0);
  }
  
  // Fetch device breakdown for pageviews
  const devicePageviews = fetchGA4Data(propertyId, startDate, endDate, [
    'screenPageViews'
  ], [
    'deviceCategory'
  ]);
  
  if (devicePageviews && devicePageviews.rows) {
    devicePageviews.rows.forEach(row => {
      const device = row.dimensionValues[0].value.toLowerCase();
      const value = parseFloat(row.metricValues[0].value || 0);
      data.pageviews[device] = value;
    });
  }
  
  // Fetch device breakdown for users
  const deviceUsers = fetchGA4Data(propertyId, startDate, endDate, [
    'totalUsers'
  ], [
    'deviceCategory'
  ]);
  
  if (deviceUsers && deviceUsers.rows) {
    deviceUsers.rows.forEach(row => {
      const device = row.dimensionValues[0].value.toLowerCase();
      const value = parseFloat(row.metricValues[0].value || 0);
      data.users[device] = value;
    });
  }
  
  // Fetch device breakdown for sessions
  const deviceSessions = fetchGA4Data(propertyId, startDate, endDate, [
    'sessions'
  ], [
    'deviceCategory'
  ]);
  
  if (deviceSessions && deviceSessions.rows) {
    deviceSessions.rows.forEach(row => {
      const device = row.dimensionValues[0].value.toLowerCase();
      const value = parseFloat(row.metricValues[0].value || 0);
      data.sessions[device] = value;
    });
  }
  
  return data;
}

/**
 * Make API call to Google Analytics Data API
 */
function fetchGA4Data(propertyId, startDate, endDate, metrics, dimensions = []) {
  const request = {
    dateRanges: [{startDate: startDate, endDate: endDate}],
    metrics: metrics.map(m => ({name: m})),
    dimensions: dimensions.map(d => ({name: d}))
  };
  
  try {
    const response = AnalyticsData.Properties.runReport(request, propertyId);
    return response;
  } catch (error) {
    Logger.log(`Error fetching GA4 data: ${error.message}`);
    return null;
  }
}

/**
 * Fetch video plays with correct event filter syntax
 */
function fetchGA4DataWithEventFilter(propertyId, startDate, endDate, eventName) {
  const request = {
    dateRanges: [{startDate: startDate, endDate: endDate}],
    metrics: [{name: 'eventCount'}],
    dimensionFilter: {
      filter: {
        fieldName: 'eventName',
        stringFilter: {
          matchType: 'EXACT',
          value: eventName
        }
      }
    }
  };
  
  try {
    const response = AnalyticsData.Properties.runReport(request, propertyId);
    return response;
  } catch (error) {
    Logger.log(`Error fetching video plays: ${error.message}`);
    return null;
  }
}

/**
 * Write month data to sheet (as a new column)
 */
function writeMonthData(sheet, monthString, data) {
  // Find the next available column
  const lastCol = sheet.getLastColumn();
  const newCol = lastCol + 1;
  
  // Format month label
  const monthLabel = formatMonthLabel(monthString);
  
  // Add month header
  sheet.getRange(1, newCol).setValue(monthLabel);
  sheet.getRange(1, newCol).setFontWeight('bold');
  
  // Prepare data column
  const columnData = [
    [monthLabel],  // Month header (row 1)
    // Pageviews
    [data.pageviews.total || 0],
    [data.pageviews.mobile || 0],
    [data.pageviews.desktop || 0],
    [data.pageviews.tablet || 0],
    [data.pageviews.perSession || 0],
    [data.pageviews.perUser || 0],
    // Users
    [data.users.total || 0],
    [data.users.mobile || 0],
    [data.users.desktop || 0],
    [data.users.tablet || 0],
    [data.users.returning || 0],
    [data.users.new || 0],
    [data.users.active || 0],
    // Sessions
    [data.sessions.total || 0],
    [data.sessions.mobile || 0],
    [data.sessions.desktop || 0],
    [data.sessions.tablet || 0],
    [data.sessions.perUser || 0],
    // Other metrics
    [data.bounceRate || 0],
    [data.videoPlays || 0]
  ];
  
  sheet.getRange(1, newCol, columnData.length, 1).setValues(columnData);
  
  // Format bounce rate as percentage (row 20)
  sheet.getRange(20, newCol).setNumberFormat('0.00%');
  
  // Auto-resize the new column
  sheet.autoResizeColumns(newCol, 1);
}

/**
 * Create a monthly trigger to run automatically
 * Run this once to set up automation
 */
function createMonthlyTrigger() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'updateAllProperties') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger for the 1st of each month at 2 AM
  ScriptApp.newTrigger('updateAllProperties')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
  
  Logger.log('Monthly trigger created successfully');
}
