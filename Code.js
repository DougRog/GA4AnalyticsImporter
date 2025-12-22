/**
 * Google Analytics 4 to Google Sheets Automation (OPTIMIZED VERSION)
 * Automatically fetches property names from GA4
 * Pulls monthly GA4 data for multiple properties into separate sheets
 * Only updates data when a full month is complete
 * Months are displayed as columns
 *
 * OPTIMIZATIONS:
 * - Reduced API calls by combining metrics into single requests
 * - Caching of property metadata (names, creation dates)
 * - Batch processing of month checks
 * - Optimized sheet operations
 */

// Configuration - Add your property IDs (IMPORTANT: Just the numbers, no 'properties/' prefix)
const PROPERTY_IDS = [
  '326567794',
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
  '504008380',
  '309532396'
];

// Cache duration: 7 days (in seconds)
const CACHE_DURATION = 7 * 24 * 60 * 60;

/**
 * Get property display name from GA4 with caching
 */
function getPropertyName(propertyId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `propName_${propertyId}`;

  // Try to get from cache first
  const cached = cache.get(cacheKey);
  if (cached) {
    return cached;
  }

  // Fetch from API
  try {
    const property = AnalyticsAdmin.Properties.get('properties/' + propertyId);
    const name = property.displayName || propertyId;

    // Store in cache
    cache.put(cacheKey, name, CACHE_DURATION);
    return name;
  } catch (error) {
    Logger.log(`Error fetching property name for ${propertyId}: ${error.message}`);
    return propertyId;
  }
}

/**
 * Get property creation date with caching
 */
function getPropertyCreationDate(propertyId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `propCreate_${propertyId}`;

  // Try to get from cache first
  const cached = cache.get(cacheKey);
  if (cached) {
    return new Date(cached);
  }

  // Fetch from API
  try {
    const property = AnalyticsAdmin.Properties.get('properties/' + propertyId);
    if (property.createTime) {
      const createTime = property.createTime;
      cache.put(cacheKey, createTime, CACHE_DURATION);
      return new Date(createTime);
    }
  } catch (error) {
    Logger.log(`Could not fetch property creation date: ${error.message}`);
  }

  // Default to 2 years ago
  const now = new Date();
  return new Date(now.getFullYear() - 2, now.getMonth(), 1);
}

/**
 * Main function to update all properties
 * Run this manually or set up a trigger to run monthly
 */
function updateAllProperties() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Debug: Log the property IDs array
  Logger.log(`Processing ${PROPERTY_IDS.length} property ID(s): ${JSON.stringify(PROPERTY_IDS)}`);

  // Filter out any undefined, null, or empty values
  const validPropertyIds = PROPERTY_IDS.filter(id => {
    if (!id || id.trim() === '') {
      Logger.log(`Skipping invalid property ID: ${id}`);
      return false;
    }
    return true;
  });

  if (validPropertyIds.length === 0) {
    Logger.log('ERROR: No valid property IDs found in PROPERTY_IDS array!');
    return;
  }

  Logger.log(`Found ${validPropertyIds.length} valid property ID(s)`);

  validPropertyIds.forEach(propertyId => {
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

  // Batch check which months already exist
  const existingMonths = getExistingMonths(sheet);
  const monthsToFetch = allMonths.filter(month => !existingMonths.has(month));

  if (monthsToFetch.length === 0) {
    Logger.log(`All months already exist - nothing to fetch`);
    return;
  }

  Logger.log(`Need to fetch ${monthsToFetch.length} new month(s)`);

  // Process each month
  let updatedCount = 0;
  monthsToFetch.forEach((monthString, index) => {
    try {
      Logger.log(`Fetching data for ${formatMonthLabel(monthString)}...`);
      const monthData = fetchMonthlyData('properties/' + propertyId, monthString);
      writeMonthData(sheet, monthString, monthData);
      updatedCount++;
      Logger.log(`✓ Added ${formatMonthLabel(monthString)} data`);

      // Add a small delay every 5 requests to avoid rate limits
      if ((index + 1) % 5 === 0 && index < monthsToFetch.length - 1) {
        Utilities.sleep(500);
      }
    } catch (error) {
      Logger.log(`✗ Error fetching data for ${monthString}: ${error.message}`);
    }
  });

  Logger.log(`Summary: ${updatedCount} month(s) added, ${allMonths.length - monthsToFetch.length} month(s) skipped (already exist)`);
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

  // Get property creation date (cached)
  const startDate = getPropertyCreationDate(propertyId);

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
 * Get all existing months from sheet in one read operation
 */
function getExistingMonths(sheet) {
  const existingMonths = new Set();
  const lastCol = sheet.getLastColumn();

  // If only 1 column exists (just the metric names), no data yet
  if (lastCol < 2) {
    return existingMonths;
  }

  // Get all month headers (row 1, starting from column 2) in one read
  const monthRow = sheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];

  monthRow.forEach(cell => {
    const cellValue = String(cell).trim();
    if (cellValue) {
      // Store both the formatted version and try to convert back to YYYY-MM
      existingMonths.add(cellValue);

      // Try to parse formatted month back to YYYY-MM format
      const parsed = parseMonthLabel(cellValue);
      if (parsed) {
        existingMonths.add(parsed);
      }
    }
  });

  return existingMonths;
}

/**
 * Parse formatted month label back to YYYY-MM format
 */
function parseMonthLabel(label) {
  const match = label.match(/^(\w+)\s+(\d{4})$/);
  if (!match) return null;

  const monthName = match[1];
  const year = match[2];

  const monthMap = {
    'January': '01', 'February': '02', 'March': '03', 'April': '04',
    'May': '05', 'June': '06', 'July': '07', 'August': '08',
    'September': '09', 'October': '10', 'November': '11', 'December': '12'
  };

  const month = monthMap[monthName];
  return month ? `${year}-${month}` : null;
}

/**
 * Fetch all metrics for a given month - OPTIMIZED VERSION
 * Reduces 5 API calls down to 2 by combining metrics
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

  // OPTIMIZATION: Combine all metrics and device breakdown in ONE API call
  const combinedMetrics = fetchGA4Data(propertyId, startDate, endDate,
    [
      'screenPageViews',
      'totalUsers',
      'newUsers',
      'activeUsers',
      'sessions',
      'bounceRate'
    ],
    ['deviceCategory']  // Add device dimension to get breakdown
  );

  if (combinedMetrics && combinedMetrics.rows && combinedMetrics.rows.length > 0) {
    // Initialize totals
    let totalPageviews = 0;
    let totalUsers = 0;
    let totalNewUsers = 0;
    let totalActiveUsers = 0;
    let totalSessions = 0;
    let totalBounceRate = 0;

    // Process each device row
    combinedMetrics.rows.forEach(row => {
      const device = row.dimensionValues[0].value.toLowerCase();
      const metrics = row.metricValues;

      const pageviews = parseFloat(metrics[0].value || 0);
      const users = parseFloat(metrics[1].value || 0);
      const newUsers = parseFloat(metrics[2].value || 0);
      const activeUsers = parseFloat(metrics[3].value || 0);
      const sessions = parseFloat(metrics[4].value || 0);
      const bounceRate = parseFloat(metrics[5].value || 0);

      // Store device-specific values
      data.pageviews[device] = pageviews;
      data.users[device] = users;
      data.sessions[device] = sessions;

      // Accumulate totals
      totalPageviews += pageviews;
      totalUsers += users;
      totalNewUsers += newUsers;
      totalActiveUsers += activeUsers;
      totalSessions += sessions;
      // For bounce rate, we'll take the weighted average
      totalBounceRate += bounceRate * sessions;
    });

    // Set total values
    data.pageviews.total = totalPageviews;
    data.users.total = totalUsers;
    data.users.new = totalNewUsers;
    data.users.active = totalActiveUsers;
    data.sessions.total = totalSessions;
    data.bounceRate = totalSessions > 0 ? totalBounceRate / totalSessions : 0;

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

  // Fetch video plays separately (still requires event filter)
  const videoMetrics = fetchGA4DataWithEventFilter(propertyId, startDate, endDate, 'video_start');
  if (videoMetrics && videoMetrics.rows && videoMetrics.rows.length > 0) {
    data.videoPlays = parseFloat(videoMetrics.rows[0].metricValues[0].value || 0);
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

  // Prepare data column (all values in one array for batch write)
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

  // Single batch write operation
  sheet.getRange(1, newCol, columnData.length, 1).setValues(columnData);

  // Format header and bounce rate
  sheet.getRange(1, newCol).setFontWeight('bold');
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

/**
 * Clear all cached property data (useful for testing or if data changes)
 */
function clearPropertyCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(PROPERTY_IDS.flatMap(id => [
    `propName_${id}`,
    `propCreate_${id}`
  ]));
  Logger.log('Property cache cleared');
}

/**
 * Diagnostic function to check PROPERTY_IDS configuration
 * Run this to verify your property IDs are set up correctly
 */
function testPropertyIdsConfiguration() {
  Logger.log('=== PROPERTY_IDS CONFIGURATION TEST ===');
  Logger.log(`Array length: ${PROPERTY_IDS.length}`);
  Logger.log(`Array contents: ${JSON.stringify(PROPERTY_IDS, null, 2)}`);

  PROPERTY_IDS.forEach((id, index) => {
    Logger.log(`\nProperty ID #${index + 1}:`);
    Logger.log(`  Value: "${id}"`);
    Logger.log(`  Type: ${typeof id}`);
    Logger.log(`  Is valid: ${id && typeof id === 'string' && id.trim() !== ''}`);
  });

  const validIds = PROPERTY_IDS.filter(id => id && typeof id === 'string' && id.trim() !== '');
  Logger.log(`\nValid property IDs: ${validIds.length} out of ${PROPERTY_IDS.length}`);

  if (validIds.length === 0) {
    Logger.log('\n⚠️ ERROR: No valid property IDs found!');
    Logger.log('Please check the PROPERTY_IDS array at the top of the script.');
    Logger.log('Make sure you have at least one uncommented property ID like: \'309532396\'');
  } else {
    Logger.log('\n✓ Configuration looks good!');
  }
}
