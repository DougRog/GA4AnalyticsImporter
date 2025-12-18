/**
 * Google Analytics 4 to Google Sheets Automation (AUTO-FETCH VERSION)
 * Automatically fetches property names from GA4
 * Pulls monthly GA4 data for multiple properties into separate sheets
 * Only updates data when a full month is complete
 * Months are displayed as columns
 */

// ============================================================================
// CONFIGURATION CONSTANTS
// ============================================================================

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

// API rate limiting
const API_DELAY_MS = 500;

// Default lookback period if property creation date unavailable
const DEFAULT_LOOKBACK_YEARS = 2;

// Trigger configuration
const TRIGGER_DAY_OF_MONTH = 1;
const TRIGGER_HOUR = 2;

// Metrics configuration - defines the order and structure of all metrics
const METRICS_CONFIG = [
  { label: 'Metric', row: 1 },
  // Pageviews
  { label: 'Pageviews - Total', row: 2, dataPath: ['pageviews', 'total'] },
  { label: 'Pageviews - Mobile', row: 3, dataPath: ['pageviews', 'mobile'] },
  { label: 'Pageviews - Desktop', row: 4, dataPath: ['pageviews', 'desktop'] },
  { label: 'Pageviews - Tablet', row: 5, dataPath: ['pageviews', 'tablet'] },
  { label: 'Pageviews per Session', row: 6, dataPath: ['pageviews', 'perSession'] },
  { label: 'Pageviews per User', row: 7, dataPath: ['pageviews', 'perUser'] },
  // Users
  { label: 'Users - Total', row: 8, dataPath: ['users', 'total'] },
  { label: 'Users - Mobile', row: 9, dataPath: ['users', 'mobile'] },
  { label: 'Users - Desktop', row: 10, dataPath: ['users', 'desktop'] },
  { label: 'Users - Tablet', row: 11, dataPath: ['users', 'tablet'] },
  { label: 'Returning Users', row: 12, dataPath: ['users', 'returning'] },
  { label: 'New Users', row: 13, dataPath: ['users', 'new'] },
  { label: 'Active Users', row: 14, dataPath: ['users', 'active'] },
  // Sessions
  { label: 'Sessions - Total', row: 15, dataPath: ['sessions', 'total'] },
  { label: 'Sessions - Mobile', row: 16, dataPath: ['sessions', 'mobile'] },
  { label: 'Sessions - Desktop', row: 17, dataPath: ['sessions', 'desktop'] },
  { label: 'Sessions - Tablet', row: 18, dataPath: ['sessions', 'tablet'] },
  { label: 'Sessions per User', row: 19, dataPath: ['sessions', 'perUser'] },
  // Other metrics
  { label: 'Bounce Rate', row: 20, dataPath: ['bounceRate'], format: '0.00%' },
  { label: 'Video Plays', row: 21, dataPath: ['videoPlays'] }
];

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get property display name from GA4
 * @param {string} propertyId - The GA4 property ID
 * @return {string} The property display name or property ID if unavailable
 */
function getPropertyName(propertyId) {
  // Explicit check at the start of this function
  if (propertyId === undefined || propertyId === null) {
    const errorMsg = `CRITICAL ERROR: getPropertyName called with ${propertyId} value. This should never happen if validation passed.`;
    Logger.log(errorMsg);
    Logger.log(`PROPERTY_IDS array contains: ${JSON.stringify(PROPERTY_IDS)}`);
    throw new Error(errorMsg);
  }

  if (typeof propertyId !== 'string' || propertyId.trim() === '') {
    const errorMsg = `CRITICAL ERROR: getPropertyName called with invalid propertyId: type=${typeof propertyId}, value="${propertyId}"`;
    Logger.log(errorMsg);
    throw new Error(errorMsg);
  }

  try {
    const property = AnalyticsAdmin.Properties.get('properties/' + propertyId);
    return property.displayName || propertyId;
  } catch (error) {
    Logger.log(`Error fetching property name for ${propertyId}: ${error.message}`);
    return propertyId;  // Fallback to property ID if name can't be fetched
  }
}

/**
 * Generate sheet name with property ID to prevent conflicts
 * @param {string} propertyName - The property display name
 * @param {string} propertyId - The GA4 property ID
 * @return {string} Sheet name in format "PropertyName (ID)"
 */
function generateSheetName(propertyName, propertyId) {
  // Include property ID in parentheses to prevent conflicts
  // Google Sheets tab names have a 100 character limit
  const sheetName = `${propertyName} (${propertyId})`;

  // Truncate if too long (reserve space for property ID)
  if (sheetName.length > 100) {
    const maxNameLength = 100 - propertyId.length - 3; // 3 for " ()"
    return `${propertyName.substring(0, maxNameLength)} (${propertyId})`;
  }

  return sheetName;
}

/**
 * Validate property IDs array
 * @throws {Error} If PROPERTY_IDS is invalid
 */
function validatePropertyIds() {
  if (!Array.isArray(PROPERTY_IDS)) {
    throw new Error('PROPERTY_IDS must be an array');
  }

  if (PROPERTY_IDS.length === 0) {
    throw new Error('PROPERTY_IDS array is empty. Please add at least one property ID.');
  }

  Logger.log(`Validating ${PROPERTY_IDS.length} property ID(s)...`);

  PROPERTY_IDS.forEach((id, index) => {
    // Log each property ID for debugging
    Logger.log(`  [${index}] Type: ${typeof id}, Value: "${id}"`);

    // Check for undefined, null, or non-string values
    if (id === undefined) {
      throw new Error(`Property ID at index ${index} is undefined. Check for trailing commas or empty elements in PROPERTY_IDS array.`);
    }

    if (id === null) {
      throw new Error(`Property ID at index ${index} is null.`);
    }

    if (typeof id !== 'string') {
      throw new Error(`Property ID at index ${index} must be a string, got ${typeof id}: ${id}`);
    }

    if (id.trim() === '') {
      throw new Error(`Property ID at index ${index} is an empty string.`);
    }

    // Validate it looks like a property ID (should be numeric)
    if (!/^\d+$/.test(id.trim())) {
      Logger.log(`  ⚠️  Warning: Property ID at index ${index} ("${id}") doesn't look like a numeric ID`);
    }
  });

  Logger.log(`✓ All ${PROPERTY_IDS.length} property ID(s) validated successfully`);
}

/**
 * Main function to update all properties
 * Run this manually or set up a trigger to run monthly
 * @throws {Error} If property IDs are invalid or spreadsheet unavailable
 */
function updateAllProperties() {
  // Validate configuration
  validatePropertyIds();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('No active spreadsheet found. Please open a Google Sheet.');
  }

  Logger.log(`\n${'='.repeat(60)}`);
  Logger.log(`Starting update for ${PROPERTY_IDS.length} propert${PROPERTY_IDS.length === 1 ? 'y' : 'ies'}`);
  Logger.log(`${'='.repeat(60)}\n`);

  let successCount = 0;
  let errorCount = 0;

  PROPERTY_IDS.forEach(propertyId => {
    try {
      const propertyName = getPropertyName(propertyId);
      const sheetName = generateSheetName(propertyName, propertyId);
      updatePropertyData(spreadsheet, propertyId, propertyName, sheetName);
      successCount++;
    } catch (error) {
      Logger.log(`✗ Error updating ${propertyId}: ${error.message}`);
      if (error.stack) {
        Logger.log(`Stack trace: ${error.stack}`);
      }
      errorCount++;
    }
  });

  Logger.log(`\n${'='.repeat(60)}`);
  Logger.log(`Update complete: ${successCount} succeeded, ${errorCount} failed`);
  Logger.log(`${'='.repeat(60)}\n`);
}

/**
 * Update data for a single property
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The active spreadsheet
 * @param {string} propertyId - The GA4 property ID
 * @param {string} propertyName - The display name for the property
 * @param {string} sheetName - The sheet name (includes property ID to prevent conflicts)
 */
function updatePropertyData(spreadsheet, propertyId, propertyName, sheetName) {
  Logger.log(`\n${'='.repeat(40)}`);
  Logger.log(`Processing: ${propertyName} (${propertyId})`);
  Logger.log(`${'='.repeat(40)}`);

  let sheet = spreadsheet.getSheetByName(sheetName);

  // Create sheet if it doesn't exist
  if (!sheet) {
    Logger.log(`Creating new sheet: ${sheetName}`);
    sheet = spreadsheet.insertSheet(sheetName);
    initializeSheet(sheet);
  } else {
    Logger.log(`Sheet "${sheetName}" already exists`);
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
      Utilities.sleep(API_DELAY_MS);
    } catch (error) {
      Logger.log(`✗ Error fetching data for ${monthString}: ${error.message}`);
      if (error.stack) {
        Logger.log(`  Stack: ${error.stack}`);
      }
    }
  });

  Logger.log(`Summary: ${updatedCount} month(s) added, ${skippedCount} month(s) skipped (already exist)`);

  if (updatedCount === 0 && skippedCount === 0) {
    Logger.log(`No data available for ${propertyName}`);
  }
}

/**
 * Initialize sheet with headers (metrics in rows, months in columns)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to initialize
 */
function initializeSheet(sheet) {
  // Generate metrics array from METRICS_CONFIG
  const metrics = METRICS_CONFIG.map(config => [config.label]);

  sheet.getRange(1, 1, metrics.length, 1).setValues(metrics);
  sheet.getRange(1, 1, metrics.length, 1).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.autoResizeColumns(1, 1);
}

/**
 * Get the last complete month in YYYY-MM format
 * @return {string} Last complete month in YYYY-MM format
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
 * @param {string} monthString - Month in YYYY-MM format
 * @return {string} Formatted month like "October 2025"
 */
function formatMonthLabel(monthString) {
  const [year, month] = monthString.split('-');
  const date = new Date(parseInt(year), parseInt(month) - 1, 1);
  const monthName = date.toLocaleString('en-US', { month: 'long' });
  return `${monthName} ${year}`;
}

/**
 * Get all complete months from GA4 property creation to last complete month
 * ONLY returns complete months - excludes the current partial month
 * @param {string} propertyId - The GA4 property ID
 * @return {string[]} Array of month strings in YYYY-MM format (only complete months)
 */
function getAllCompleteMonths(propertyId) {
  const months = [];
  const now = new Date();

  // Calculate last complete month (always previous month, never current month)
  // This ensures we only pull months with a full month's worth of data
  const lastCompleteMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);

  Logger.log(`Current date: ${now.toISOString().split('T')[0]}`);
  Logger.log(`Last complete month: ${lastCompleteMonth.getFullYear()}-${String(lastCompleteMonth.getMonth() + 1).padStart(2, '0')}`);

  // Try to get property creation date
  let startDate;
  try {
    const property = AnalyticsAdmin.Properties.get('properties/' + propertyId);
    if (property.createTime) {
      startDate = new Date(property.createTime);
      Logger.log(`Property created: ${property.createTime}`);
    } else {
      // Default to lookback period if creation date not available
      startDate = new Date(now.getFullYear() - DEFAULT_LOOKBACK_YEARS, now.getMonth(), 1);
      Logger.log(`Property creation date unavailable, using ${DEFAULT_LOOKBACK_YEARS} year lookback`);
    }
  } catch (error) {
    Logger.log(`Could not fetch property creation date: ${error.message}`);
    // Default to lookback period
    startDate = new Date(now.getFullYear() - DEFAULT_LOOKBACK_YEARS, now.getMonth(), 1);
  }

  // Start from the first full month after creation
  let currentMonth = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 1);

  // Generate all complete months up to last complete month (excluding current month)
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
 * Prevents duplicate imports by checking existing column headers
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to check
 * @param {string} monthString - Month in YYYY-MM format (e.g., "2025-11")
 * @return {boolean} True if month data already exists, false if it needs to be imported
 */
function hasMonthData(sheet, monthString) {
  const lastCol = sheet.getLastColumn();

  // If only 1 column exists (just the metric names), no data yet
  if (lastCol < 2) {
    Logger.log(`No data columns exist yet for ${formatMonthLabel(monthString)}`);
    return false;
  }

  // Get all month headers (row 1, starting from column 2)
  const monthRow = sheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
  const formattedMonth = formatMonthLabel(monthString);

  // Check if this month already exists in any column
  // We check for both formats: "2025-11" and "November 2025" to be robust
  const exists = monthRow.some(cell => {
    if (!cell) return false; // Skip empty cells
    const cellValue = String(cell).trim();
    return cellValue === monthString || cellValue === formattedMonth;
  });

  if (exists) {
    Logger.log(`✓ Month "${formattedMonth}" already exists in sheet - SKIPPING to prevent duplicate`);
  } else {
    Logger.log(`✗ Month "${formattedMonth}" not found - will import`);
  }

  return exists;
}

/**
 * Helper function to fetch device breakdown for a specific metric
 * @param {string} propertyId - The GA4 property ID
 * @param {string} startDate - Start date in YYYY-MM-DD format
 * @param {string} endDate - End date in YYYY-MM-DD format
 * @param {string} metricName - The GA4 metric name
 * @param {Object} targetObject - The object to populate with device data
 */
function fetchDeviceBreakdown(propertyId, startDate, endDate, metricName, targetObject) {
  const deviceData = fetchGA4Data(propertyId, startDate, endDate, [metricName], ['deviceCategory']);

  if (deviceData && deviceData.rows) {
    deviceData.rows.forEach(row => {
      const device = row.dimensionValues[0].value.toLowerCase();
      const value = parseFloat(row.metricValues[0].value || 0);
      targetObject[device] = value;
    });
  }
}

/**
 * Fetch all metrics for a given month
 * @param {string} propertyId - The GA4 property ID (with 'properties/' prefix)
 * @param {string} monthString - Month in YYYY-MM format
 * @return {Object} Object containing all fetched metrics
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

  // Fetch device breakdowns using helper function
  fetchDeviceBreakdown(propertyId, startDate, endDate, 'screenPageViews', data.pageviews);
  fetchDeviceBreakdown(propertyId, startDate, endDate, 'totalUsers', data.users);
  fetchDeviceBreakdown(propertyId, startDate, endDate, 'sessions', data.sessions);

  return data;
}

/**
 * Make API call to Google Analytics Data API
 * @param {string} propertyId - The GA4 property ID (with 'properties/' prefix)
 * @param {string} startDate - Start date in YYYY-MM-DD format
 * @param {string} endDate - End date in YYYY-MM-DD format
 * @param {string[]} metrics - Array of metric names to fetch
 * @param {string[]} dimensions - Array of dimension names to fetch
 * @return {Object|null} GA4 API response or null if error
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
 * Fetch event data with dimension filter (used for video plays)
 * @param {string} propertyId - The GA4 property ID (with 'properties/' prefix)
 * @param {string} startDate - Start date in YYYY-MM-DD format
 * @param {string} endDate - End date in YYYY-MM-DD format
 * @param {string} eventName - The event name to filter by
 * @return {Object|null} GA4 API response or null if error
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
    Logger.log(`Error fetching event data for ${eventName}: ${error.message}`);
    return null;
  }
}

/**
 * Helper function to get nested value from object using path array
 * @param {Object} obj - The object to traverse
 * @param {string[]} path - Array of property names to traverse
 * @return {*} The value at the path or 0 if not found
 */
function getValueByPath(obj, path) {
  let value = obj;
  for (const key of path) {
    value = value?.[key];
    if (value === undefined || value === null) return 0;
  }
  return value;
}

/**
 * Write month data to sheet (as a new column)
 * Uses METRICS_CONFIG to ensure data is written in the correct order
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to write to
 * @param {string} monthString - Month in YYYY-MM format
 * @param {Object} data - The data object containing all metrics
 */
function writeMonthData(sheet, monthString, data) {
  // Find the next available column
  const lastCol = sheet.getLastColumn();
  const newCol = lastCol + 1;

  // Format month label
  const monthLabel = formatMonthLabel(monthString);

  // Build column data from METRICS_CONFIG
  const columnData = METRICS_CONFIG.map((config, index) => {
    if (index === 0) {
      // First row is the month header
      return [monthLabel];
    }
    // Get value from data object using the dataPath
    const value = getValueByPath(data, config.dataPath);
    return [value];
  });

  // Write all data at once
  sheet.getRange(1, newCol, columnData.length, 1).setValues(columnData);

  // Apply special formatting based on METRICS_CONFIG
  METRICS_CONFIG.forEach(config => {
    if (config.format && config.row) {
      sheet.getRange(config.row, newCol).setNumberFormat(config.format);
    }
  });

  // Make month header bold
  sheet.getRange(1, newCol).setFontWeight('bold');

  // Auto-resize the new column
  sheet.autoResizeColumns(newCol, 1);
}

/**
 * Create a monthly trigger to run automatically
 * Run this once to set up automation
 * Deletes existing triggers for updateAllProperties to avoid duplicates
 */
function createMonthlyTrigger() {
  // Delete existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'updateAllProperties') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });

  if (deletedCount > 0) {
    Logger.log(`Deleted ${deletedCount} existing trigger(s)`);
  }

  // Create new trigger for the configured day/hour of each month
  ScriptApp.newTrigger('updateAllProperties')
    .timeBased()
    .onMonthDay(TRIGGER_DAY_OF_MONTH)
    .atHour(TRIGGER_HOUR)
    .create();

  Logger.log(`Monthly trigger created successfully`);
  Logger.log(`Will run on day ${TRIGGER_DAY_OF_MONTH} of each month at ${TRIGGER_HOUR}:00`);
}
