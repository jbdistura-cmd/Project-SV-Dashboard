// Sheet cache and utilities
const SHEET_CACHE = {};
const SPREADSHEET_ID = '1fYx3TXgiKzmkHTwy30aOUSRNChB5h0AUP-f84QklXio';

function getSheet(sheetName) {
  if (!SHEET_CACHE[sheetName]) {
    SHEET_CACHE[sheetName] = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  }
  return SHEET_CACHE[sheetName];
}

function withRetry(fn, maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return fn();
    } catch (error) {
      if (i === maxRetries - 1) throw error;
      Utilities.sleep(1000 * Math.pow(2, i));
    }
  }
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// Serve the HTML file for the dashboard
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Admin Dashboard Panel')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Get the HTML content for a specific page
function getPage(page) {
  const content = HtmlService.createHtmlOutputFromFile(pageToFileName(page)).getContent();
  return content;
}

// Mapping pages to file names
function pageToFileName(page) {
  const pageMap = {
    'dashboard': 'Dashboard',
    'analysis': 'Analysis',
    'sales-dump-ffh': 'SalesDumpFFH',
    'sales-dump-wls': 'SalesDumpWLS',
    'dnc': 'DNC'
  };
  return pageMap[page] || 'Dashboard';
}

// Fetch sales data from Google Sheets and include timestamp for polling
function getSalesData() {
  return withRetry(() => {
    const sheet = getSheet('Sales Dump');
    const svRosterSheet = getSheet('SV Roster');
    const evalSheetWLS = getSheet('Eval Form Submissions WLS');
    const evalSheetFFH = getSheet('Eval Form Submissions FFH');

  if (!sheet || !svRosterSheet) {
    Logger.log('Sales Dump or SV Roster sheet not found');
    return { data: [], svRoster: [], ids: [], assigneeEmails: [], userEmail: '', timestamp: new Date().getTime() };
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  const processedData = data.slice(1).map(row => row.slice(1, 19));
  const ids = data.slice(1).map(row => row[0].toString());
  const assigneeEmails = data.slice(1).map(row => row[19]);

  const svRosterData = svRosterSheet.getRange('B2:C').getValues();
  const svRosterNames = svRosterData
    .filter(row => row[1] === 'WLS' || row[1] === 'FFH')
    .map(row => row[0])
    .filter(name => name);

  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = getLastEditTimestamp();

    const takenRows = data.slice(1).map(row => ({
        id: row[0],
        svName: row[2],
        assignee: row[18],
        status: row[1]  // Assuming Column B is index 1
    }));

  let evalData = { WLS: {}, FFH: {} };

  // Handle WLS eval data
  if (evalSheetWLS) {
    const evalRangeWLS = evalSheetWLS.getDataRange();
    const evalValuesWLS = evalRangeWLS.getValues();
    evalValuesWLS.slice(1).forEach(row => {
      if (row[5]) {
        const flags = parseFlags(row.slice(18, 28));
        const nonEmptyFlags = Object.fromEntries(
          Object.entries(flags).filter(([_, value]) => value.length > 0)
        );
        
        evalData.WLS[row[5]] = {
          callDisposition: row[16],
          zviqEvalID: row[17],
          selectedFlags: nonEmptyFlags,
          profFlagRemarks: row[28] || '',
          goodSaleRemarks: row[29] || ''
        };
      }
    });
  }

  // Handle FFH eval data
if (evalSheetFFH) {
    const evalRangeFFH = evalSheetFFH.getDataRange();
    const evalValuesFFH = evalRangeFFH.getValues();
    evalValuesFFH.slice(1).forEach(row => {
        if (row[5]) {
            const flags = parseFlagsFFH(row.slice(18, 29));
            const nonEmptyFlags = Object.fromEntries(
                Object.entries(flags).filter(([_, value]) => value.length > 0)
            );
            
            evalData.FFH[row[5]] = {
                callDisposition: row[16],
                zviqEvalID: row[17],
                selectedFlags: nonEmptyFlags,
                profFlagRemarks: row[29] || '',  // Changed from row[30]
                goodSaleRemarks: row[30] || ''   // Changed from row[31]
            };
        }
    });
}
  Logger.log('Eval Data:', evalData);

    return {
      data: processedData,
      svRoster: svRosterNames,
      ids: ids,
      assigneeEmails: assigneeEmails,
      userEmail: userEmail,
      timestamp: timestamp,
      takenRows: takenRows,
      evalData: evalData
    };
  });
}


function getBacklogData() {
  return withRetry(() => {
    const sheet = getSheet('Backlogs');
    const svRosterSheet = getSheet('SV Roster');
  
  if (!sheet || !svRosterSheet) {
    Logger.log('Backlogs or SV Roster sheet not found');
    return { data: [], svRoster: [], ids: [], assigneeEmails: [], userEmail: '', timestamp: new Date().getTime() };
  }

  const dataRange = sheet.getDataRange();
  const rawData = dataRange.getValues();
  Logger.log('Raw Backlog Data rows: ' + rawData.length);

  const processedData = rawData.slice(1).map(row => row.slice(1, 19));
  const ids = rawData.slice(1).map(row => row[0].toString()); // Ensure IDs are strings
  const assigneeEmails = rawData.slice(1).map(row => row[19] || '');

  // Keep SV Roster for WLS functionality
  const svRosterData = svRosterSheet.getRange('B2:C').getValues();
  const svRosterNames = svRosterData
    .filter(row => row[1] === 'WLS' || row[1] === 'FFH')
    .map(row => row[0])
    .filter(name => name);

  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = new Date().getTime();

  Logger.log('Processed Backlog Data rows: ' + processedData.length);

    // Return all data for both WLS and FFH
    return {
      data: processedData,
      svRoster: svRosterNames,
      ids: ids,
      assigneeEmails: assigneeEmails,
      userEmail: userEmail,
      timestamp: timestamp,
      evalData: {}
    };
  });
}

function updateBacklogSVName(uniqueId, svName) {
  return withLock(() => withRetry(() => {
    const sheet = getSheet('Backlogs');

  if (!sheet) {
    Logger.log('Backlogs sheet not found');
    return null;
  }

  try {
    const data = sheet.getDataRange().getValues();
    Logger.log('Looking for uniqueId:', uniqueId, 'in Backlogs');

    const rowIndex = data.findIndex(row => String(row[0]).trim() === String(uniqueId).trim());
    Logger.log('Found row index:', rowIndex);

    if (rowIndex !== -1) {
      // Update SV Name (Column C)
      sheet.getRange(rowIndex + 1, 3).setValue(svName);
      
      // Update Assignee (Column S)
      const userEmail = Session.getActiveUser().getEmail();
      sheet.getRange(rowIndex + 1, 20).setValue(userEmail);
      
      const statusRange = sheet.getRange(rowIndex + 1, 2);
      const currentFormula = statusRange.getFormula();
      if (currentFormula) {
        statusRange.setFormula(currentFormula);
      }

      Logger.log('Updated SV Name to:', svName, 'and Assignee to:', userEmail, 'for row:', rowIndex + 1);
      
      // Return both the status, updated SV name, and assignee
      return {
        status: statusRange.getValue(),
        svName: svName,
        assignee: userEmail
      };
    } else {
      Logger.log('No matching row found for ID:', uniqueId);
      return null;
    }
  } catch (error) {
    Logger.log('Error in updateBacklogSVName:', error.toString());
    throw error;
  }
  }));
}

function updateBacklogDateValidated(uniqueId, dateValidated) {
  const sheetId = '1fYx3TXgiKzmkHTwy30aOUSRNChB5h0AUP-f84QklXio';
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Backlogs');

  if (!sheet) {
    Logger.log('Backlogs sheet not found');
    return null;
  }

  try {
    const data = sheet.getDataRange().getValues();
    Logger.log('Looking for uniqueId:', uniqueId, 'in Backlogs');

    const rowIndex = data.findIndex(row => String(row[0]).trim() === String(uniqueId).trim());
    Logger.log('Found row index:', rowIndex);

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 1, 4).setValue(dateValidated);
      Logger.log('Updated Date Validated to:', dateValidated, 'for row:', rowIndex + 1);
      
      const statusRange = sheet.getRange(rowIndex + 1, 2);
      
      // Return both the status and the updated date validated
      return {
        status: statusRange.getValue(),
        dateValidated: dateValidated
      };
    } else {
      Logger.log('No matching row found for ID:', uniqueId);
      return null;
    }
  } catch (error) {
    Logger.log('Error in updateBacklogDateValidated:', error.toString());
    throw error;
  }
}

function updateBacklogDateValidatedFFH(uniqueId, dateValidated) {
  const sheetId = '1fYx3TXgiKzmkHTwy30aOUSRNChB5h0AUP-f84QklXio';
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Backlogs');

  if (!sheet) {
    Logger.log('Backlogs sheet not found');
    return null;
  }

  try {
    const data = sheet.getDataRange().getValues();
    Logger.log('Looking for uniqueId:', uniqueId, 'in Backlogs');

    const rowIndex = data.findIndex(row => String(row[0]).trim() === String(uniqueId).trim());
    Logger.log('Found row index:', rowIndex);

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 1, 4).setValue(dateValidated);
      Logger.log('Updated Date Validated to:', dateValidated, 'for row:', rowIndex + 1);
      
      const statusRange = sheet.getRange(rowIndex + 1, 2);
      
      return {
        status: statusRange.getValue(),
        dateValidated: dateValidated
      };
    } else {
      Logger.log('No matching row found for ID:', uniqueId);
      return null;
    }
  } catch (error) {
    Logger.log('Error in updateBacklogDateValidatedFFH:', error.toString());
    throw error;
  }
}

// Helper function to parse WLS flags
function parseFlags(flagsArray) {
    const categories = [
        "Regulatory", "Privacy", "Price Information", "Plan Information", 
        "Order Processing", "Account Documentation", "Call Summary", 
        "Business Intelligence", "Marketing Objectives", "Professionalism"
    ];
    
    let selectedFlags = {};
    flagsArray.forEach((flags, index) => {
        if (flags) {
            const flagList = flags.split(', ').filter(flag => flag.trim() !== '');
            if (flagList.length > 0) {
                selectedFlags[categories[index]] = flagList;
            }
        }
    });
    return selectedFlags;
}

// Helper function to parse FFH flags
function parseFlagsFFH(flagsArray) {
    const categories = [
        "Regulatory", "Privacy", "Price Information", "Plan Information", 
        "Order Processing", "Account Documentation", "Call Summary", 
        "Customer Experience", "Business Intelligence", "Marketing Objectives", 
        "Professionalism"
    ];
    
    let selectedFlags = {};
    flagsArray.forEach((flags, index) => {
        if (flags) {
            const flagList = flags.split(', ').filter(flag => flag.trim() !== '');
            if (flagList.length > 0) {
                selectedFlags[categories[index]] = flagList;
            }
        }
    });
    return selectedFlags;
}

// Simplified row availability check using LockService
function checkRowAvailability(uniqueId) {
  return withRetry(() => {
    const sheet = getSheet('Sales Dump');
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => String(row[0]).trim() === String(uniqueId).trim());
    
    if (rowIndex === -1) return false;
    
    const assigneeEmail = data[rowIndex][18];
    const currentUserEmail = Session.getActiveUser().getEmail();
    return !assigneeEmail || assigneeEmail === currentUserEmail;
  });
}

// Modify updateSVName to use the locking mechanism
function updateSVName(uniqueId, svName) {
  return withLock(() => withRetry(() => {
    const sheet = getSheet('Sales Dump');

  if (!sheet) {
    Logger.log('❌ Sales Dump sheet not found');
    return { success: false, message: 'Sales Dump sheet not found' };
  }

  try {
    // 🔍 Find row with the given uniqueId
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => String(row[0]).trim() === String(uniqueId).trim());

    if (rowIndex === -1) {
      Logger.log(`❌ No matching row found for ID: ${uniqueId}`);
      return { success: false, message: `No matching row for ID: ${uniqueId}` };
    }

    // ✏️ Update SV Name (Column C) and Assignee (Column S)
    const userEmail = Session.getActiveUser().getEmail();
    sheet.getRange(rowIndex + 1, 3).setValue(svName);
    sheet.getRange(rowIndex + 1, 20).setValue(userEmail);

    // 🔄 Force refresh of the Status formula (Column B)
    const statusRange = sheet.getRange(rowIndex + 1, 2);
    const currentFormula = statusRange.getFormula();
    if (currentFormula) {
      statusRange.setFormula(currentFormula);
    }

    Logger.log(`✅ Updated SV Name to: ${svName} for row: ${rowIndex + 1}`);

    return { success: true, svName: svName, assignee: userEmail };

  } catch (error) {
    Logger.log(`❌ Error in updateSVName: ${error.toString()}`);
    return { success: false, message: `Error: ${error.toString()}` };
  }
  }));
}

function updateLastEditTimestamp() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales Dump');
  const timestampCell = sheet.getRange('U2');
  const currentTime = new Date().getTime();
  timestampCell.setValue(currentTime);
}

function updateDateValidated(uniqueId, dateValidated) {
  return withLock(() => withRetry(() => {
    const sheet = getSheet('Sales Dump');

  if (!sheet) {
    Logger.log('Sales Dump sheet not found');
    return null;
  }

  try {
    const data = sheet.getDataRange().getValues();
    Logger.log('Looking for uniqueId:', uniqueId, 'in Sales Dump');

    const rowIndex = data.findIndex(row => String(row[0]).trim() === String(uniqueId).trim());
    Logger.log('Found row index:', rowIndex);

    if (rowIndex !== -1) {
      // Update Date Validated (Column D)
      sheet.getRange(rowIndex + 1, 4).setValue(dateValidated);
      Logger.log('Updated Date Validated to:', dateValidated, 'for row:', rowIndex + 1);
      
      // Get the current status
      const statusRange = sheet.getRange(rowIndex + 1, 2);
      
      // Return both the status and the updated date validated
      return {
        status: statusRange.getValue(),
        dateValidated: dateValidated
      };
    } else {
      Logger.log('No matching row found for ID:', uniqueId);
      return null;
    }
  } catch (error) {
    Logger.log('Error in updateDateValidated:', error.toString());
    throw error;
  }
  }));
}


function onEdit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const salesDumpSheet = sheet.getSheetByName('Sales Dump');
  const backlogsSheet = sheet.getSheetByName('Backlogs');
  const currentTime = new Date().getTime();

  // Update timestamp for Sales Dump
  if (salesDumpSheet) {
    const timestampCell = salesDumpSheet.getRange('U2');
    timestampCell.setValue(currentTime);
  }

  // Update timestamp for Backlogs
  if (backlogsSheet) {
    const timestampCell = backlogsSheet.getRange('U2');
    timestampCell.setValue(currentTime);
  }
}


function getLastEditTimestamp() {
  const sheet = SpreadsheetApp.openById('1fYx3TXgiKzmkHTwy30aOUSRNChB5h0AUP-f84QklXio').getSheetByName('Sales Dump');
  const timestamp = sheet.getRange('U2').getValue();
  return timestamp;
}

function updateEvalFormSubmissionWLS(formData) {
    const sheetId = '1fYx3TXgiKzmkHTwy30aOUSRNChB5h0AUP-f84QklXio';
    const sheetName = 'Eval Form Submissions WLS';

    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

    if (!sheet) {
        Logger.log('Eval Form Submissions WLS sheet not found');
        return;
    }

    Logger.log('Form Data Received for WLS: ' + JSON.stringify(formData));

    const data = sheet.getDataRange().getValues();
    const phoneColumnIndex = 5; // Phone Number column (F)
    const orderIdColumnIndex = 10; // Order ID column (K)
    const productIdColumnIndex = 9; // Product ID column (J)
    
    let existingRowIndex = -1;
    
    // Check for existing entry by phone number, order ID, and product ID combination
    for (let i = 1; i < data.length; i++) {
        const existingPhone = data[i][phoneColumnIndex];
        const existingOrderId = data[i][orderIdColumnIndex];
        const existingProductId = data[i][productIdColumnIndex];
        
        // Match by phone number AND (order ID OR product ID) to identify same item
        if (existingPhone === formData.phoneNumber && 
            (existingOrderId === formData.orderID || existingProductId === formData.productID)) {
            existingRowIndex = i + 1;
            Logger.log(`Found existing WLS entry at row ${existingRowIndex} for phone: ${formData.phoneNumber}, orderID: ${formData.orderID}, productID: ${formData.productID}`);
            break;
        }
    }

    const rowData = [
        formData.svName,
        formData.dateValidated,
        formData.saleDate,
        formData.xID,
        formData.ban,
        formData.phoneNumber,
        formData.soldBAN,
        formData.soldPhone,
        formData.cssPortfolio,
        formData.productID,
        formData.orderID,
        formData.dueDate,
        formData.multipleSales,
        formData.numberOfSales,
        formData.agentName,
        '',
        formData.callDisposition,
        formData.zviqEvalID,
        formData.selectedFlags["Regulatory"].join(', '),
        formData.selectedFlags["Privacy"].join(', '),
        formData.selectedFlags["Price Information"].join(', '),
        formData.selectedFlags["Plan Information"].join(', '),
        formData.selectedFlags["Order Processing"].join(', '),
        formData.selectedFlags["Account Documentation"].join(', '),
        formData.selectedFlags["Call Summary"].join(', '),
        formData.selectedFlags["Business Intelligence"].join(', '),
        formData.selectedFlags["Marketing Objectives"].join(', '),
        formData.selectedFlags["Professionalism"].join(', '),
        formData.profFlagRemarks,
        formData.goodSaleRemarks
    ];

    if (existingRowIndex !== -1) {
        // Replace existing entry
        sheet.getRange(existingRowIndex, 1, 1, rowData.length).setValues([rowData]);
        Logger.log(`Updated existing WLS entry for ${formData.phoneNumber} at row ${existingRowIndex}`);
        return `WLS Data updated successfully (replaced existing entry for same item)`;
    } else {
        // Add new entry
        const newRowIndex = sheet.getLastRow() + 1;
        sheet.getRange(newRowIndex, 1, 1, rowData.length).setValues([rowData]);
        Logger.log(`Inserted new WLS entry for ${formData.phoneNumber} at row ${newRowIndex}`);
        return `WLS Data updated successfully (new entry created)`;
    }
}

function updateEvalFormSubmissionFFH(formData) {
    const sheetId = '1fYx3TXgiKzmkHTwy30aOUSRNChB5h0AUP-f84QklXio';
    const sheetName = 'Eval Form Submissions FFH';

    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

    if (!sheet) {
        Logger.log('Eval Form Submissions FFH sheet not found');
        return;
    }

    Logger.log('Form Data Received for FFH: ' + JSON.stringify(formData));

    const data = sheet.getDataRange().getValues();
    const phoneColumnIndex = 5; // Phone Number column (F)
    const orderIdColumnIndex = 10; // Order ID column (K)
    const productIdColumnIndex = 9; // Product ID column (J)
    
    let existingRowIndex = -1;
    
    // Check for existing entry by phone number, order ID, and product ID combination
    for (let i = 1; i < data.length; i++) {
        const existingPhone = String(data[i][phoneColumnIndex] || '').trim();
        const existingOrderId = String(data[i][orderIdColumnIndex] || '').trim();
        const existingProductId = String(data[i][productIdColumnIndex] || '').trim();
        
        const formPhone = String(formData.phoneNumber || '').trim();
        const formOrderId = String(formData.orderID || '').trim();
        const formProductId = String(formData.productID || '').trim();
        
        Logger.log(`FFH Row ${i}: Comparing phone '${existingPhone}' vs '${formPhone}', orderID '${existingOrderId}' vs '${formOrderId}', productID '${existingProductId}' vs '${formProductId}'`);
        
        // Match by phone number AND (order ID OR product ID) to identify same item
        if (existingPhone === formPhone && 
            (existingOrderId === formOrderId || existingProductId === formProductId)) {
            existingRowIndex = i + 1;
            Logger.log(`Found existing FFH entry at row ${existingRowIndex} for phone: ${formPhone}, orderID: ${formOrderId}, productID: ${formProductId}`);
            break;
        }
    }

    const rowData = [
        formData.svName,
        formData.dateValidated,
        formData.saleDate,
        formData.xID,
        formData.ban,
        formData.phoneNumber,
        formData.soldBAN,
        formData.soldPhone,
        formData.cssPortfolio,
        formData.productID,
        formData.orderID,
        formData.dueDate,
        formData.multipleSales,
        formData.numberOfSales,
        formData.agentName,
        '',
        formData.callDisposition,
        formData.zviqEvalID,
        formData.selectedFlags["Regulatory"].join(', '),
        formData.selectedFlags["Privacy"].join(', '),
        formData.selectedFlags["Price Information"].join(', '),
        formData.selectedFlags["Plan Information"].join(', '),
        formData.selectedFlags["Order Processing"].join(', '),
        formData.selectedFlags["Account Documentation"].join(', '),
        formData.selectedFlags["Call Summary"].join(', '),
        formData.selectedFlags["Customer Experience"].join(', '),
        formData.selectedFlags["Business Intelligence"].join(', '),
        formData.selectedFlags["Marketing Objectives"].join(', '),
        formData.selectedFlags["Professionalism"].join(', '),
        formData.profFlagRemarks,
        formData.goodSaleRemarks
    ];

    if (existingRowIndex !== -1) {
        // Replace existing entry
        sheet.getRange(existingRowIndex, 1, 1, rowData.length).setValues([rowData]);
        Logger.log(`Updated existing FFH entry for ${formData.phoneNumber} at row ${existingRowIndex}`);
        return `FFH Data updated successfully (replaced existing entry for same item)`;
    } else {
        // Add new entry
        const newRowIndex = sheet.getLastRow() + 1;
        sheet.getRange(newRowIndex, 1, 1, rowData.length).setValues([rowData]);
        Logger.log(`Inserted new FFH entry for ${formData.phoneNumber} at row ${newRowIndex}`);
        return `FFH Data updated successfully (new entry created)`;
    }
}

function saveFormData(formType, uniqueId, formData) {
  const userEmail = Session.getActiveUser().getEmail();
  const userProperties = PropertiesService.getUserProperties();
  const key = `${userEmail}_${formType}_${uniqueId}`;
  userProperties.setProperty(key, JSON.stringify(formData));
  return "Data saved successfully";
}

function getRealtimeDataFromSheet(uniqueId, formType) {
    Logger.log(`Starting getRealtimeDataFromSheet - uniqueId: ${uniqueId}, formType: ${formType}`);
    
    try {
        const sheetId = '1fYx3TXgiKzmkHTwy30aOUSRNChB5h0AUP-f84QklXio';
        const evalSheetName = formType === 'WLS' ? 'Eval Form Submissions WLS' : 'Eval Form Submissions FFH';
        const evalSheet = SpreadsheetApp.openById(sheetId).getSheetByName(evalSheetName);

        if (!evalSheet) {
            Logger.log(`${evalSheetName} sheet not found`);
            return null;
        }

        const evalData = evalSheet.getDataRange().getValues();
        const phoneNumberIndex = 5; // Column F in Eval sheets

        // Find the matching row in eval sheet
        for (let i = 1; i < evalData.length; i++) {
            if (evalData[i][phoneNumberIndex] === uniqueId) {
                Logger.log(`Found matching eval data for ${uniqueId}`);

                const flagsStartIndex = 18;
                const flagsEndIndex = formType === 'WLS' ? 28 : 29;
                const flagsArray = evalData[i].slice(flagsStartIndex, flagsEndIndex);

                // Parse flags based on form type
                const flags = formType === 'WLS' 
                    ? parseFlags(flagsArray)
                    : parseFlagsFFH(flagsArray);

                // Return structured data
                return {
                    callDisposition: evalData[i][16] || '', // Column Q
                    zviqEvalID: evalData[i][17] || '', // Column R
                    selectedFlags: flags,
                    profFlagRemarks: formType === 'WLS' ? (evalData[i][28] || '') : (evalData[i][29] || ''),
                    goodSaleRemarks: formType === 'WLS' ? (evalData[i][29] || '') : (evalData[i][30] || '')
                };
            }
        }

        Logger.log(`No eval data found for ${uniqueId}`);
        return null;

    } catch (error) {
        Logger.log(`Error in getRealtimeDataFromSheet: ${error.toString()}`);
        Logger.log(`Stack: ${error.stack}`);
        return null;
    }
}

function clearEvalFormData(uniqueId, formType = 'WLS') {
    const sheetId = '1fYx3TXgiKzmkHTwy30aOUSRNChB5h0AUP-f84QklXio';
    const sheetName = `Eval Form Submissions ${formType}`;
    const salesDumpSheet = SpreadsheetApp.openById(sheetId).getSheetByName('Sales Dump');
    const evalSheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

    if (!evalSheet || !salesDumpSheet) {
        Logger.log(`${sheetName} sheet or Sales Dump sheet not found`);
        return "Sheet not found";
    }

    // Clear data in Sales Dump sheet
    const salesData = salesDumpSheet.getDataRange().getValues();
    const phoneColumnIndex = 6; // Phone Number column index
    
    for (let i = 1; i < salesData.length; i++) {
        if (salesData[i][phoneColumnIndex] === uniqueId) {
            // Clear all columns except the phone number and unique identifier
            const row = i + 1;
            salesDumpSheet.getRange(row, 3).clearContent(); // SV Name
            salesDumpSheet.getRange(row, 4).clearContent(); // Date Validated
            salesDumpSheet.getRange(row, 19).clearContent(); // Clear assignee email
            break;
        }
    }

    // Clear data in Eval Form sheet
    const evalData = evalSheet.getDataRange().getValues();
    const evalPhoneColumnIndex = 5; // Phone Number column index in eval sheet
    
    for (let i = 1; i < evalData.length; i++) {
        if (evalData[i][evalPhoneColumnIndex] === uniqueId) {
            const row = i + 1;
            // Clear all columns in the eval form
            evalSheet.getRange(row, 1, 1, evalSheet.getLastColumn()).clearContent();
            break;
        }
    }

    return `Data cleared successfully for ${uniqueId}`;
}

function handleFormSubmission(formType, formData) {
  try {
    saveFormData(formType, formData.phoneNumber, formData);

    if (formType === 'WLS') {
      return updateEvalFormSubmissionWLS(formData);
    } else if (formType === 'FFH') {
      return updateEvalFormSubmissionFFH(formData);
    }
  } catch (error) {
    Logger.log(`Error in handleFormSubmission for ${formType}: ` + error.toString());
    throw error;
  }
}

function getSVRoster() {
  const sheetId = '1fYx3TXgiKzmkHTwy30aOUSRNChB5h0AUP-f84QklXio';
  const svRosterSheetName = 'SV Roster';

  try {
    const svRosterSheet = SpreadsheetApp.openById(sheetId).getSheetByName(svRosterSheetName);
    
    if (!svRosterSheet) {
      Logger.log('SV Roster sheet not found');
      return [];
    }

    const svRosterData = svRosterSheet.getRange('B2:C').getValues();
    const svRosterNames = svRosterData
      .filter(row => row[1] === 'WLS' || row[1] === 'FFH')
      .map(row => ({ name: row[0], type: row[1] }))
      .filter(item => item.name);

    return svRosterNames;
  } catch (error) {
    Logger.log('Error in getSVRoster: ' + error.toString());
    throw error;
  }
}

function getUserInfo() {
  const userEmail = Session.getActiveUser().getEmail();
  return {
    email: userEmail,
  };
}

function refreshData() {
  try {
    const newData = getSalesData();
    return newData;
  } catch (error) {
    Logger.log('Error in refreshData: ' + error.toString());
    throw error;
  }
}

function logError(errorMessage, functionName) {
  const timestamp = new Date().toISOString();
  Logger.log(`[${timestamp}] Error in ${functionName}: ${errorMessage}`);
}

function getDNCData() {
  return withRetry(() => {
    Logger.log('Starting getDNCData function');
    const sheet = getSheet('DNC');
    
    if (!sheet) {
      Logger.log('DNC sheet not found');
      return {
        error: 'DNC sheet not found',
        data: [],
        ids: [],
        userEmail: Session.getActiveUser().getEmail(),
        timestamp: new Date().getTime()
      };
    }

    Logger.log('DNC sheet found');
    const dataRange = sheet.getDataRange();
    const rawData = dataRange.getValues();
    Logger.log('Raw DNC Data rows: ' + rawData.length);
    Logger.log('Headers: ' + JSON.stringify(rawData[0])); // Log headers

    if (rawData.length <= 1) {
      Logger.log('DNC sheet is empty or contains only headers');
      return {
        data: [],
        ids: [],
        userEmail: Session.getActiveUser().getEmail(),
        timestamp: new Date().getTime()
      };
    }

    // Process data starting from row 1 (skipping headers)
    const processedData = rawData.slice(1).map((row, index) => {
      try {
        const formattedDate = row[2] ? 
          (row[2] instanceof Date ? 
            Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : 
            row[2]) : '';

        return [
          String(row[0] || ''),         // Unique ID
          String(row[1] || ''),         // SV Name
          formattedDate,                // Call Date
          String(row[3] || ''),         // Status
          String(row[4] || ''),         // xID
          String(row[5] || ''),         // Campaign
          String(row[6] || ''),         // BAN
          String(row[7] || ''),         // PHONE1
          String(row[8] || ''),         // Category
          String(row[9] || ''),         // Sub-Category
          String(row[10] || ''),        // Portfolio
          String(row[11] || ''),        // DNC Driver
          String(row[12] || ''),        // Agent Name
          String(row[13] || '')         // FLM Name
        ];
      } catch (error) {
        Logger.log(`Error processing row ${index + 1}: ${error}`);
        Logger.log(`Problematic row data: ${JSON.stringify(row)}`);
        // Return an empty row in case of error
        return Array(14).fill('');
      }
    });

    const ids = rawData.slice(1).map(row => String(row[0] || ''));
    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = sheet.getRange('U2').getValue() || new Date().getTime();

    Logger.log('Processed DNC Data rows: ' + processedData.length);
    Logger.log('Sample processed row: ' + JSON.stringify(processedData[0]));

    return {
      data: processedData,
      ids: ids,
      userEmail: userEmail,
      timestamp: timestamp
    };
  });
}


// Update the updateDNCField function to match the actual column structure
function updateDNCFields(uniqueId, updates) {
    return withLock(() => withRetry(() => {
        const sheet = getSheet('DNC');

    if (!sheet) {
        Logger.log('DNC sheet not found');
        return { success: false, message: 'DNC sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => String(row[0]).trim() === String(uniqueId).trim());

    if (rowIndex === -1) {
        Logger.log(`No matching row found for ID: ${uniqueId}`);
        return { success: false, message: `No matching row for ID: ${uniqueId}` };
    }

    // Column mappings for fields
    const columnMap = {
        category: 8,       // Adjust for the actual column index of "Category"
        subcategory: 9,    // Adjust for the actual column index of "Sub-Category"
        dncDriver: 11      // Adjust for the actual column index of "DNC Driver"
    };

    Object.entries(updates).forEach(([field, value]) => {
        const columnIndex = columnMap[field];
        if (columnIndex !== undefined) {
            sheet.getRange(rowIndex + 1, columnIndex + 1).setValue(value || ''); // Clear if null/empty
        }
    });

        Logger.log(`Batch updated fields for ID: ${uniqueId} with values: ${JSON.stringify(updates)}`);
        return { success: true };
    }));
}

 function onEdit(e) {
  const sheet = e.source.getSheetByName('DNC');  // Ensure we're working with the correct sheet
  const range = e.range;
  const columnToWatch = 8; // Column I (Category) — Adjust this for any column you want to watch

  // Check if the edited column is the Category column (Column 8)
  if (range.getColumn() === columnToWatch) {
    const row = range.getRow();
    const timestampCell = sheet.getRange(row, 20); // Column T for the timestamp
    timestampCell.setValue(new Date());  // Set the current timestamp whenever Category is updated

    // Now check if the category is cleared (empty)
    const categoryValue = sheet.getRange(row, columnToWatch).getValue();

    // If the category is cleared, reset the dropdowns for Category, Sub-Category, and DNC Driver
    if (!categoryValue) {
      const subCategoryCell = sheet.getRange(row, 9);  // Column J (Sub-Category)
      const dncDriverCell = sheet.getRange(row, 11);   // Column L (DNC Driver)

      subCategoryCell.setValue('');  // Reset Sub-Category
      dncDriverCell.setValue('');    // Reset DNC Driver
    }
  }
}







