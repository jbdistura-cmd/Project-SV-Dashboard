// Trigger setup for real-time updates instead of polling
function setupTriggers() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create onEdit trigger for Sales Dump sheet
  ScriptApp.newTrigger('onSalesDumpEdit')
    .forSpreadsheet(SPREADSHEET_ID)
    .onEdit()
    .create();
    
  // Create onChange trigger for data changes
  ScriptApp.newTrigger('onDataChange')
    .forSpreadsheet(SPREADSHEET_ID)
    .onChange()
    .create();
}

// Handle edit events on Sales Dump sheet
function onSalesDumpEdit(e) {
  if (!e || !e.source) return;
  
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  
  // Only process relevant sheets
  if (['Sales Dump', 'Backlogs', 'DNC'].includes(sheetName)) {
    updateLastEditTimestamp();
  }
}

// Handle data changes
function onDataChange(e) {
  if (!e || !e.source) return;
  updateLastEditTimestamp();
}

// Optimized timestamp update
function updateLastEditTimestamp() {
  withRetry(() => {
    const timestamp = new Date().getTime();
    
    // Update timestamps for all relevant sheets
    ['Sales Dump', 'Backlogs', 'DNC'].forEach(sheetName => {
      try {
        const sheet = getSheet(sheetName);
        if (sheet) {
          sheet.getRange('U2').setValue(timestamp);
        }
      } catch (error) {
        Logger.log(`Error updating timestamp for ${sheetName}: ${error}`);
      }
    });
  });
}
