function rebuildSetup() {
    var ss = SpreadsheetApp.openById('1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw');
      var sheet = ss.getSheetByName('工作表1');
        
          // Step 1: Set columns A, B, C, F to plain text
            var maxRows = 60000;
              sheet.getRange(1, 1, maxRows, 1).setNumberFormat('@'); // A
                sheet.getRange(1, 2, maxRows, 1).setNumberFormat('@'); // B
                  sheet.getRange(1, 3, maxRows, 1).setNumberFormat('@'); // C
                    sheet.getRange(1, 6, maxRows, 1).setNumberFormat('@'); // F
                      
                        // Step 2: Clear data rows (keep row 1 header)
                          var lastRow = sheet.getLastRow();
                            if (lastRow > 1) {
                                sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
                                  }
                                    
                                      SpreadsheetApp.flush();
                                        Logger.log('Setup done: A/B/C/F formatted as text, cleared rows 2-' + lastRow);
                                        }