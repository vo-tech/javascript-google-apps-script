function onOpen() {
    // Add a custom menu
    SpreadsheetApp.getUi()
      .createMenu('PDF Tools')
      .addItem('Generate PDF', 'generatePdf')
      .addToUi();
  }
  
  function generatePdf() {
    try {
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = spreadsheet.getActiveSheet();
      
      // Columns to hide: A, E, F, H, K, M, N
      var columnsToHide = [1, 5, 6, 8, 11, 13, 14];
      
      // Store original column visibility
      var originalVisibility = [];
      for (var i = 0; i < columnsToHide.length; i++) {
        var col = columnsToHide[i];
        originalVisibility.push(sheet.isColumnHiddenByUser(col));
        sheet.hideColumns(col);
      }
      
      // Prepare export options
      var exportOptions = {
        'gid': sheet.getSheetId(),
        'exportFormat': 'application/pdf',
        'size': 'A4',
        'portrait': 1,
        'gridlines': false,
        'printtitle': true,
        'sheetnames': true,
        'pagenum': 'CENTER',
        'scale': 4,
        'top_margin': 0.5,
        'bottom_margin': 0.5,
        'left_margin': 0.5,
        'right_margin': 0.5
      };
      
      // Export the sheet as PDF using Drive API
      var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() +
                '/export?' +
                'format=pdf' +
                '&gid=' + sheet.getSheetId() +
                '&size=A4' +
                '&portrait=1' +
                '&gridlines=false' +
                '&printtitle=true' +
                '&sheetnames=true' +
                '&pagenum=CENTER' +
                '&scale=4' +
                '&top_margin=0.50' +
                '&bottom_margin=0.50' +
                '&left_margin=0.50' +
                '&right_margin=0.50';
      
      var token = ScriptApp.getOAuthToken();
      var response = UrlFetchApp.fetch(url, {
        headers: {
          'Authorization': 'Bearer ' + token
        },
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() == 200) {
        var pdfBlob = response.getBlob().setName(sheet.getName() + '.pdf');
        var pdfFile = DriveApp.createFile(pdfBlob);
        
        // Restore original column visibility
        for (var i = 0; i < columnsToHide.length; i++) {
          var col = columnsToHide[i];
          if (originalVisibility[i]) {
            sheet.hideColumns(col);
          } else {
            sheet.showColumns(col);
          }
        }
        
        // Notify user
        var pdfUrl = pdfFile.getUrl();
        SpreadsheetApp.getUi().alert('PDF generated successfully:\n' + pdfUrl);
      } else {
        throw new Error('Failed to generate PDF: ' + response.getContentText());
      }
      
    } catch (e) {
      Logger.log("Error: " + e.toString());
      SpreadsheetApp.getUi().alert('An error occurred: ' + e.toString());
    }
  }