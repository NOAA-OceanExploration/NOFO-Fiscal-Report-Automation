function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Create a new menu in the Google Docs UI.
  ui.createMenu('Figure Generators')
    .addItem("Generate All Figures", "genAll")
    .addItem('Generate Figure One', 'figureOne')
    .addItem('Generate Figure Two', 'figureTwo')
    .addItem('Generate Figure Three', 'figureThree')
    .addItem('Generate Figure Four', 'figureFour')
    .addItem('Generate Figure Five', 'figureFive')
    .addItem('Generate Figure Six', 'figureSix')
    .addItem('Clear Spreadsheet', 'clearSpreadsheet')
    .addToUi();
}

function figureOne() {
var originalDocId = '17WNH_DGlP9pvsFj4v4EkQnrRv2ygJhR7kxUKliq0qYI';
  var newDocId = '1YoalnLbU7oKYKfQu-mB4KM-xqVgOEhmafXcdzI4QwHw';
  // Access the original spreadsheet
  var originalSpreadsheet = SpreadsheetApp.openById(originalDocId);
  var originalSheet = originalSpreadsheet.getSheetByName('Calculations');

  var originalColumnGroups = [['H2:M2'], ['P2:U2'], ['X2:AC2'], ['AN2:AS2']];

  // Access the new spreadsheet
  var newSpreadsheet = SpreadsheetApp.openById(newDocId);
  var newSheet = newSpreadsheet.getSheetByName('Sheet2');
  
    // Set the headers in the new spreadsheet
  var columnNames = ['Academia', 'Federal: NOAA', 'Federal: Other', 'Nongovernmental', 'Private Sector', 'State'];
  var rowNames = ['Pre-Proposals', 'Encouraged', 'Submitted', 'Awarded'];

  for (var i = 0; i < columnNames.length; i++) {
    newSheet.getRange(1, 1 + i).setValue(columnNames[i]);
  }

  // Copy data from original to new spreadsheet
  for (var i = 0; i < rowNames.length; i++) {
    newSheet.getRange(2 + i, 1).setValue(rowNames[i]);

    var columnValues = originalSheet.getRange(originalColumnGroups[i][0]).getValues();
    newSheet.getRange(2 + i, 2, 1, columnValues[0].length).setValues(columnValues);
  }

  var rowNumber = 1; // row number to be shifted
  var lastColumn = newSheet.getLastColumn(); // get last column
  var rowData = newSheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0]; // get the data in the row
  
  rowData.unshift("Figure 1"); // insert a blank cell at the beginning of the data
  rowData.pop(); // remove the last element to keep the row length consistent
  
  newSheet.getRange(rowNumber, 1, 1, lastColumn).setValues([rowData]); // write the data back into the row

  // Create a chart
  var chartBuilder = newSheet.newChart();
  chartBuilder.addRange(newSheet.getRange('A1:G5'))
              .setChartType(Charts.ChartType.COLUMN)
              .setOption('series', {
                  0: {color: 'blue'},
                  1: {color: 'lightblue'},
                  2: {color: 'darkblue'},
                  3: {color: '#ADD8E6'},
                  4: {color: '#0000FF'},
                  5: {color: '#00008B'}
              }) 
              .setOption('legend', {position: 'right'})
              .setNumHeaders(1)
              .setPosition(6, 1, 0, 0);

  var chart = chartBuilder.build();
  newSheet.insertChart(chart);
}

function figureTwo() {
  var originalDocId = '17WNH_DGlP9pvsFj4v4EkQnrRv2ygJhR7kxUKliq0qYI';
  var newDocId = '1YoalnLbU7oKYKfQu-mB4KM-xqVgOEhmafXcdzI4QwHw';
  // Access the original spreadsheet
  var originalSpreadsheet = SpreadsheetApp.openById(originalDocId);
  var originalSheet = originalSpreadsheet.getSheetByName('Calculations');

  // Access the new spreadsheet
  var newSpreadsheet = SpreadsheetApp.openById(newDocId);
  var newSheet = newSpreadsheet.getSheetByName('Sheet2');

  // Set the headers in row 28 of the new spreadsheet
  newSheet.getRange('A28').setValue('Figure 2');
  newSheet.getRange('B28').setValue('Pre-Proposals');
  newSheet.getRange('C28').setValue('Encouraged');
  newSheet.getRange('D28').setValue('Full Proposals');
  newSheet.getRange('E28').setValue('Awards');

  // Copy column A from original to new spreadsheet
  var columnAValues = originalSheet.getRange('A2:A6').getValues();
  var convertedValues = columnAValues.map(function(row) { return [Number(row[0])]; }); // Convert to numbers
  newSheet.getRange('A29:A33').setValues(convertedValues);

  // Let's assume your headers are "Header1", "Header2", "Header3", "Header4".
  var originalHeaders = ['# Preproposals', '# Enc Preproposals', '# Full Proposals', '# Awarded'];

  // Get the first row which contains the headers.
  var headers = originalSheet.getRange(1, 1, 1, originalSheet.getLastColumn()).getValues()[0];

  // Find the columns corresponding to the headers.
  var originalColumns = [];
  for (var j = 0; j < originalHeaders.length; j++) {
    var colIndex = headers.indexOf(originalHeaders[j]) + 1; // +1 because column index starts from 1
    if (colIndex !== 0) {
      originalColumns.push(colIndex);
    } else {
      Logger.log("Header not found: " + originalHeaders[j]);
    }
  }

  // Now originalColumns contains column numbers instead of letters. The rest of the code remains the same.
  var columnValues;
  for (var i = 0; i < originalColumns.length; i++) {
    columnValues = originalSheet.getRange(2, originalColumns[i], 5).getValues(); // start from row 2, for 5 rows.
    newSheet.getRange(29, 2 + i, 5).setValues(columnValues); // Columns B to E
  }
  
  // Set number format for columns B to E in new spreadsheet
  for (var i = 0; i < 4; i++) {
    newSheet.getRange(29, 2 + i, 5).setNumberFormat('0'); // Set to whole number format
  }

  // Sort the data in ascending order based on column A (Figure 2) values
  var dataRange = newSheet.getRange('A29:E33');
  dataRange.sort(1);

  // Create a chart
  var chartBuilder = newSheet.newChart();
  chartBuilder.addRange(newSheet.getRange('A28:E33')) // Use only the data range, excluding headers
              .setChartType(Charts.ChartType.COLUMN)
              .setOption('hAxis.title', 'Fiscal Year')
              .setOption('series', {
                  0: {color: 'blue'},
                  1: {color: 'lightblue'},
                  2: {color: 'darkblue'},
                  3: {color: 'blue'}
              }) 
              .setOption('legend', {position: 'right'})
              .setNumHeaders(1)
              .setPosition(34, 1, 0, 0);

  var chart = chartBuilder.build();
  newSheet.insertChart(chart);
}

function figureThree() {
  var originalDocId = '17WNH_DGlP9pvsFj4v4EkQnrRv2ygJhR7kxUKliq0qYI';
  var newDocId = '1YoalnLbU7oKYKfQu-mB4KM-xqVgOEhmafXcdzI4QwHw';

  // Access the original spreadsheet
  var originalSpreadsheet = SpreadsheetApp.openById(originalDocId);
  var originalSheet = originalSpreadsheet.getSheetByName('Calculations');

  // Access the new spreadsheet
  var newSpreadsheet = SpreadsheetApp.openById(newDocId);
  var newSheet = newSpreadsheet.getSheetByName('Sheet2');

  // Set the headers in row 56 of the new spreadsheet
  newSheet.getRange('A56').setValue('Figure 3');
  newSheet.getRange('B56').setValue('Ocean Exploration');
  newSheet.getRange('C56').setValue('Maritime Heritage');
  newSheet.getRange('D56').setValue('Technology');

  // Copy column A from original to new spreadsheet
  var columnAValues = originalSheet.getRange('A2:A6').getValues();
  var convertedValues = columnAValues.map(function(row) { return [Number(row[0])]; }); // Convert to numbers
  newSheet.getRange('A57:A61').setValues(convertedValues);

  // Get header row from the original sheet
  var headers = originalSheet.getRange(1, 1, 1, originalSheet.getLastColumn()).getValues()[0];

  // Header names of columns to be copied from original spreadsheet
  var headersOfInterest = ['# Exp Award', '# Arch Award', '# Tech Award'];

  headersOfInterest.forEach(function(header, i) {
    var originalColumnNumber = headers.indexOf(header) + 1; // +1 because column numbering starts from 1

    if (originalColumnNumber > 0) { // If header is found
      var columnValues = originalSheet.getRange(2, originalColumnNumber, 5).getValues();
      newSheet.getRange(57, 2 + i, 5).setValues(columnValues);
    }
  });

  // Set number format for columns B to D in new spreadsheet
  for (var i = 0; i < 3; i++) {
    newSheet.getRange(57, 2 + i, 5).setNumberFormat('0'); // Set to whole number format
  }

  // Sort the data in ascending order based on column A (Figure 3) values
  var dataRange = newSheet.getRange('A57:D61');
  dataRange.sort(1);

  // Create a chart
  var chartBuilder = newSheet.newChart();
  chartBuilder.addRange(newSheet.getRange('A56:D61')) // Use only the data range, excluding headers
              .setChartType(Charts.ChartType.COLUMN)
              .setOption('hAxis.title', 'Fiscal Year')
              .setOption('series', {
                  0: {color: 'blue', label: 'Ocean Exploration'},
                  1: {color: 'lightblue', label: 'Maritime Heritage'},
                  2: {color: 'darkblue', label: 'Technology'}
              }) 
              .setOption('legend', {position: 'right'})
              .setNumHeaders(1)
              .setPosition(62, 1, 0, 0);

  var chart = chartBuilder.build();
  newSheet.insertChart(chart);
}


function figureFour() {
  var originalDocId = '17WNH_DGlP9pvsFj4v4EkQnrRv2ygJhR7kxUKliq0qYI';
  var newDocId = '1YoalnLbU7oKYKfQu-mB4KM-xqVgOEhmafXcdzI4QwHw';

  // Access the original spreadsheet
  var originalSpreadsheet = SpreadsheetApp.openById(originalDocId);
  var originalSheet = originalSpreadsheet.getSheetByName('Calculations');

  // Access the new spreadsheet
  var newSpreadsheet = SpreadsheetApp.openById(newDocId);
  var newSheet = newSpreadsheet.getSheetByName('Sheet2');

  // Set the headers in row 84 of the new spreadsheet
  newSheet.getRange('A84').setValue('Figure 4');
  newSheet.getRange('B84').setValue('Academia');
  newSheet.getRange('C84').setValue('Federal: NOAA');
  newSheet.getRange('D84').setValue('Federal: Other');
  newSheet.getRange('E84').setValue('Nongovernmental');
  newSheet.getRange('F84').setValue('Private Sector');

  // Copy column A from original to new spreadsheet
  var columnAValues = originalSheet.getRange('A2:A6').getValues();
  var convertedValues = columnAValues.map(function(row) { return [Number(row[0])]; }); // Convert to numbers
  newSheet.getRange('A85:A89').setValues(convertedValues);

  // Get header row from the original sheet
  var headers = originalSheet.getRange(1, 1, 1, originalSheet.getLastColumn()).getValues()[0];

  // Header names of columns to be copied from original spreadsheet
  var headersOfInterest = ['PI Academic', 'PI NOAA', 'PI NOAA', 'PI NGO', 'PI Private'];

  headersOfInterest.forEach(function(header, i) {
    var originalColumnNumber = headers.indexOf(header) + 1; // +1 because column numbering starts from 1

    if (originalColumnNumber > 0) { // If header is found
      var columnValues = originalSheet.getRange(2, originalColumnNumber, 5).getValues();
      newSheet.getRange(85, 2 + i, 5).setValues(columnValues);
    }
  });

  // Set number format for columns B to F in new spreadsheet
  for (var i = 0; i < 5; i++) {
    newSheet.getRange(85, 2 + i, 5).setNumberFormat('$#,##0.00');
  }

  // Sort the data in ascending order based on column A (Figure 4) values
  var dataRange = newSheet.getRange('A85:F89');
  dataRange.sort(1);

  var dataRange = newSheet.getRange('A84:F89');
  // Create a chart
  var chart = newSheet.newChart()
     .setChartType(Charts.ChartType.COLUMN) // change to column chart
     .addRange(dataRange)
     .setPosition(90, 1, 0, 0)
     .setOption('hAxis.title', 'Fiscal Year')
     .setOption('legend', {position: 'right'})
    .setOption('series', {0: {color: 'blue'}, 1: {color: 'lightblue'}, 2: {color: 'darkblue'}, 3: {color: 'skyblue'}, 4: {color: 'steelblue'}})
    .setNumHeaders(1)
    .build();

  newSheet.insertChart(chart);
}

function figureFive() {
  var originalDocId = '17WNH_DGlP9pvsFj4v4EkQnrRv2ygJhR7kxUKliq0qYI';
  var newDocId = '1YoalnLbU7oKYKfQu-mB4KM-xqVgOEhmafXcdzI4QwHw';

  // Access the original spreadsheet
  var originalSpreadsheet = SpreadsheetApp.openById(originalDocId);
  var originalSheet = originalSpreadsheet.getSheetByName('Calculations');

  // Access the new spreadsheet
  var newSpreadsheet = SpreadsheetApp.openById(newDocId);
  var newSheet = newSpreadsheet.getSheetByName('Sheet2');

  // Set the headers in row 112 of the new spreadsheet
  newSheet.getRange('A112').setValue('Figure 5');
  newSheet.getRange('B112').setValue('Academia');
  newSheet.getRange('C112').setValue('Federal: NOAA');
  newSheet.getRange('D112').setValue('Federal: Other');
  newSheet.getRange('E112').setValue('Nongovernmental');
  newSheet.getRange('F112').setValue('Private Sector');

  // Copy column A from original to new spreadsheet
  var columnAValues = originalSheet.getRange('A2:A6').getValues();
  newSheet.getRange('A113:A117').setValues(columnAValues);

  // Get header row from the original sheet
  var headers = originalSheet.getRange(1, 1, 1, originalSheet.getLastColumn()).getValues()[0];

  // Header names of columns to be copied from original spreadsheet
  var headersOfInterest = ['$ Award Acad', '$ Award NOAA', '$ Award Other', '$ Award NGO', '$ Award Priv'];

  headersOfInterest.forEach(function(header, i) {
    var originalColumnNumber = headers.indexOf(header) + 1; // +1 because column numbering starts from 1

    if (originalColumnNumber > 0) { // If header is found
      var columnValues = originalSheet.getRange(2, originalColumnNumber, 5).getValues();
      newSheet.getRange(113, 2 + i, 5).setValues(columnValues);
    }
  });

  // Set number format for columns B to F in new spreadsheet
  for (var i = 0; i < 5; i++) {
    newSheet.getRange(113, 2 + i, 5).setNumberFormat('$#,##0.00');
  }

  // Sort the data in ascending order based on column A (Figure 5) values
  var dataRange = newSheet.getRange('A113:F117');
  dataRange.sort(1);

  var dataRange = newSheet.getRange('A112:F117');
  // Create a chart
  var chart = newSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN) // change to column chart
    .addRange(dataRange)
    .setPosition(118, 1, 0, 0)
    .setOption('hAxis.title', 'Fiscal Year')
    .setOption('legend', {position: 'right'})
    .setOption('series', {0: {color: 'blue'}, 1: {color: 'lightblue'}, 2: {color: 'darkblue'}, 3: {color: 'skyblue'}, 4: {color: 'steelblue'}})
    .setNumHeaders(1)
    .build();

  newSheet.insertChart(chart);
}

function figureSix() {
  var originalDocId = '17WNH_DGlP9pvsFj4v4EkQnrRv2ygJhR7kxUKliq0qYI';
  var newDocId = '1YoalnLbU7oKYKfQu-mB4KM-xqVgOEhmafXcdzI4QwHw';
  // Access the original spreadsheet
  var originalSpreadsheet = SpreadsheetApp.openById(originalDocId);
  var originalSheet = originalSpreadsheet.getSheetByName('Calculations');

  // Access the new spreadsheet
  var newSpreadsheet = SpreadsheetApp.openById(newDocId);
  var newSheet = newSpreadsheet.getSheetByName('Sheet2');

  // Set the headers in row 140 of the new spreadsheet
  newSheet.getRange('A140').setValue('Figure 6');
  newSheet.getRange('B140').setValue('Awarded');
  newSheet.getRange('C140').setValue('Leveraged');

  // Copy column A from original to new spreadsheet
  var columnAValues = originalSheet.getRange('A2:A6').getValues();
  newSheet.getRange('A141:A145').setValues(columnAValues);

  var headerRow = originalSheet.getRange(1, 1, 1, originalSheet.getLastColumn()).getValues()[0];

  var headerIndexStart = -1;

  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === "$ Award Acad") {
      headerIndexStart = i + 1;  // Add 1 because column indices are 1-based in Apps Script
      break;
    }
  }

  var headerIndexEnd = -1;

  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === "$ Award Tribal") {
      headerIndexEnd = i + 1;  // Add 1 because column indices are 1-based in Apps Script
      break;
    }
  }

  Logger.log(headerIndexStart)
  Logger.log(headerIndexEnd)

  // Calculate sum of columns BZ to CG (columns 78 to 87) for each row and copy to column B of new spreadsheet
  for (var i = 2; i <= 6; i++) {
    var sum = 0;
    for (var j = headerIndexStart; j <= headerIndexEnd; j++) {
      sum += originalSheet.getRange(i, j).getValue();
    }
    newSheet.getRange(i + 139, 2).setValue(sum); // row number in new sheet is original row number + 139
  }

  // Copy column BO (column 68) from original to column C of new spreadsheet
  var columnBOValues = originalSheet.getRange('BO2:BO6').getValues();
  newSheet.getRange('C141:C145').setValues(columnBOValues);

  // Set number format for columns B and C in new spreadsheet
  newSheet.getRange('B141:B145').setNumberFormat('$#,##0.00');
  newSheet.getRange('C141:C145').setNumberFormat('$#,##0.00');

  var dataRange = newSheet.getRange('A141:C145');
  dataRange.sort(1);

  var dataRange = newSheet.getRange('A140:C145');
  // Create a chart
  var chart = newSheet.newChart()
     .setChartType(Charts.ChartType.COLUMN) // change to column chart
     .addRange(dataRange)
     .setPosition(146, 1, 0, 0)
     .setOption('hAxis.title', 'Fiscal Year')
     .setOption('legend', {position: 'right'})
     .setOption('series', {0: {color: 'blue'}, 1: {color: 'lightblue'}}) // No need to specify targetAxisIndex
     .setNumHeaders(1)
     .build();

  newSheet.insertChart(chart);
}

function genAll() {
  figureOne()
  figureTwo()
  figureThree()
  figureFour()
  figureFive()
  figureSix()
}

function clearSpreadsheet() {
  var spreadsheetId = '1YoalnLbU7oKYKfQu-mB4KM-xqVgOEhmafXcdzI4QwHw'; // replace with your Spreadsheet ID
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName('Sheet2');

  sheet.clear(); // this will clear the content, formats, and comments
  var charts = sheet.getCharts();

  for (var j = 0; j < charts.length; j++) {
    sheet.removeChart(charts[j]); // this will remove the charts
  }
}
