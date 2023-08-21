// FY22 NOFO Fiscal Report Figures (Graphs)
// Graph Generation Sheet
targetSheet = "1YoalnLbU7oKYKfQu-mB4KM-xqVgOEhmafXcdzI4QwHw"
// S&T Metrics Sheet
stMetricsSheet = "17WNH_DGlP9pvsFj4v4EkQnrRv2ygJhR7kxUKliq0qYI"
// This document's ID
var docId = '1pLWFshH96w_NLUpRYqDXIeCra5w96dM0WoKK8fvCyMs';
// Change this for current fiscal year
var fiscalYear = 22

function onOpen() {
  var ui = DocumentApp.getUi();
  
  // Create a new menu in the Google Docs UI.
  ui.createMenu('Functions')
    .addItem('Copy Graphs From Spreadsheet', 'copyGraphsFromSpreadsheet')
    .addItem('Clear Images And Charts', 'clearImagesAndCharts')
    .addItem('Generate Table', 'copyToTable')
    .addItem('Remove Table', 'deleteTable')
    .addToUi();
}

function insertChartAtBookmark(targetSheetID, targetBookmarkId, chart_pos) {
  // Get the chart from Google Sheets
  var sheet = SpreadsheetApp.openById(targetSheetID);
  var charts = sheet.getSheets()[1].getCharts();

  if (charts.length > 0) {
    var chartBlob = charts[chart_pos].getBlob();

    // Insert the chart at the bookmark location
    var doc = DocumentApp.openById(docId);
    var bookmarks = doc.getBookmarks();
    var bookmarkPosition;

    for (var i = 0; i < bookmarks.length; i++) {
      if (bookmarks[i].getId() === targetBookmarkId) {
        bookmarkPosition = bookmarks[i].getPosition();
        break;
      }
    }

    if (bookmarkPosition) {
      var element = bookmarkPosition.getElement();
      var parent = element.getParent();
      var childIndex = parent.getChildIndex(element) + 1;

      // Create a new paragraph to insert the chart
      var newParagraph = parent.insertParagraph(childIndex, '');
      newParagraph.insertInlineImage(0, chartBlob);
    } else {
      Logger.log('No bookmark found with the specified ID.');
    }
  } else {
    Logger.log('No chart found in the specified Google Sheet.');
  }
}

function clearImagesAndCharts() {
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  var numChildren = body.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    var type = child.getType();

    if (type === DocumentApp.ElementType.INLINE_IMAGE || type === DocumentApp.ElementType.CHART) {
      child.removeFromParent();
    } else if (type === DocumentApp.ElementType.PARAGRAPH) {
      var numChildElements = child.asParagraph().getNumChildren();
      for (var j = 0; j < numChildElements; j++) {
        var childElement = child.asParagraph().getChild(j);
        if (childElement.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
          childElement.removeFromParent();
        }
      }
    } else if (type === DocumentApp.ElementType.TABLE) {
      var table = child.asTable();
      var numRows = table.getNumRows();
      var numCols = table.getRow(0).getNumCells();
      for (var row = 0; row < numRows; row++) {
        for (var col = 0; col < numCols; col++) {
          var cell = table.getCell(row, col);
          var numCellElements = cell.getNumChildren();
          for (var k = 0; k < numCellElements; k++) {
            var cellElement = cell.getChild(k);
            if (cellElement.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
              cellElement.removeFromParent();
            }
          }
        }
      }
    }
  }
}

function copyGraphsFromSpreadsheet() {
  // figure 1
  insertChartAtBookmark(targetSheet, "id.7a3x6x2kdol2", 0)
  // figure 2
  insertChartAtBookmark(targetSheet, "id.xe9lrfgmu2zy", 1)
  // figure 3
  insertChartAtBookmark(targetSheet, "id.aavkmxdty26n", 2)
  // figure 4
  insertChartAtBookmark(targetSheet, "id.hg1a6ep4c7po", 3)
  // figure 5
  insertChartAtBookmark(targetSheet, "id.k8bu1ljhdnk0", 4)
  // figure 6
  insertChartAtBookmark(targetSheet, "id.8vtwp89mtxsd", 5)
}

function listBookmarks() {
  var doc = DocumentApp.openById(docId);
  var bookmarks = doc.getBookmarks();

  bookmarks.forEach(function (bookmark) {
    var position = bookmark.getPosition();
    var location = position.getElement().asParagraph().getText() + ' (offset: ' + position.getOffset() + ')';
    Logger.log('Bookmark ID: %s, Location: %s', bookmark.getId(), location);
  });
}

function copyToTable() {
  var ssId = '17WNH_DGlP9pvsFj4v4EkQnrRv2ygJhR7kxUKliq0qYI'; // replace with your spreadsheet ID
  var spreadsheet = SpreadsheetApp.openById(ssId);
  var sheet = spreadsheet.getSheetByName('Grants');
  var doc = DocumentApp.openById(docId);
  var bookmarks = doc.getBookmarks();
  var body = doc.getBody();

  var targetBookmarkId = "id.ege8q1rd9wka"
  var bookmarks = doc.getBookmarks();
  var bookmarkPosition;

  for (var i = 0; i < bookmarks.length; i++) {
    if (bookmarks[i].getId() === targetBookmarkId) {
      bookmarkPosition = bookmarks[i].getPosition();
      break;
    }
  }

  var element = bookmarkPosition.getElement();
  var parent = element.getParent();
  var childIndex = parent.getChildIndex(element) + 1;

  // Create a table in the document
  var table = body.insertTable(childIndex);

  // Create a header row in the table
  var headerRow = table.appendTableRow();
  headerRow.appendTableCell("Project Title").setBackgroundColor('#00008B').setForegroundColor('#FFFFFF');;
  headerRow.appendTableCell("Entity, PI").setBackgroundColor('#00008B').setForegroundColor('#FFFFFF');;
  headerRow.appendTableCell("Theme").setBackgroundColor('#00008B').setForegroundColor('#FFFFFF');;
  headerRow.appendTableCell("Expected Close Out").setBackgroundColor('#00008B').setForegroundColor('#FFFFFF');;

  // Get all data from the sheet
  var data = sheet.getDataRange().getValues();

  // Assume the first row of the data contains the headers.
  var headers = data[0];
  var columnIndices = {};
  for (var j = 0; j < headers.length; j++) {
    columnIndices[headers[j]] = j;
  }

  // Loop through each row in the sheet
  for (var i = 0; i < data.length; i++) {
    // Now in the loop, you can use the column names:
    for (var i = 1; i < data.length; i++) { // start from 1 to skip the header row
      if (String(data[i][columnIndices['Funding Mech']]).includes('NOFO') && data[i][columnIndices['Project FY\n(project funded)']] === fiscalYear) {
        // Create a new row in the table
        var tableRow = table.appendTableRow();

        // Use column names instead of indices
        tableRow.appendTableCell(String(data[i][columnIndices['LT Assigned\nProject Name (from LT - do NOT type here)']])).setForegroundColor('#000000'); 
        tableRow.appendTableCell(String(data[i][columnIndices['PI Institution']]) + ", " + String(data[i][columnIndices['PI Last Name']]) + " " + String(data[i][columnIndices['PI First Name']])); 
        tableRow.appendTableCell(String(data[i][columnIndices['Mission Type']])); 
        var date = new Date(data[i][columnIndices['Grant\nEnd Date']]);
        var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/yyyy');

        tableRow.appendTableCell(formattedDate);
      }
    }
  }

  // Save and close the document
  doc.saveAndClose();
}

function deleteTable() {
  var targetBookmarkId = "id.ege8q1rd9wka";
  
  var doc = DocumentApp.openById(docId);
  var bookmarks = doc.getBookmarks();
  var bookmarkPosition;

  for (var i = 0; i < bookmarks.length; i++) {
    if (bookmarks[i].getId() === targetBookmarkId) {
      bookmarkPosition = bookmarks[i].getPosition();
      break;
    }
  }

  var element = bookmarkPosition.getElement();
  var parent = element.getParent();
  var childIndex = parent.getChildIndex(element) + 1;

  // Get the element at the child index
  var tableElement = parent.getChild(childIndex);

  // Check if the element is a table
  if (tableElement.getType() === DocumentApp.ElementType.TABLE) {
    // Remove the table
    parent.removeChild(tableElement);
  }

  // Save and close the document
  doc.saveAndClose();
}

