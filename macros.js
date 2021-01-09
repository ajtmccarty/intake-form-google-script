function onOpen() {
  var ui = SpreadsheetApp.getUi();
  //  addColumns([COLUMN.latitude, COLUMN.longitude, COLUMN.cluster], true);
  ui.alert(
    'READ THIS MESSAGE BEFORE EDITING! \n There is a new AUTOMATED data system that sorts data into the DELIVERY tab when their STATUS is changed. \n NEEDS URGENT is no longer a status, it is a check box you will find in the row. \n Please fill in ALL other information for an order before changing the status column. \n If you change status to "pending delivery", that row will automatically be copied into those tabs. \n However, any changes made to the row in the new intake form after that WILL NOT be reflected in the pickup or delivery tabs. \n That is why its important to change the status after filling in ALL other information. Otherwise, please go and update information in the other tabs as well. Thank you!'
  );
  ui.createMenu("Automation")
    .addItem("Post-delivery automation", "startUpMessage")
    .addToUi();
  //  ui.createMenu("Clustering")
  //      .addItem("Sort Rows by Priority", "prioritizeRows")
  //      .addToUi();
}

function startUpMessage() {
  let ui = SpreadsheetApp.getUi();
  let buttonPressed = ui.alert(
    "Please update status for all deliveries in DD sheet prior. Begin DD automation?",
    ui.ButtonSet.YES_NO
  );
  if (buttonPressed == ui.Button.YES) {
    ddAutomation();
  }
}

// Names of tabs we care about.
const SHEET = {
  intakeForm: "Intake Form",
  deliveries: "Deliveries",
  closedCompleted: "Closed/Completed",
};

// Names of columns we care about.
const COLUMN = {
  address: "Address",
  cluster: "Cluster",
  dateCompleted: "Date Delivered",
  date: "Date",
  latitude: "Lat",
  longitude: "Lon",
  status: "Status",
  time: "Time",
  uid: "UID",
  urgent: "Needs Urgent?",
};

// Names of statuses we care about.
const STATUS = {
  closed: "Closed",
  delivered: "Delivered",
  pendingDelivery: "Pending - Delivery",
};

/**
 * Handle edits to sheets.
 */
function onEdit(event) {
  const sheet = event.source.getActiveSheet();
  const cell = sheet.getActiveCell();
  const cellR = cell.getRow();
  const cellC = cell.getColumn();
  const cellValue = cell.getValue();
  const active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get all the column names.
  const columnNames = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()
    .shift();
  // Define a function that returns the 1-based index for the given name
  const getColIndex = (name) => {
    const index = columnNames.indexOf(name);
    if (index >= 0) {
      return index + 1;
    }
  };

  // Abort if this is a header edit or a change to multiple cells.
  if (cellR == 1) {
    return;
  }

  //Abort if this is not the status column.
  if (cellC != getColIndex(COLUMN.status)) {
    return;
  }

  if (sheet.getName() == SHEET.intakeForm) {
    //if new intake status pending delivery or needs urgent will be moved to DD sheet
    if (cellValue == STATUS.pendingDelivery) {
      cell.setValue("Scheduled Delivery");
      SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(SHEET.deliveries)
        .appendRow(
          sheet.getRange(cellR, 1, 1, sheet.getLastColumn()).getValues()[0]
        );
      return;
    }

    // if delivery intake status is closed/delivered moves to closed/completed and hides from new intake/ had to put 30 as the max columns because getLastColumn function for 'range' was not pulling all columns.
    if (cellValue == STATUS.closed) {
      SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(SHEET.intakeForm)
        .getRange(cellR, getColIndex(COLUMN.dateCompleted))
        .setValue(new Date());
      SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(SHEET.closedCompleted)
        .appendRow(sheet.getRange(cellR, 1, 1, 50).getValues()[0]);
      SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(SHEET.intakeForm)
        .deleteRows(cellR);
      return;
    }
  }
}

//if DD sheet status is set to delivered, changes status to delivered in New Intake and thus triggers moving it to the closed/completed tab
//if(sheet.getName() == SHEET.deliveries){
//  //if(cellC == getColIndex(COLUMN.status)){
//    if(cellValue == STATUS.delivered){
//     for(y=1; y < SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getLastRow(); y++){
//       if(sheet.getRange(cellR, 1).getValue() == SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, 1).getValue()){
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, getColIndex(COLUMN.status)).setValue('Delivered');
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, getColIndex(COLUMN.dateCompleted)).setValue(new Date());
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.closedCompleted).appendRow(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, 1, 1, 50).getValues()[0]);
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).deleteRow(y);
//        sheet.deleteRow(cellR);
//       }
//    }
//}
//}

//
//
function ddAutomation() {
  var ui = SpreadsheetApp.getUi();
  let now = new Date();
  let ddRows = [];
  let ddRows2 = [];
  //  let ddRows3 = [];
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //  const cell = sheet.getActiveCell();
  //  const cellR = cell.getRow();
  //  const cellC = cell.getColumn();
  //  const cellValue = cell.getValue();

  //if DD sheet status is set to delivered, changes status to delivered in New Intake, deletes from NI and D tab and copies to closed/completed
  for (
    d = 1;
    d < spreadsheet.getSheetByName(SHEET.deliveries).getLastRow();
    d++
  ) {
    if (spreadsheet.getName() == SHEET.deliveries) {
      if (spreadsheet.getRange(d, 3) == STATUS.delivered) {
        ddRows.push(
          spreadsheet.getSheetByName(SHEET.deliveries).getRange(d, 1).getValue()
        );
        ddRows2.push(d);
        // the [0] bit is changing a Matrix to a list because getRange().getValues returns a section of the sheet not a row
        spreadsheet
          .getSheetByName(SHEET.closedCompleted)
          .appendRow(
            spreadsheet
              .getSheetByName(SHEET.deliveries)
              .getRange(
                d,
                1,
                1,
                spreadsheet.getSheetByName(SHEET.deliveries).getLastColumn()
              )
              .getValues()[0]
          );
        spreadsheet.getSheetByName(SHEET.deliveries).deleteRow(d);
      }
    }

    ui.alert(`UIDs: ${ddRows} \n Row numbers: ${ddRows2}`);
  }

  for (
    n = 1;
    n < spreadsheet.getsheetbyname(SHEET.intakeForm).getLastRow();
    n++
  ) {
    for (let deliveredUID of ddRows) {
      if (
        spreadsheet
          .getSheetByName(SHEET.intakeForm)
          .getRange(n, 1)
          .getValue() == deliveredUID
      ) {
        spreadsheet.getSheetByName(SHEET.intakeForm).deleteRow(n);
      }
    }
  }

  // Delete rows from the bottom up so that you don't change row indices as you're iterating
  // Caution, reverse() changes the actual contents of rowNumsToDelete!
  //deletes the rows from the SSNI sheet that are closed/completed and moved to that tab
  for (let rowNum of rowNumsToDelete.reverse()) {
    ssNI.deleteRow(rowNum);
  }

  for (let deliveredRow of ddrows2.reverse()) {
    spreadsheet.getSheetByName(Deliveries).deleteRow(deliveredRow);
  }
}

//
//
//     for(y=1; y < SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getLastRow(); y++){
//       if(sheet.getRange(cellR, 1).getValue() == SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, 1).getValue()){
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, getColIndex(COLUMN.status)).setValue('Delivered');
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, getColIndex(COLUMN.dateCompleted)).setValue(new Date());
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.closedCompleted).appendRow(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, 1, 1, 50).getValues()[0]);
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).deleteRow(y);
//        sheet.deleteRow(cellR);
//       }
//    }
//}
//}
//
//
//  }
//
//  function processDelivered() {
//    let rowNumsToDelete = [];
//
//    //if a row is marked delivered or closed in NI sheet it will move that entire row to the closed/completed tab and delete those rows
//    for (j = 2; j < ssNI.getMaxRows(); j++){
//      if(ssNI.getRange(j, niStatCol).getValue() == 'Delivered' || ssNI.getRange(j, niStatCol).getValue() == 'Closed'){
//        // the [0] bit is changing a Matrix to a list because getRange().getValues returns a section of the sheet not a row
//        ssCC.appendRow(ssNI.getRange(j, 1, 1, ssNI.getLastColumn()).getValues()[0]);
//        rowNumsToDelete.push(j);
//      }
//    }
//
//    // Delete rows from the bottom up so that you don't change row indices as you're iterating
//    // Caution, reverse() changes the actual contents of rowNumsToDelete!
//    //deletes the rows from the SSNI sheet that are closed/completed and moved to that tab
//    for (let rowNum of rowNumsToDelete.reverse()){
//      ssNI.deleteRow(rowNum);
//    }
//  }

function getHeaderRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.deliveries
  );
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getColumnIdx(columnName) {
  return getHeaderRow().indexOf(columnName);
}

function addColumns(columnNames, hidden = false) {
  var headers = getHeaderRow();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.deliveries
  );
  var lastColIdx = sheet.getLastColumn();
  for (var i = 0; i < columnNames.length; i++) {
    const name = columnNames[i];
    if (!headers.includes(name)) {
      sheet.insertColumnAfter(lastColIdx + i + 1);
      var range = sheet.getRange(1, lastColIdx + i + 1);
      range.setValue(name);
      if (hidden) {
        sheet.hideColumns(lastColIdx + i + 1);
      }
    }
  }
}

function prioritizeRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.deliveries
  );
  let range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());
  let rows = range.getValues();
  setDateForRows(rows);
  let sortedRows = rows.sort(rowsCompareFunction);
  range.setValues(sortedRows);
}

function rowsCompareFunction(row1, row2) {
  let row1_is_urgent = isRowUrgent(row1);
  let row2_is_urgent = isRowUrgent(row2);
  if (row1_is_urgent && !row2_is_urgent) {
    return -1;
  }
  if (row2_is_urgent && !row1_is_urgent) {
    return 1;
  }
  const dateIdx = getColumnIdx(COLUMN.date);
  let row1_timestamp = row1[dateIdx];
  let row2_timestamp = row2[dateIdx];
  if (isNaN(row1_timestamp)) {
    return 1;
  }
  if (isNaN(row2_timestamp)) {
    return -1;
  }
  return row1_timestamp - row2_timestamp;
}

function isRowUrgent(row) {
  const urgentIdx = getColumnIdx(COLUMN.urgent);
  let val = row[urgentIdx];
  if (!val) {
    return false;
  }
  if (["yes", "true", "1"].includes(val.toLowerCase())) {
    return true;
  }
  return false;
}

function setDateForRows(rows) {
  const dateIdx = getColumnIdx(COLUMN.date);
  const timeIdx = getColumnIdx(COLUMN.time);
  for (r of rows) {
    var dateVal = r[dateIdx];
    var timeVal = r[timeIdx];
    if (!dateVal || !timeVal) {
      r[dateIdx] = Date.parse("12/31/2500");
      continue;
    }
    r[dateIdx].setHours(
      timeVal.getHours(),
      timeVal.getMinutes(),
      timeVal.getSeconds()
    );
  }
}
