/*
 CONSTANTS START
*/
// URL for AWS Lambda clustering service
const CLUSTERING_SERVICE_URL =
  "https://6t2mrznt84.execute-api.us-east-2.amazonaws.com/default/clusterAddresses";

// Names of tabs we care about.
const SHEET = {
  intakeForm: "Intake Form",
  deliveries: "Deliveries",
  closedCompleted: "Closed/Completed",
  geocoding: "Geocoding",
  multiSelects: "MultiSelects",
};

// Names of columns we care about.
const DELIVERY_COLUMNS = {
  address: "Address",
  cluster: "Cluster",
  dateCompleted: "Date Delivered",
  date: "Date",
  status: "Status",
  time: "Time",
  uid: "UID",
  urgent: "Needs Urgent?",
};

const GEOCODING_COLUMNS = {
  address: "addresskey",
  latitude: "latitude",
  longitude: "longitude",
};

// Names of statuses we care about.
const STATUS = {
  closed: "Closed",
  delivered: "Delivered",
  pendingDelivery: "Pending - Delivery",
};
/*
 CONSTANTS END
*/

/*
 UTILITIES START
*/
function getHeaderRow(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getColumnIdx(columnName, sheetName = null) {
  // Returns 0-indexed column index
  if (sheetName === null) {
    sheetName = SHEET.deliveries;
  }
  const idx = getHeaderRow(sheetName).indexOf(columnName);
  // If no UID column, assume it is the first column
  if (idx == -1 && columnName == DELIVERY_COLUMNS.uid) {
    return 0;
  }
  return idx;
}

function colIdxToLetter(value) {
  // 0-indexed col to letter for column value up to 701
  if (value < 26) {
    return (value + 10).toString(36).toUpperCase();
  }
  var letters = "";
  letters += (Math.floor(value / 26) + 9).toString(36);
  letters += ((value % 26) + 10).toString(36);
  return letters.toUpperCase();
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
/*
 UTILITIES END
*/

/*
  MULTI-SELECT START
*/
class MultiSelectSheet {
  constructor(mappingSheetName, delimiter = ", ") {
    this.delimiter = delimiter;
    var mappingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      mappingSheetName
    );
    this._multi_select_map = {};
    for (let colIdx = 0; colIdx < mappingSheet.getLastColumn(); colIdx++) {
      // colIdx+1 is to convert 0-index to 1-index
      const colVals = mappingSheet
        .getRange(1, colIdx + 1, mappingSheet.getLastRow())
        .getValues()
        .flat();
      const header = colVals[0];
      const splitHeader = header.split(".");
      const sheetName = splitHeader[0];
      const colName = splitHeader[1];
      if (!colName || !sheetName) {
        continue;
      }
      const sheetColIdx = getColumnIdx(colName, sheetName);
      if (sheetColIdx === -1) {
        const msg = `Could not find column '${colName}' on sheet '${sheetName}'`;
        console.error(msg);
        Logger.log(msg);
        continue;
      }
      if (this._multi_select_map[sheetName] == undefined) {
        this._multi_select_map[sheetName] = { by_name: {}, by_idx: {} };
      }

      const colOptions = colVals
        .slice(1)
        .map((val) => val.trim().toLowerCase())
        .filter((val) => val.length > 0);
      var data = {
        options: colOptions,
        colName: colName,
        colIdx: sheetColIdx,
        colLetter: colIdxToLetter(sheetColIdx),
      };
      this._multi_select_map[sheetName]["by_name"][colName] = data;
      this._multi_select_map[sheetName]["by_idx"][sheetColIdx] = data;
    }
    Logger.log(
      `MultiSelect Validation config ${JSON.stringify(this._multi_select_map)}`
    );
  }

  getColOptionsByName(sheetName, colName) {
    // return null if this sheetName/colName combo is not found
    if (sheetName in this._multi_select_map) {
      if (colName in this._multi_select_map[sheetName]["by_name"]) {
        return this._multi_select_map[sheetName]["by_name"][colName]["options"];
      }
    }
    return null;
  }

  getColOptionsByIdx(sheetName, colIdx) {
    // return null if this sheetName/colIdx combo is not found
    const colIdxStr = String(colIdx);
    if (sheetName in this._multi_select_map) {
      if (colIdxStr in this._multi_select_map[sheetName]["by_idx"]) {
        return this._multi_select_map[sheetName]["by_idx"][colIdxStr][
          "options"
        ];
      }
    }
    return null;
  }

  isColumnTracked(sheetName, colName = null, colIdx = null) {
    if (colName !== null) {
      return Boolean(this.getColOptionsByName(sheetName, colName));
    }
    if (colIdx !== null) {
      return Boolean(this.getColOptionsByIdx(sheetName, colIdx));
    }
    return false;
  }

  isSheetTracked(sheetName) {
    return sheetName in this._multi_select_map;
  }

  setAllDataValidations() {
    for (let sheetName in this._multi_select_map) {
      let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        sheetName
      );
      Logger.log(`Processing data validations for ${sheetName}`);
      for (let colIdxStr in this._multi_select_map[sheetName]["by_idx"]) {
        const colLetter = this._multi_select_map[sheetName]["by_idx"][
          colIdxStr
        ]["colLetter"];
        const a1Range = `${colLetter}:${colLetter}`;
        let colRange = sheet.getRange(a1Range);
        const allowedVals = this.getColOptionsByIdx(sheetName, colIdxStr);
        Logger.log(
          `Adding data validation to column ${colLetter}. Options ${allowedVals}`
        );
        let rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(allowedVals)
          .build();
        colRange.setDataValidation(rule);
      }
    }
    // clear validations on the header row
    for (let sheetName in this._multi_select_map) {
      let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        sheetName
      );
      var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      headerRange.clear({ validationsOnly: true });
    }
  }
}
/*
  MULTI-SELECT END
*/

/*
 onOpen TRIGGER START
*/
const multiSelectValidator = new MultiSelectSheet(SHEET.multiSelects);

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  //  addColumns([DELIVERY_COLUMNS.latitude, DELIVERY_COLUMNS.longitude, DELIVERY_COLUMNS.cluster], true);
  ui.alert(
    'READ THIS MESSAGE BEFORE EDITING! \n There is a new AUTOMATED data system that sorts data into the DELIVERY tab when their STATUS is changed. \n NEEDS URGENT is no longer a status, it is a check box you will find in the row. \n Please fill in ALL other information for an order before changing the status column. \n If you change status to "pending delivery", that row will automatically be copied into those tabs. \n However, any changes made to the row in the new intake form after that WILL NOT be reflected in the pickup or delivery tabs. \n That is why its important to change the status after filling in ALL other information. Otherwise, please go and update information in the other tabs as well. Thank you!'
  );
  ui.createMenu("Automation")
    .addItem("Post-delivery automation", "startUpMessage")
    .addToUi();
  ui.createMenu("Delivery Clustering")
    .addItem("Sort Delivery Rows by Priority", "prioritizeRows")
    .addItem("Geocode Delivery Addresses", "geocode")
    .addItem("Cluster Delivery Rows", "clusteringPrompt")
    .addToUi();
  multiSelectValidator.setAllDataValidations();
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
/*
 onOpen TRIGGER END
*/

/*
 onEdit TRIGGER START
*/
function onEdit(event) {
  statusChangeAutomation(event);
  // multiSelectValidation(event);
}

function statusChangeAutomation(event) {
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
  if (cellC != getColIndex(DELIVERY_COLUMNS.status)) {
    return;
  }

  if (sheet.getName() == SHEET.intakeForm) {
    //if new intake status pending delivery or needs urgent will be moved to DD sheet
    if (cellValue == STATUS.pendingDelivery) {
      SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(SHEET.deliveries)
        .appendRow(
          sheet.getRange(cellR, 1, 1, sheet.getLastColumn()).getValues()[0]
        );
      cell.setValue("Scheduled Delivery");
      return;
    }

    // if delivery intake status is closed/delivered moves to closed/completed and hides from new intake/ had to put 30 as the max columns because getLastColumn function for 'range' was not pulling all columns.
    if (cellValue == STATUS.closed) {
      SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(SHEET.intakeForm)
        .getRange(cellR, getColIndex(DELIVERY_COLUMNS.dateCompleted))
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

function multiSelectValidation(e) {
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) {
    // only care about single-cell edits
    return;
  }
  const sheetName = e.range.getSheet().getName();
  if (!multiSelectValidator.isSheetTracked(sheetName)) {
    // only care about edits on particular sheets
    return;
  }
  if (e.range.getRow() === 1) {
    // don't care about header row
    return;
  }
  // make it 0-indexed
  const colIdx = e.range.getColumn() - 1;
  const colOptions = multiSelectValidator.getColOptionsByIdx(sheetName, colIdx);
  if (!colOptions) {
    return;
  }

  const cell = SpreadsheetApp.getActiveSpreadsheet().getCurrentCell();
  if (!e.value) {
    cell.setValue("");
  } else if (!e.oldValue) {
    cell.setValue(value);
  } else {
    cell.setValue(`${e.oldValue}${multiSelectValidator.delimiter}${e.value}`);
  }
}
/*
 onEdit TRIGGER END
*/

/*
 CLUSTERING FUNCTIONS START
*/
function prioritizeRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.deliveries
  );
  let range = sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  );
  range.sort([
    { column: getColumnIdx(DELIVERY_COLUMNS.urgent) + 1, ascending: false },
    { column: getColumnIdx(DELIVERY_COLUMNS.date) + 1, ascending: true },
    { column: getColumnIdx(DELIVERY_COLUMNS.time) + 1, ascending: true },
  ]);
}

function getGeocodedAddrs() {
  var geocodingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.geocoding
  );
  const geocodeAddrIdx = getColumnIdx(
    GEOCODING_COLUMNS.address,
    SHEET.geocoding
  );
  const geocodeLatIdx = getColumnIdx(
    GEOCODING_COLUMNS.latitude,
    SHEET.geocoding
  );
  const geocodeLonIdx = getColumnIdx(
    GEOCODING_COLUMNS.longitude,
    SHEET.geocoding
  );
  const geocodeCells = geocodingSheet.getRange(
    2,
    1,
    geocodingSheet.getLastRow(),
    geocodingSheet.getLastColumn()
  );
  const geocodedNested = geocodeCells.getValues();
  var geocodedAddrsMap = {};
  for (var r of geocodedNested) {
    const addr = r[geocodeAddrIdx];
    const lat = r[geocodeLatIdx];
    const lon = r[geocodeLonIdx];
    if (addr && lat && lon) {
      geocodedAddrsMap[addr] = [lat, lon];
    }
  }
  return geocodedAddrsMap;
}

function addGeocodedAddr(addr, lat, lon) {
  var geocodingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.geocoding
  );
  const geocodeAddrIdx = getColumnIdx(
    GEOCODING_COLUMNS.address,
    SHEET.geocoding
  );
  const geocodeLatIdx = getColumnIdx(
    GEOCODING_COLUMNS.latitude,
    SHEET.geocoding
  );
  const geocodeLonIdx = getColumnIdx(
    GEOCODING_COLUMNS.longitude,
    SHEET.geocoding
  );

  let range = geocodingSheet.getRange(
    geocodingSheet.getLastRow() + 1,
    geocodeAddrIdx + 1
  );
  range.setValue(addr);
  range = geocodingSheet.getRange(
    geocodingSheet.getLastRow(),
    geocodeLatIdx + 1
  );
  range.setValue(lat);
  range = geocodingSheet.getRange(
    geocodingSheet.getLastRow(),
    geocodeLonIdx + 1
  );
  range.setValue(lon);
}

function geocode() {
  var deliveriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.deliveries
  );
  const deliveryAddrIdx = getColumnIdx(
    DELIVERY_COLUMNS.address,
    SHEET.deliveries
  );
  let geocodedAddrsMap = getGeocodedAddrs();
  let geocodedAddrs = Object.keys(geocodedAddrsMap);
  var geocoder = Maps.newGeocoder().setBounds(
    38.81604,
    -77.14538,
    39.00865,
    -76.90918
  );
  for (let rowIdx = 2; rowIdx <= deliveriesSheet.getLastRow(); rowIdx++) {
    const addr = deliveriesSheet
      .getRange(rowIdx, deliveryAddrIdx + 1)
      .getValue();
    if (!addr) {
      continue;
    }
    if (geocodedAddrs.includes(addr)) {
      continue;
    }
    var resp = geocoder.geocode(addr);
    if (resp.status !== "OK") {
      Logger.log(
        addr,
        " failed to geocode. Error type: '",
        resp.status_code,
        "'. Error message: '",
        resp.error_message,
        "'"
      );
      continue;
    }
    var result = resp.results[0];
    geocodedAddrs.push(addr);
    const lat = result.geometry.location.lat;
    const lon = result.geometry.location.lng;
    geocodedAddrsMap[addr] = [lat, lon];
    Logger.log(addr, " --> ", lat, ", ", lon);
    addGeocodedAddr(addr, lat, lon);
  }
}

function clusterAddresses(numberOfRows = 45) {
  var deliveriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.deliveries
  );
  const rows = deliveriesSheet
    .getRange(2, 1, numberOfRows, deliveriesSheet.getLastColumn())
    .getValues();
  let geocodedAddrsMap = getGeocodedAddrs();
  const payload = prepareClusteringPayload(rows, geocodedAddrsMap);
  Logger.log("Clustering Payload: ", payload);
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  };
  let response = UrlFetchApp.fetch(CLUSTERING_SERVICE_URL, options);
  var responseData = JSON.parse(response.getContentText());
  updateRowsWithClusters(responseData);
}

function clusteringPrompt() {
  var ui = SpreadsheetApp.getUi();
  let numRows = NaN;

  while (isNaN(numRows)) {
    var result = ui.prompt(
      "How many rows do you want to cluster?",
      ui.ButtonSet.OK_CANCEL
    );

    var button = result.getSelectedButton();
    var text = result.getResponseText();
    if (button == ui.Button.OK) {
      numRows = parseInt(text);
      if (!isNaN(numRows) && numRows > 0) {
        clusterAddresses(numRows);
        return;
      }
    }
    ui.alert("Please enter a positive, whole number.");
  }
}

function prepareClusteringPayload(rows, geocodingMap) {
  const uidIdx = getColumnIdx(DELIVERY_COLUMNS.uid, SHEET.deliveries);
  const addrIdx = getColumnIdx(DELIVERY_COLUMNS.address, SHEET.deliveries);
  const urgentIdx = getColumnIdx(DELIVERY_COLUMNS.urgent, SHEET.deliveries);
  let points = [];
  for (var r of rows) {
    let uid = r[uidIdx];
    let addr = r[addrIdx];
    let isUrgent = false;
    if (r[urgentIdx] && r[urgentIdx].toLowerCase().trim() == "yes") {
      isUrgent = true;
    }
    if (uid && addr && addr in geocodingMap) {
      points.push({
        _id: uid,
        extra_data: {
          is_urgent: isUrgent,
        },
        coords: geocodingMap[addr],
      });
    }
  }
  return {
    points: points,
  };
}

function updateRowsWithClusters(clusterData) {
  var deliveriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.deliveries
  );
  const uidIdx = getColumnIdx(DELIVERY_COLUMNS.uid, SHEET.deliveries);
  const clusterIdx = getColumnIdx(DELIVERY_COLUMNS.cluster, SHEET.deliveries);
  var range = deliveriesSheet.getRange(
    2,
    1,
    deliveriesSheet.getLastRow(),
    deliveriesSheet.getLastColumn()
  );
  var sheetValues = range.getValues();
  for (var rowWithCluster of clusterData) {
    if (!rowWithCluster["_id"] || !rowWithCluster["cluster"]) {
      continue;
    }
    let uid = rowWithCluster["_id"];
    let cluster = rowWithCluster["cluster"];
    for (var sheetRow of sheetValues) {
      if (sheetRow[uidIdx] != uid) {
        continue;
      }
      sheetRow[clusterIdx] = cluster;
      break;
    }
  }
  range.setValues(sheetValues);
}
/*
 CLUSTERING FUNCTIONS END
*/

/*
  UNUSED AUTOMATION START
*/
//if DD sheet status is set to delivered, changes status to delivered in New Intake and thus triggers moving it to the closed/completed tab
//if(sheet.getName() == SHEET.deliveries){
//  //if(cellC == getColIndex(DELIVERY_COLUMNS.status)){
//    if(cellValue == STATUS.delivered){
//     for(y=1; y < SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getLastRow(); y++){
//       if(sheet.getRange(cellR, 1).getValue() == SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, 1).getValue()){
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, getColIndex(DELIVERY_COLUMNS.status)).setValue('Delivered');
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, getColIndex(DELIVERY_COLUMNS.dateCompleted)).setValue(new Date());
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
    var d = 1;
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
    var n = 1;
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
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, getColIndex(DELIVERY_COLUMNS.status)).setValue('Delivered');
//        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.intakeForm).getRange(y, getColIndex(DELIVERY_COLUMNS.dateCompleted)).setValue(new Date());
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
/*
 UNUSED AUTOMATION END
*/
