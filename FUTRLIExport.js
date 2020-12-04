function mergeObjects(destination, source) {
  for (var property in source) {
    if (!source.hasOwnProperty(property)) {
      continue;
    }

    destination[property] = source[property];
  }
  return destination;
}

function addDays(date, days) {
  newDate = new Date();
  newDate.setTime(date.getTime() + days * 24 * 60 * 60 * 1000);
  return newDate;
}

function addHours(date, hours) {
  newDate = new Date();
  newDate.setTime(date.getTime() + hours * 60 * 60 * 1000);
  return newDate;
}

function appendRow(row) {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(row);
}

function createNewSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet(name);
}

function transpose(a) {
  return Object.keys(a[0]).map(function (c) {
    return a.map(function (r) {
      return r[c];
    });
  });
}

function getRows() {
  return SpreadsheetApp.getActiveSheet().getDataRange().getValues();
}

function getColumns() {
  var rows = getRows();
  return transpose(rows);
}

function getColumn(idx) {
  var columns = getColumns();
  return columns[idx];
}

function getRow(idx) {
  var rows = getRows();
  return rows[idx];
}

function find(arr, elemToFind) {
  for (var i = 0; i < arr.length; i++) {
    if (arr[i] === elemToFind) {
      return i;
    }
  }
  return -1;
}

function findRow(columnIdx, value) {
  var col = getColumn(columnIdx);
  var rowIdx = find(col, value);
  return rowIdx;
}

function findCol(rowIdx, value) {
  var row = getRow(rowIdx);
  var columnIdx = find(row, value);
  return columnIdx;
}

function cellValue(cell) {
  return SpreadsheetApp.getActiveSheet().getRange(cell).getValue();
}

function numToLetter(num) {
  return "ABCDEFGHIKJLMNOPQRSTUVWXQZ"[num];
}

function letterToNum(letter) {
  var letters = "ABCDEFGHIKJLMNOPQRSTUVWXQZ".split("");
  var letterIdx = find(letters, letter);
  return letterIdx;
}

function nextLetter(letter) {
  letter = letterToNum(letter);
  letter += 1;
  return numToLetter(letter);
}

function prevLetter(letter) {
  letter = letterToNum(letter);
  letter -= 1;
  return numToLetter(letter);
}

function getAccountingHeaders() {
  return [
    "Account Category",
    "Account",
    "Forecast Name",
    "Credit Terms",
    "Sales tax %",
    "20-Jan",
    "20-Feb",
    "20-Mar",
    "20-Apr",
    "20-May",
    "20-Jun",
    "20-Jul",
    "20-Aug",
    "20-Sep",
    "20-Oct",
    "20-Nov",
    "20-Dec",
  ];
}

function nextEmptyRow(col, row) {
  var cell = col + row;
  var value = cellValue(cell);
  while (value) {
    row += 1;
    cell = "A" + row;
    value = cellValue(cell);
  }
  return row;
}

function nextEmptyCol(col, row) {
  var cell = col + row;
  var value = cellValue(cell);
  while (value) {
    Logger.log(value);
    col = nextLetter(col);
    cell = col + row;
    value = cellValue(cell);
  }
  return col;
}

function deleteRow(grid, idx, elem) {
  return grid.filter(function (row) {
    return row[idx] !== elem;
  });
}

function removeEventNames(eventsGrid) {
  // Take row where row[0] !=== "Event Name"
  return eventsGrid.filter(function (row) {
    return row[0] !== "Event Name";
  });
}

function getEventsGrid() {
  var startRow = findRow(0, "Website") + 1;
  var endRow = nextEmptyRow("A", startRow) - 1;
  var columns = getColumns();
  var numOfCols = columns.length;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];
  var grid = sheet.getRange(startRow, 1, endRow, numOfCols).getValues();
  return grid;
}

function getAccountingGrid() {
  var startRow = findRow(0, "Accounting Export") + 1;
  var endRow = nextEmptyRow("A", startRow) - 1;
  var columns = getColumns();
  var numOfCols = columns.length;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];
  var grid = sheet.getRange(startRow, 1, endRow, numOfCols).getValues();
  return grid;
}

function getEventNames(eventsGrid) {
  // Take row where row[0] === Event Name
  var eventNames = eventsGrid.filter(function (row) {
    return row[0] === "Event Name";
  });
  // eventNames is [[ ... ]] because of .filter
  eventNames = eventNames[0];
  // Remove the header: "Event Name"
  return eventNames.slice(1);
}

function getAccountingLabels(accountingGrid) {
  var labels = [];
  accountingGrid.forEach(function (row) {
    // first elem of row is label
    // ex. ["Hotels Revene", 100, 300, 200]
    labels.push(row[0]);
  });
  // Remove header -> "Accounting Export"
  return labels.slice(1);
}

function createEvent(eventsGrid, accountingGrid, idx) {
  hash = {};
  eventsGrid.forEach(function (row) {
    hash[row[0]] = row[idx + 1];
  });
  accountingGrid.forEach(function (row) {
    hash[row[0]] = row[idx + 1];
  });
  return hash;
}

function createEvents(eventsGrid, accountingGrid, eventNames) {
  events = {};
  eventNames.forEach(function (eventName, idx) {
    var event = createEvent(eventsGrid, accountingGrid, idx);
    events[eventName] = event;
    return;
  });
  return events;
}

var START_DATE = {
  DEFAULT: [{ percent: 100, endDate: false, days: 0 }],
};

var END_DATE = {
  DEFAULT: [{ percent: 100, endDate: true, days: 0 }],
};

var STAFF_EXPENSE = {
  DEFAULT: [
    { percent: 50, endDate: false, days: -30 },
    { percent: 50, endDate: true, days: 0 },
  ],
};

var HOTELS_REVENUE = {
  DEFAULT: [{ percent: 100, endDate: true, days: 30 }],
};

var SHIPPING_EXPENSE = {
  DEFAULT: [
    { percent: 50, endDate: false, days: -2 },
    { percent: 50, endDate: true, days: 2 },
  ],
};

var TEAM_REGISTRATION_REVENUE = {
  "Region Finals": [
    { percent: 40, days: "JANURARY" },
    { percent: 30, days: "FEBURARY" },
  ],
  "Prep Hoops": [
    { percent: 20, days: -40 },
    { percent: 60, days: -15 },
    { percent: 20, days: -10 },
  ],
  PGH: [
    { percent: 20, days: -40 },
    { percent: 60, days: -15 },
    { percent: 20, days: -10 },
  ],
  "Prep Girls Hoops": [
    { percent: 20, days: -40 },
    { percent: 60, days: -15 },
    { percent: 20, days: -8 },
  ],
  "Prep Dig": [{ percent: 25, days: -31 }],
};

var ACCOUNTING_DATES = {
  "Team Registration Revenue": TEAM_REGISTRATION_REVENUE,
  "Hotels Revenue": HOTELS_REVENUE,
  "Staff Expense": STAFF_EXPENSE,
  "Shipping Expense": SHIPPING_EXPENSE,
  "Scheduling Expense": START_DATE,
  "Officials Expense": START_DATE,
  "Printing Expense": START_DATE,
  "Gate Revenue": END_DATE,
  "Concessions Revenue": END_DATE,
  "Workers Expense": END_DATE,
  "Athletic Trainers Expense": END_DATE,
  "Meals Expense": END_DATE,
  "Facility Expense": END_DATE,
  "Media Expense": END_DATE,
  "Team Accommodations Expense": END_DATE,
  "Apparel Expense": END_DATE,
  "Championship Balls Expense": END_DATE,
};

function getLabelsRow(eventName, accountingLabel) {
  var account = eventName + ": " + accountingLabel;
  var labels = ["", account, "", "", ""];
  return labels;
}

function getFinances(events, eventName, accountingLabel) {
  if (!events[eventName]) {
    return;
  }

  var event = events[eventName];

  if (!ACCOUNTING_DATES[accountingLabel]) {
    return;
  }

  var accountingDates = ACCOUNTING_DATES[accountingLabel];

  var dates = ["", "", "", "", "", "", "", "", "", "", "", ""];

  var startDate = event["Event Start Date"];
  var endDate = event["Event End Date"];

  var totalDollars = event[accountingLabel];

  var isRegionFinal = eventName.indexOf("Region Finals") !== -1;

  var edgeCase = isRegionFinal ? "Region Finals" : event["Website"];

  var dayChanges = accountingDates[edgeCase]
    ? accountingDates[edgeCase]
    : accountingDates["DEFAULT"];

  dayChanges.forEach(function (dayChange) {
    var date = dayChange["endDate"] ? endDate : startDate;
    var dollars = totalDollars * (dayChange["percent"] / 100);
    var days = dayChange["days"];

    // Add 1 because GMT is 1hr behind CST
    var newDate = addDays(date, days);

    // If expenses go into prev year, ignore them
    if (date.getMonth() === 0 && newDate.getMonth() === 11) {
      return;
    }

    // If expenses go into next year, ignore them
    if (date.getMonth() === 11 && newDate.getMonth() === 0) {
      return;
    }

    if (dayChange["days"] === "JANURARY") {
      newDate = new Date();
      newDate.setMonth(0);
    }

    if (dayChange["days"] === "FEBURARY") {
      newDate = new Date();
      newDate.setMonth(1);
    }

    var month = newDate.getMonth();
    dates[month] ? (dates[month] += dollars) : (dates[month] = dollars);
  });

  return dates;
}

function setAccountingMonths(events, eventNames, accountingLabels) {
  var row = 2;

  eventNames.forEach(function (eventName) {
    accountingLabels.forEach(function (accountingLabel) {
      Logger.log(accountingLabel);

      var account = eventName + ": " + accountingLabel;

      var website = events[eventName]["Website"];

      if (website === "Prep Dig") {
        account = "Prep Dig " + account;
      }

      if (website === "PGH") {
        account = "PGH " + account;
      }

      var finances = getFinances(events, eventName, accountingLabel);
      var row = ["", account, "", "", ""].concat(finances);

      if (!finances) {
        return;
      }

      appendRow(row);
    });
  });
}

function getSheetName() {
  var today = new Date();
  var day = today.getDate();
  var month = today.getMonth() + 1;
  return month + "-" + day + " Forcast";
}

function lettersToNumber(letters) {
  for (var p = 0, n = 0; p < letters.length; p++) {
    n = letters[p].charCodeAt() - 64 + n * 26;
  }
  return n;
}

function test() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
}

function FUTRLIExport() {
  var eventsGrid = getEventsGrid();
  var eventsGridNoName = deleteRow(eventsGrid, 0, "Event Name");
  var eventNames = getEventNames(eventsGrid);
  var accountingGrid = getAccountingGrid();
  var accountingGridNoHeader = accountingGrid.slice(1);
  var accountingLabels = getAccountingLabels(accountingGrid);
  var accountingHeaders = getAccountingHeaders();

  var events = createEvents(
    eventsGridNoName,
    accountingGridNoHeader,
    eventNames
  );

  var sheetName = getSheetName();

  createNewSheet(sheetName);
  appendRow(accountingHeaders);

  setAccountingMonths(events, eventNames, accountingLabels);
}
