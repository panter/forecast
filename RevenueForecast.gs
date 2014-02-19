// Configuration:
// enter a valid authentication token here, user must have controllr api access.
// enter the root url to your controllr server
var AUTH_TOKEN = ""
var API_HOST_URL = "http://controllr-staging.panter.biz/"

// Configuration end. Don't change anything below here.

// url to load the revenue forecast data
var finalResourceURL = API_HOST_URL + "api/revenue_forecast.json?user_token=" + AUTH_TOKEN;

// map column order to var names
var projectNameColumn = 1;
var startDateColumn = 3;
var endDateColumn = 4;
var revenueColumn = 5;
var probabilityColumn = 6;
var stateColumn = 7;
var isActiveColumn = 8;
var firstMonthColumn = 9; // first column with a month

// column mapping for A1 notation
var columnMapping = "0ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");

// input validation with dropdowns for "state" and "active" columns
var yesNoValidation = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build();
var stateValidation = SpreadsheetApp.newDataValidation().requireValueInList(['lead', 'offered', 'won', 'running', 'closing', 'permanent'], true).build();

function loadCurrentForecast() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var data = loadForecastJSON(finalResourceURL);
  var rows = data.rows;

  renderProjectRows(rows, data.base_url);
  formatHeaderRow();
  renderTotalRow(rows.length);
  applyCellValidations(rows.length-1);
};

function renderProjectRows(projects, baseURL){
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();

  if(!projects){
    throw "Es sind keine Projektdaten vorhanden"
  }

  // the header row has the definite count of columns
  var lastMonthColumn = projects[0].length;

  projects.forEach(function(project, idx){
    // header row case
    if(idx == 0) {
      sheet.appendRow(project);
      return;
    }

    // handle project rows
    var projectName = project[0];
    var leader = project[1];
    var startDate = project[2];
    var endDate = project[3];
    var revenue = project[4];
    var probability = project[5];
    var state = project[6];
    var isActive = project[7];
    // .. rest are monthly revenue values
    var accountingPath = project.pop();
    var editPath = project.pop();
    var path = project.pop();
    var row = idx + 1 // row starts at 1 + header row offset

    // cell references in A1 notation
    var totalRevenueCell = "$"+columnMapping[revenueColumn]+"$"+row;
    var probabilityCell = "$"+columnMapping[probabilityColumn]+"$"+row;
    var startDateCell = "$"+columnMapping[startDateColumn]+"$"+row;
    var endDateCell = "$"+columnMapping[endDateColumn]+"$"+row;
    var projectNameCell = "$"+columnMapping[projectNameColumn]+"$"+row;

    // round dates to first/last day of month so we get meaningful dateranges
    // to distribute the revenue accross
    var firstDayOfStartDate = "DATE(YEAR("+startDateCell+"); MONTH("+startDateCell+"); 1)";
    var lastDayOfEndDate = "DATE(YEAR("+endDateCell+"); MONTH("+endDateCell+"); 28)"; // 28 matches february and is always close enough

    // convert true/false to yes/no
    project[isActiveColumn-1] = project[isActiveColumn-1] ? 'Yes' : 'No';

    sheet.appendRow(project)

    // apply formula to monthly revenue cells with a value
    for(var column = firstMonthColumn; column <= lastMonthColumn; column++){
      var cellHasRevenue = project[column-1] != 0;

      // format project revenue per month
      if(cellHasRevenue){
        // the monthly revenue cell
        var cell = sheet.getRange(row, column);
        var formula = "=ROUND("+totalRevenueCell+"/(ROUND(("+lastDayOfEndDate+"-"+firstDayOfStartDate+")/30)))*"+probabilityCell;
        cell.setFormula(formula);
      }
    }

    // format project name
    var projectNameFormula = '=hyperlink("'+baseURL+path+'"; "'+projectName+'")';
    var projectNameRange = sheet.getRange(projectNameCell)
    projectNameRange.setFormula(projectNameFormula);

    // format start date
    var startDateFormula = '=hyperlink("'+baseURL+editPath+'"; "'+startDate+'")';
    var startDateRange = sheet.getRange(startDateCell)
    startDateRange.setFormula(startDateFormula);

    // format end date
    var endDateFormula = '=hyperlink("'+baseURL+editPath+'"; "'+endDate+'")';
    var endDateRange = sheet.getRange(endDateCell)
    endDateRange.setFormula(endDateFormula);

    // format revenue
    var revenueFormula = '=hyperlink("'+baseURL+accountingPath+'"; "'+revenue+'")';
    var revenueRange = sheet.getRange(totalRevenueCell)
    var revenue = revenueRange.getValue();
    revenueRange.setFormula(revenueFormula);
  })
}

function renderTotalRow(numberOfProjects){
  var sheet = SpreadsheetApp.getActiveSheet();

  var rowOffset = 2; // empty rows between projects and total
  var rowNumber = numberOfProjects + 1 + rowOffset;

  var columnOffset = firstMonthColumn; // first column with a month
  var lastColumn = sheet.getLastColumn();

  for(var column = columnOffset; column <= lastColumn; column++){
    var cell = sheet.getRange(rowNumber, column);
    // 9 = SUM function, use SUBTOTAL so filtered rows are correctly summed for the total
    cell.setFormulaR1C1("=SUBTOTAL(9;R2C"+column+":R"+(rowNumber-rowOffset)+"C"+column+")");
    cell.setFontWeight("bold");
  }
}

function applyCellValidations(numProjects){
  var sheet = SpreadsheetApp.getActiveSheet();

  var stateRange = sheet.getRange(2, stateColumn, numProjects);
  stateRange.setDataValidation(stateValidation);

  var isActiveRange = sheet.getRange(2, isActiveColumn, numProjects);
  isActiveRange.setDataValidation(yesNoValidation);
}

function formatHeaderRow(){
  var sheet = SpreadsheetApp.getActive();
  var range = sheet.getRange("A1:Z1");
  range.setFontWeight("bold");
}

function loadForecastJSON(url) {
  var response = UrlFetchApp.fetch(url, { method : "GET" });
  var data = JSON.parse(response.getContentText());
  return data;
}

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Aktuellen Revenue Forecast laden",
    functionName : "loadCurrentForecast"
  }];
  sheet.addMenu("Controllr", entries);
};
