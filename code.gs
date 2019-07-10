

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SEMRush API')
  .addItem('Get Traffic', 'getTrafficFromSEMRush')
  .addItem('Update Traffic', 'updateTrafficFromSEMRush')
  .addItem('Show Chart', 'showChartDialog')
  .addToUi();
}

var REASONABLE_TIME_TO_WAIT = 60*1000;
var APIKEY = 'efbace1d49e76b459fcf5d02e27b92ea';
var errorMessage = 'Data not found';

function getTrafficFromSEMRush(){
  var startTime = new Date();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var domains = sheet.getDataRange().getValues().splice(1);
  var keywordArr = [];
  var dateArr = [];

  for(var i=0;i<domains.length;i++){
    var domain = domains[i][0];
    var trafficArr = [];
    if(domains[i]!=''){
      var response = UrlFetchApp.fetch('http://api.semrush.com/?type=domain_rank_history&key='+APIKEY+'&export_columns=Dt,Ot&display_limit=24&display_sort=dt_desc&domain=' + domain + '&database=uk');
      var parsableData = response.getContentText();
      if(parsableData.indexOf('ERROR')<=-1){
        var array = parsableData.split('\n');
        if(dateArr.length==0){
          for(var j=1;j<array.length;j++){
            var row = array[j];
            if(row.indexOf(';')>-1){
              dateArr.push(row.split(';')[0]);
            }
          }
        }
        for(var j=1;j<array.length;j++){
          var row = array[j];
          if(row.indexOf(';')>-1){
            trafficArr.push(row.split(';')[1]);
          }
        }
        sheet.getRange(i+2,3,1 ,trafficArr.length).setValues(new Array(trafficArr.reverse()));
        
        Utilities.sleep(200);
      }else{
        keywordArr.push([errorMessage,errorMessage]);
      }
    }
    var currentTime = new Date();
    var timeDiff = currentTime - startTime;
    if(timeDiff>(270*1000)){
      ScriptApp.newTrigger("getKeywordsFromSEMRush")
      .timeBased()
      .at(new Date(currentTime.getTime()+REASONABLE_TIME_TO_WAIT))
      .create();
      break;
    }
  }
  sheet.getRange(1,3,1,24).setValues(new Array(dateArr.reverse()));
}

function updateTrafficFromSEMRush() {
  var startTime = new Date();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var domains = sheet.getDataRange().getValues().splice(1);
  var keywordArr = [];
  var dateArr = [];

  for (var i = 0; i < domains.length; i++) {
    var domain = domain[i][0];
    var trafficArr = [];
    if (domains[i] != '') {
      var response = UrlFetchApp.fetch('http://api.semrush.com/?type=domain_rank_history&key='+APIKEY+'&export_columns=Dt,Ot&display_limit=1&display_sort=dt_desc&domain=' + domain + '&database=uk');
      var parsableData = response.getContentText();
      if (parsableData.indexOf('ERROR') == -1) {
        
      }
    }
  }
}
function showChartDialog() {

  var html = HtmlService.createHtmlOutputFromFile('index')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(600)
      .setHeight(400);

  SpreadsheetApp.getUi() 
      .showModalDialog(html, 'Graph');
}

function getSpreadsheetData() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var activeCellRange = sheet.getActiveCell();
  var activeRow = activeCellRange.getRow();
  var data;
  if(activeRow>1){
    data = sheet.getRange(activeRow, 3,1,24).getValues();
    
    var rows = [];
    var URL = sheet.getRange(activeRow, 1).getValue();
    
    var header = sheet.getRange(1, 3,1,24).getValues();
    for(var i=0;i<header[0].length;i++){
      //    yyyymmdd
      var year = header[0][i].toString().substring(0,4);
      var month = header[0][i].toString().substring(4,6);
      var day = header[0][i].toString().substring(6,8);
    
      var date = new Date(year, month-1, day);
      var monthText = date.toLocaleString('en-us', { month: 'short' });
      var yearText = date.toLocaleString('en-us', { month: 'long' });
      
      rows.push([monthText + ',' + yearText,data[0][i]]);
    }
//    rows = rows.reverse();
    rows.unshift(["Month","Traffic"]);
    
    Logger.log(rows);
    return JSON.stringify([URL,rows]);
  }
  return JSON.stringify(["Please select any URL Row"]);
}