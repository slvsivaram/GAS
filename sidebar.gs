function openSidebarForm() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile("index");
  htmlOutput.setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle("Form");
  var ui = SpreadsheetApp.getUi();
  ui.showSidebar(htmlOutput);
}

function processForm(sh,itemno) {
  // here you can process the data from the form
  //Browser.msgBox(sh);
  // Browser.msgBox(itemno);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tab = ss.getSheetByName("Search");
  tab.getRange("A1").setValue(sh);
  tab.getRange("C1").setValue(itemno);
  
  var desc=tab.getRange("D4").getValue();
  var unit=tab.getRange("E4").getValue();
  var boq_qty=tab.getRange("F4").getValue();
  var boq_rate=tab.getRange("G4").getValue();
  var total_qty=tab.getRange("H4").getValue();
  var total_amt=tab.getRange("I4").getValue();
  var RA_qty=tab.getRange("J4").getValue();
  var RA_amt=tab.getRange("K4").getValue();
  var brkup=tab.getRange("L4").getValue();
  var loc=tab.getRange("M4").getValue();
  var msrmnt=tab.getRange("N4").getValue();

  var data = [desc,unit,boq_qty,boq_rate,total_qty,total_amt,RA_qty,RA_amt,brkup,loc,msrmnt];
  //Browser.msgBox(data);

  return data;
  
}
