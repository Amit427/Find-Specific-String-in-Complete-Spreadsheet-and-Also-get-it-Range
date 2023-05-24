function getMonthEmployee() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("MASTER");
  var employees = masterSheet.getRange("A2:F").getValues().filter(f=>f[5]=="May 2023").map(m=>m[1]);
  console.log(employees);
}
