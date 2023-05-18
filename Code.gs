const ss = SpreadsheetApp.getActive()
const sheet2 = ss.getSheetByName('Sheet2');
const master = ss.getSheetByName('MASTER');



function getAllSheetNames() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets();
  const sheetNames = [];

  for (let i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName();
    sheetNames.push(sheetName);
  }
Logger.log(sheetNames);

return sheetNames
}



const sheetList = getAllSheetNames();
const sheetLength = sheetList.length




function findCellsContainingText() { 

  const lastRow = sheet2.getLastRow();
  const lastColumn = sheet2.getLastColumn();
  const searchText = "Employee:";
  const cellsWithText = [];

  for (let row = 1; row <= lastRow; row++) {
    for (let column = 1; column <= lastColumn; column++) {
      const cellValue = sheet2.getRange(row, column).getValue();
      if (cellValue && cellValue.toString().includes(searchText)) {
        const cellAddress = sheet2.getRange(row, column).getA1Notation();
        cellsWithText.push(cellAddress);
      }
    }
  }
  Logger.log(cellsWithText);

var allData = []


for(i=0;i<cellsWithText.length;i++){


  Logger.log(cellsWithText[i]);


var range = sheet2.getRange(cellsWithText[i])

const currentRow = range.getRow();
  
const nextRow = currentRow + 1;

Logger.log(nextRow)

const nextRange = sheet2.getRange( "A" + nextRow);

Logger.log(nextRange.getValue())

const empData = sheet2.getRange( "D" + currentRow);

var value =  empData.getValue()
var splitValue = value.split(":")

var empCode = splitValue[0]
var empName = splitValue[1]
Logger.log(empName)

Logger.log(sheet2.getRange( 'C'+ nextRow + ':AK' + nextRow).getValues().flat())


var present = sheet2.getRange( 'C'+ nextRow + ':AK' + nextRow).getValues().flat().filter(f=>f[0]== "P")
var absent = sheet2.getRange( 'C'+ nextRow + ':AK' + nextRow).getValues().flat().filter(f=>f[0]== "A")


// var wo = sheet2.getRange( 'C'+ nextRow + ':AK' + nextRow).getValues().flat().filter(f=>f[0]== "  WO")


var wo = sheet2.getRange( 'C'+ nextRow + ':AK' + nextRow).getValues().flat().filter(f=> f.trim() === "WO")

Logger.log(present)
Logger.log(absent)
Logger.log(wo)

var data = [empCode,empName,present.length,absent.length,wo.length]

Logger.log(data)


allData.push(data)

}

Logger.log(allData)

 const numRows = allData.length;
  const numColumns = allData[0].length;

Logger.log(master.getLastRow())

master.getRange(master.getLastRow()+1,1,numRows,numColumns).setValues(allData)




  } 








