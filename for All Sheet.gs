function onOpen(e){
const ui = SpreadsheetApp.getUi()
ui.createMenu("Automation")
.addItem("Filter DATA" , "forAllSheet")
.addToUi()
}

var ss = SpreadsheetApp.getActive()
var masterSheet = ss.getSheetByName('MASTER');

const sheetList = getAllSheetNames();
const sheetLength = sheetList.length

var masterData = masterSheet.getRange(2,1,masterSheet.getLastRow(),6).getValues()

var masterNames = arrayOfMaster()[0]
var masterDates = arrayOfMaster()[1]
var newAllData = []



function forAllSheet() { 

masterSheet.getRange('F2:F').setNumberFormat('@')

for(i=0;i<sheetLength;i++){

if(sheetList[i] !== "MASTER"){



const sheet2 = ss.getSheetByName(sheetList[i]);
// Logger.log(sheetList[i])
const dateString = sheet2.getRange('B3').getValue().toString()
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
  // Logger.log(cellsWithText);

var [allData, allEmp, newData] = [[], [], []];

for(s=0;s<cellsWithText.length;s++){


  // Logger.log(cellsWithText[i]);


var range = sheet2.getRange(cellsWithText[s])

const currentRow = range.getRow();
  
const nextRow = currentRow + 1;

// Logger.log(nextRow)

// Logger.log(nextRange.getValue())

const empData = sheet2.getRange( "D" + currentRow);

var value =  empData.getValue()
var splitValue = value.split(":")

var empCode = splitValue[0]
var empName = splitValue[1]



var present = sheet2.getRange( 'C'+ nextRow + ':AK' + nextRow).getValues().flat().filter(f=>f[0]== "P")
var absent = sheet2.getRange( 'C'+ nextRow + ':AK' + nextRow).getValues().flat().filter(f=>f[0]== "A")

var wo = sheet2.getRange( 'C'+ nextRow + ':AK' + nextRow).getValues().flat().filter(f=> f.trim() === "WO")

// Logger.log(present)
// Logger.log(absent)
// Logger.log(wo)

var month = dateString.split(" ")[0].toString() + " " + dateString.split(" ")[2].toString()


var data = [empCode,empName,present.length,absent.length,wo.length,month.toString()]


allData.push(data)

}



Logger.log(month)

var employeeName = masterData.filter(d=>d[5] == month).map(m=>m[1])
Logger.log(employeeName)

allData.forEach(d=>{
  if(masterDates.indexOf(d[5]) == -1){

newAllData.push(d)

  }else if(employeeName.indexOf(d[1]) == -1 ){


      newAllData.push(d)

  }

}
)


Logger.log(newAllData)

  }
  }
  
try{
 
 const numRows = newAllData.length;
  const numColumns = newAllData[0].length;

master.getRange(master.getLastRow()+1,1,numRows,numColumns).setValues(newAllData)
SpreadsheetApp.getActive().toast('Data is Added in Master' , "Data Updated" , 5)
}
catch{

SpreadsheetApp.getActive().toast('Nothing to Add in Master', "Data Updated" , 5 )

}
  
  
  }




function arrayOfMaster(){
  // Logger.log(masterData)

  // const nameArray = masterData.map(d=>d[1])
  const nameArray = masterData.map(
    function(d){
return          d[1].toString()
    }

  )
  const dateArray = masterData.map(d=>d[5])


  return [nameArray,dateArray]
}




