///Sidebar 
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Fill Out');
  DocumentApp.getUi().showSidebar(ui);
}

function LoadScript() {
  getTable() ;  
  return "Script Executed..!"
}
//Sidebar Finish 


//Load Excel File 

function LoadExcelFileName() 
{
  var uniQueName = DocumentApp.getActiveDocument().getName().substr(0, 8);
  var files = DriveApp.searchFiles('title contains "' + uniQueName + '" and mimeType = "application/vnd.google-apps.spreadsheet"');
  while (files.hasNext()) {
    var file = files.next();
    Logger.log(file.getId());
    Logger.log(file.getName());    
  }
  
  return file.getId(); 
}

function myFunction(realTableCountPara,rowLocationPara,year1,year2,year3,year4) {
  
  var rowLocation =0 ; 
  rowLocation = rowLocationPara ; 
  
  var id = LoadExcelFileName() ;   
  var ss = SpreadsheetApp.openById(id);  
  var sheet = ss.getSheetByName("Decisions");
  
  var startRow = 10;  // First row of data to process    
  var lastRow = sheet.getLastRow(); //Last Row 
  var numRows = lastRow-1;   // Number of rows to process
  
  
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 40);
  
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
    
  if(realTableCountPara === 1) {
    rowLocation += 8; 
  }
  else if(realTableCountPara === 2) {
    rowLocation += 11; 
  }
    else if(realTableCountPara === 3) {
    rowLocation += 14; 
  }
    else if(realTableCountPara === 4) {
    rowLocation += 18; 
  }
    else if(realTableCountPara === 5) {
    rowLocation += 19; 
  }
    else if(realTableCountPara === 6) {
    rowLocation += 20; 
  }
    else if(realTableCountPara === 7) {
    rowLocation += 21; 
  }
    else if(realTableCountPara === 8) {
    rowLocation += 22; 
  }
    else if(realTableCountPara === 9) {
    rowLocation += 23; 
  }
    else if(realTableCountPara === 10) {
    rowLocation += 24; 
  }
    else if(realTableCountPara === 11) {
    rowLocation += 27; 
  }
  
  //Insert Sheet Name on Top 
  sheet.getRange(rowLocation, 4).setValue(year1);
  sheet.getRange(rowLocation, 5).setValue(year2);
  sheet.getRange(rowLocation, 6).setValue(year3);
  sheet.getRange(rowLocation, 7).setValue(year4);
  
}


function getTable() 
{  
 
  // Get the body section of the active document.
  var body = DocumentApp.getActiveDocument().getBody();
  
  // Define the search parameters.
  var searchType = DocumentApp.ElementType.TABLE;
  
  var searchResult = null;  
  var realTableCount = 0 ; 
  
  // Search until the paragraph is found.
  while (searchResult = body.findElement(searchType,searchResult)) {
    
    var tableP = searchResult.getElement().asTable();     
    var firstCellValue = tableP.getCell(0,0).getText().trim();
    
    if( firstCellValue === "Decisions" || firstCellValue === "Decision" ) {
      
      //If table 6 is hit, only the last row should be considered for data transfer 
      var skipRows =0 ;       
      realTableCount++;
      
      if(realTableCount === 6 ) {
        skipRows =7 ; 
      }
      else {
        skipRows =0 ; 
      }
      
      var numberOfRows = tableP.getNumRows();       
       var count =0 ; 
      
      for(var i=2 + skipRows ; i<numberOfRows; i++){
        
        var firstColumn = tableP.getCell(i,0).getText(); //Not transferred 
        var thirdColumn = tableP.getCell(i,2).getText(); 
        var fourthColumn = tableP.getCell(i,3).getText(); 
        var fifthColumn = tableP.getCell(i,4).getText(); 
        var sixthColumn = tableP.getCell(i,5).getText(); 
        
        myFunction(realTableCount,count,thirdColumn,fourthColumn,fifthColumn,sixthColumn);
        count++;
        Logger.log("True"); 
      }            
    }    
  }  
}
