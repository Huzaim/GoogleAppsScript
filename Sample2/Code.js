function onOpen() {
 
 //Waiting time 
 var time = 20; 
  
 var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();  
 spreadSheet.toast("This message will disappear after " + time + " seconds");
  
 Utilities.sleep(time*1000);  
 spreadSheet.toast("We are now sending this private note to the shredder");

 spreadSheet.getActiveSheet()
   .getRange(1, 1, spreadSheet.getLastRow(), spreadSheet.getLastColumn()).clear();  
}
