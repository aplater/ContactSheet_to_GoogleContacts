function CSVConvert() {
  
  var ss = SpreadsheetApp.getActive();
  var main = ss.getSheetByName("Main List");
  var csv = ss.getSheetByName("CSV Format");
  
  csv.getRange("A2:CK").clearContent();
  
  var data = main.getDataRange().getValues();
  
  for (m = 1; m < data.length; m++) {
    
    var rowArr = [];
    
    //name
    
    var nameArr = data[m][0].split(" ");
    
    if (nameArr.length == 1) {
      
      rowArr.push(nameArr[0], "", "");
      
    }
    
    if (nameArr.length == 2) {
    
      rowArr.push(nameArr[0], "", nameArr[1]);
      
    }
    
    if (nameArr.length == 3) {
      
      rowArr.push(nameArr[0], nameArr[1], nameArr[2]);
      
    }
    
    var groupArr = ["My Contacts;"]

    for (g = 7; g < data[m].length; g++) {
      
      var HoYCheck = data[m][g].split(" ");
      
      if (data[m][g] != "" && HoYCheck[0] != "HoY") {
        
        groupArr.push(data[m][g] + ";");

      }
      
      if (HoYCheck[0] == "HoY") {
        
        groupArr.push("HoY;");

      }
      
    }
    
    rowArr.push("", "", "", "", "", "", "", "", "", "", "", data[m][3], data[m][4], "", data[m][1], "", "", data[m][2], "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", data[m][1], "", "", "IES Älvsjö", data[m][5], data[m][7],
               "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", groupArr.join(""))
    
    csv.appendRow(rowArr);
    
  }
  
  saveAsCSV();
  
}

function saveAsCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csv = ss.getSheetByName("CSV Format");
  var range = csv.getRange("A1:CJ" + csv.getLastRow());
  var vals = range.getValues();
  var csvString = vals.join("\n");
  
  var files = DriveApp.getFolderById("1uoJX7na-_vNF_-s2r4oIak68w4hTO1H1").getFiles();

  while (files.hasNext()) {
    var file = files.next();
    
    file.setTrashed(true);
    
  }
  
  
  DriveApp.getFolderById("1uoJX7na-_vNF_-s2r4oIak68w4hTO1H1").createFile(ss.getName() + ".csv", csvString);
  
  for (r = 1; r < vals.length; r++) {
    
    var groups = vals[r][87].split(";");
    
    for (g = 0; g < groups.length; g++) {
      
      if (groups[g] == "My Contacts") {
        
        groups.splice(g, 1, "System Group: My Contacts");
        
      }
      
    }

    csv.getRange("CJ" + (r + 1)).setValue(groups.join(";"));
    
  }
  
}

