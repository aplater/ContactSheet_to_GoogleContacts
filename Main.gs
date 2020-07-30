function main(){
  deleteContacts();
  createGroups();
  createContacts();
  addToContact();
}

function trigger(){
  
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "main") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  for (var x = 1; x <= range.getWidth(); x++) {
    for (var y = 1; y <= range.getHeight(); y++) {
      var number = Math.floor(Math.random() * 6) + 1;   
    }
  }
  ScriptApp.newTrigger('main')
  .timeBased()
  .onWeekDay(ScriptApp.WeekDay.MONDAY)
  .atHour(number)
  .create();
}
