function share() {
  
  var ss = SpreadsheetApp.getActive()
  var classList = SpreadsheetApp.openById("1t6HpkCkWgouxNFcgCEfhdMgiECsMbfQDku7NiXu0A4Y");
  var main = ss.getSheetByName("Main List");
  var id = ss.getId();
  var editors = ss.getEditors();
  var staff = [];
  
  var data = main.getDataRange().getValues();
  
  for (i=2; i < data.length; i++) {
    
    staff.push(data[i][3]);
    
  }  
  
  for (i=0; i < staff.length; i++) {
    
    for (x=0; x < editors.length; x++) {
      
      if (staff[i] == editors[x]) {
        
        staff.splice(i, 1);
        
        var i = i - 1;
        
      }
      
    }
    
  }
  
  if (staff.length > 0) {
    
    for (i=0; i < staff.length; i++) {
      
      try {
        
        ss.addEditor(staff[i]);
        classList.addViewer(staff[i]);
        
      } catch(e) {
        
        continue;
        
      } 
      
    }
    
  }
  
}

function SFShare() {
  
  var ss = SpreadsheetApp.getActive()
  var classList = SpreadsheetApp.openById("1t6HpkCkWgouxNFcgCEfhdMgiECsMbfQDku7NiXu0A4Y");
  var main = ss.getSheetByName("Main List");
  var UL = ss.getSheetByName("Upload Format");
  
  var pictureFolder = DriveApp.getFolderById("1kOu4Sr0WaEVR_86jBPUVnUhuX7VZGmaB");
  var SFList = DriveApp.getFileById("1wkV4dUZ5M0BcU92lxkiYz0naTYl8pH0kcZCiPfQ0-9w");
  var viewersArr = SFList.getViewers();
  var editorsArr = SFList.getEditors();
  var owner = SFList.getOwner().getEmail();
  var owner2 = pictureFolder.getOwner().getEmail()
  var sharedArr = [owner, owner2];
  var staffArr = []
  
  
  for (i = 0; i < viewersArr.length; i++) {
    
    sharedArr.push(viewersArr[i].getEmail());
    
    //Logger.log(viewersArr[i].getEmail());
    
  }
  
  for (i = 0; i < editorsArr.length; i++) {
    
    sharedArr.push(editorsArr[i].getEmail());
    //Logger.log(editorsArr[i].getEmail());
    
  }
  
  var data = main.getRange("D3:D").getValues();
  
  for (i=0; i < data.length; i++) {
    
    if (data[i][0] != "") {
      
      //Logger.log(data[i]);
      
      staffArr.push(data[i][0]);
        
      //Logger.log(data[i][4]);
      
    }
    
  }
  
  for (var i = 0; i < staffArr.length; i++) {
    
    //Logger.log(staffArr[i][0]);

    for (var x = 0; x < sharedArr.length; x++) {
      
      //Logger.log(sharedArr[x]);
      
      if (staffArr[i] === sharedArr[x]) {
        
       staffArr.splice(i, 1);
        
        var i = i - 1;

        
      }
      
    }
    
  }
  
  for (i = 0; i < staffArr.length; i++) {
   
    pictureFolder.addViewer(staffArr[i]);
    SFList.addViewer(staffArr[i]);
    DriveApp.getFolderById("1uoJX7na-_vNF_-s2r4oIak68w4hTO1H1").addViewer(staffArr[i]);
    
  }
  
}


function importStaffEmails() {
 
  var ss = SpreadsheetApp.getActive();
  var main = ss.getSheetByName("Main List");
  var list = ss.getSheetByName("Resource Calendar Sync");
  
  var mainData = main.getRange("D3:D").getValues();
  var listData = list.getRange("A2:A").getValues();
  
  var mList = [];
  var uList = [];
  
  for (i = 0; i < mainData.length; i++) {
    
    if (mainData[i][0] != "" && mainData[i][0] != "john.hammer.alvsjo@engelska.se") {
      
      mList.push(mainData[i][0]);
      
    }
    
  }
   
  for (i = 0; i < listData.length; i++) {
    
    if (listData[i][0] != "") {
      
      uList.push(listData[i][0]);
      
      //Logger.log(listData[i][0])
      
    }
    
  }
  
  for (i = 0; i < mList.length; i++) {
    
    for (x = 0; x < uList.length; x++) {
     
      if (mList[i] === uList[x]) {
        
        mList.splice(i, 1);
        
        var i = i - 1;
        
      }
      
    }
    
  }
  
  for (i = 0; i < mList.length; i++) {
    
    var Avals = list.getRange("A1:A").getValues();
    var lastRow = Avals.filter(String).length;
    
    list.getRange("A" + (lastRow + 1)).setValue(mList[i]);
    
    Logger.log(mList[i]);
    
  }
  
  //calShare();
  
}

function calShare() {
  
  importStaffEmails();
  
  var ss = SpreadsheetApp.getActive();
  var list = ss.getSheetByName("Resource Calendar Sync");
  
  var calArr = [
    "engelska.se_188atfrjjhd8aipvh6o6m40skf6qm6gb6oojichj74o36dph64@resource.calendar.google.com",
    "engelska.se_1884vmvvl0322i0bkcgo2869rbpiq6g96op30d9p70s36c0@resource.calendar.google.com",
    "engelska.se_1880erdt802lcjaqnv7peuhdkdgmk6gb6op32e1h6gs38cpp6k@resource.calendar.google.com",
    "engelska.se_18834t094e0raj60ifbut4oclftla6gb6op34e9i70s3ic9l64@resource.calendar.google.com",
    "engelska.se_18892d998diruhboit1lqedhq93us6gb6op36c9j6sojgcpp60@resource.calendar.google.com",
    "engelska.se_188as20aknf88hbpkqut23pl413b66g96op38e1o64qjedg@resource.calendar.google.com",
    "engelska.se_1880te1fh3su2i4nl8h5hjk9vjbam6gb6op3acpm6ssjidpl68@resource.calendar.google.com",
    "engelska.se_1881r75l23q0ugg5h9b9epdkn4upc6gb6cr3ge1j74q34chp6g@resource.calendar.google.com",
    "engelska.se_188er10c1mfs6hgrgjrk7u0le7igg6gb6op3edhg6kp32cpk6o@resource.calendar.google.com",
    "engelska.se_188c42r20gckoil0kh4m98nbtfsgu6gb6krj6dpm6cq36dho74@resource.calendar.google.com",
    "engelska.se_188aunm2hnb8misnlp50qk5s724no6gb6gp3ae9j60rjidpp6s@resource.calendar.google.com",
    "engelska.se_1887mno795td0jqskhjgl2f9edfeq6gb6krjccph6crj8d1k6s@resource.calendar.google.com",
    "engelska.se_1888r8sf36vjgh6dk1q4q68bdpvhc6g96krjgd9i6ks34cg@resource.calendar.google.com",
    "engelska.se_188dt6ufolon8ggonp0l3tb21ml4o6gb6krjidpg6sqj4e9l6g@resource.calendar.google.com",
    "engelska.se_188869ngceg86hh6j39ker5g0ktte6gb6ks30e1m60s3cdhh6o@resource.calendar.google.com",
    "engelska.se_188cv8s0bv4mshfmlo9sa61v303l46gb6gp3cchg70pjgdpj6o@resource.calendar.google.com",
    "engelska.se_188asjdms5lbshfpnulahj5tv8l6u6gb68sj4chn64p3idpg6g@resource.calendar.google.com",
    "engelska.se_18892rr7cihggho4j47ksgchg2h8e6ga68sj2dpj6cr3gd1m@resource.calendar.google.com"
  ]  
  
  Logger.log(calArr.length);
  
  var listData = list.getRange("A2:B").getValues();
  
  for (i=0; i < listData.length; i++) {
    var toggle = 0;
    
    if (listData[i][0] != "" && listData[i][1] == false) {           
      
      var userEmail = listData[i][0];
      var resource = {
        'scope': {
          'type': 'user',
          'value': userEmail
        },
        'role': 'reader'
      };
      
      for (x = 0; x < calArr.length; x++) {
        Calendar.Acl.insert(resource, calArr[x]);
        
        var toggle = toggle + 1;
        Logger.log(toggle);
        
        if (toggle == calArr.length) {
          
          list.getRange("B" + (i + 2)).setValue(true);
          
        }
        
      }
      
    } 
    
  }
  
}

function calStaffShare() {
  
  var ss = SpreadsheetApp.getActive();
  var list = ss.getSheetByName("Main List");
  
  var listData = list.getRange("D2:D").getValues();
  
  for (i=0; i < listData.length; i++) {
    
    if (listData[i][0] != "") {
      
      CalendarApp.subscribeToCalendar(listData[i][0]);
      
      Logger.log(listData[i][0]);
 
        var cal = listData[i][0];
        
        var resource = {
          'scope': {
            'type': 'user',
            'value': "it.representative.alvsjo@engelska.se"
          },
          'role': 'writer'
        };
        
        Calendar.Acl.insert(resource, cal);

    }
  }
}

function shareButton() {
  
  var cal = Session.getActiveUser().getEmail();
  
  Logger.log(cal);
  
  var resource = {
    'scope': {
      'type': 'user',
      'value': "it.representative.alvsjo@engelska.se"
    },
    'role': 'owner'
  };
  
  Calendar.Acl.insert(resource, cal);
  
}

