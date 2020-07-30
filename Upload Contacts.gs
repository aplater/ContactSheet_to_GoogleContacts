function deleteContacts() {
  
  var contacts = ContactsApp.getContactsByEmailAddress('.alvsjo@engelska.se');
  var ss = SpreadsheetApp.getActive();
  var csv = ss.getSheetByName("CSV Format");
  var csvData = csv.getDataRange().getValues();
  
  Logger.log(contacts.length);
  
  for (i = 0; i < contacts.length; i++) {
    
    var toggle = 0;
    
    try {
      
      var cEmail = contacts[i].getPrimaryEmail();
      
      for (c = 1; c < csvData.length; c++) {
        
        var email = csvData[c][14];
        
        if (email != "") {
          
          if (cEmail == email) {
            
            var toggle = 1;
            
          }
        } 
      }
      if (toggle == 0) {
        contacts[i].deleteContact();
      }
    } catch(e) {
      continue; 
    }
  } 
}

function createGroups() {
  
  var ss = SpreadsheetApp.getActive();
  var csv = ss.getSheetByName("CSV Format");
  var background = ss.getSheetByName("Background");
  
  var cGroups = ContactsApp.getContactGroups();
  
  var dGroups = [];
  
  for (g = 0; g < cGroups.length; g++) {
    
    dGroups.push(cGroups[g].getName());
    
  }
  
  var bData = background.getRange("D2:D").getValues();
  
  var uGroups = ["HoY"];
  
  for (b = 0; b < bData.length; b++) {
    
    var split = bData[b][0].split(" ");
    
    if (split[0] != "HoY") {
      
      uGroups.push(bData[b][0]);
      
    }
    
  }

  for (b = 0; b < uGroups.length; b++) {
    
    var toggle = 0;
    
    for (g = 0; g < uGroups.length; g++) {
      
      if(uGroups[b] == dGroups[g]) {
        
        var toggle = 1;
        
      }
      
    }
    
    if (toggle == 0 && uGroups[b] != "") {
      
      ContactsApp.createContactGroup(uGroups[b]);

    }
    
  }
  
}


function createContacts() {
  
  var ss = SpreadsheetApp.getActive();
  var csv = ss.getSheetByName("CSV Format");
  var background = ss.getSheetByName("Background");
  
  var csvData = csv.getDataRange().getValues();
  var bData = background.getRange("D2:D").getValues();
  
  for (c = 1; c < csvData.length; c++) {
    
    try {
      
      if (csvData[c][14] != "") {
        
        var contact = ContactsApp.getContact(csvData[c][14]);
        
        if (csvData[c][0] > "" && contact == null) {
          
          var contact = ContactsApp.createContact(csvData[c][0], csvData[c][2], csvData[c][14]);
          
        }
        
        var masterGroupArr = [];
        var contactGroupArr = [];
        var joinedGroupArr = [];
        
        for (b = 0; b < bData.length; b++) {
          
          if (bData[b][0] != "") {
            
            if (group != null) {
              
              masterGroupArr.push(bData[b][0]);
              
            }
            
          }
          
        }
        
        var groups = contact.getContactGroups();
        
        for (g = 0; g < groups.length; g++) {
          
          contactGroupArr.push(groups[g].getName());
          
        }
    
        var groups = csvData[c][87].split(";");
        for (g = 0; g < groups.length - 1; g++) {
          
          if (groups[g] != "") {
            
            var group = ContactsApp.getContactGroup(groups[g]);
            
            if (group.getName() == groups[g]) {
              
              contact.addToGroup(group);
              
              joinedGroupArr.push(groups[g]);
              
            }
            
          }
          
        }
        
        for (cg = 0; cg < contactGroupArr.length; cg++) {
          
          for (jg = 0; jg < joinedGroupArr.length; jg++) {
            
            if (contactGroupArr[cg] == joinedGroupArr[jg]) {
              
              contactGroupArr.splice(cg, 1);
              
              var cg = cg - 1;
              
            }
            
          }
          
        }
        
        if (contactGroupArr.length > 0) {
          
          for (cg = 0; cg < contactGroupArr.length; cg++) {
            
            for (mg = 0; mg < masterGroupArr.length; mg++) {
              
              if (contactGroupArr[cg] == masterGroupArr[mg]) {
                
                var group = ContactsApp.getContactGroup(contactGroupArr[cg]);
                
                contact.removeFromGroup(group);
                
              }                
              
            }
            
          }
          
        }
        
      }
      
    } catch(e) {
      
      var c = c - 1;
      
      //Logger.log(e);
      
      continue;
      
    }
    
  }
  
}

function addToContact() {
  
  var ss = SpreadsheetApp.getActive();
  var csv = ss.getSheetByName("CSV Format");
  var background = ss.getSheetByName("Background")
  
  var csvData = csv.getDataRange().getValues();
  
  for (c = 1; c < csvData.length; c++) {
    
    try {
      
      if (csvData[c][14] != "") {
        
        var contact = ContactsApp.getContact(csvData[c][14]);
        
        if (csvData[c][1] != "") {
          
          contact.setMiddleName(csvData[c][1]);
          
        }
        
        var email = contact.getEmails();
        
        var toggle = 0;
        
        for (i = 0; i < email.length; i++) {
          
          if (email[i].getAddress() == csvData[c][14]) {
            
            email[i].setLabel(ContactsApp.Field.WORK_EMAIL).setAsPrimary();
            
          }
          
          if (email[i].getAddress() == csvData[c][15] && csvData[c][15] > "") {
            
            email[i].setLabel(ContactsApp.Field.HOME_EMAIL);
            
            var toggle = 1;
            
          }
          
        }
        
        if (csvData[c][15] > "" && toggle == 0) {
          contact.addEmail(ContactsApp.Field.HOME_EMAIL, csvData[c][15]);
        }
        
        var phone = contact.getPhones();
        
        var toggle1 = 0;
        var toggle2 = 0;
        
        for (i = 0; i < phone.length; i++) {
          
          var number = phone[i].getPhoneNumber();
          
          if (number == csvData[c][17] && csvData[c][17] > "") {
            
            phone[i].setLabel(ContactsApp.Field.WORK_PHONE).setAsPrimary()
            
            var toggle1 = 1
            
            }
          
          if (number == "073 917 42 98" && csvData[c][20] > "" && csvData[c][0] == "John" &&  csvData[c][2] == "Hammer") {
            
            phone[i].deletePhoneField();
            
          }
          
          if (number == csvData[c][20] && csvData[c][20] > "") {
            
            phone[i].setLabel(ContactsApp.Field.MOBILE_PHONE)
            
            var toggle2 = 1
            
            }
          
        }
        
        if (csvData[c][17] > "" && toggle1 == 0) {
          
          contact.addPhone(ContactsApp.Field.WORK_PHONE, csvData[c][17]).setAsPrimary();
          
        }
        
        if (csvData[c][20] > "" && toggle2 == 0) {
          
          contact.addPhone(ContactsApp.Field.MOBILE_PHONE, csvData[c][20]);
          
        }
        
        var company = contact.getCompanies();
        
        for (i = 0; i < company.length; i++) {
          
          if (company[i].getCompanyName() == csvData[c][42]) {
            
            company[i].deleteCompanyField();
            
          }
          
        }
        
        contact.addCompany(csvData[c][42], csvData[c][43]);
        
      }
      
    } catch(e) {
      
      var c = c - 1;
      
      Logger.log(e);
      
    }
    
  }
  
}

function importSuspended() {
  
  var users = [];
  var options = {
    domain: "engelska.se",     // Google Apps domain name 
    maxResults: 100,
    projection: "full",      // Fetch full details of users
    query: "isSuspended=true",
    orderBy: "email"          // Sort results by users
  }
  
  do {
    
    var response = AdminDirectory.Users.list(options);
    response.users.forEach(function(user) {
      users.push([user.name.fullName, user.primaryEmail, user.organizations]); 
    });
    
    // For domains with many users, the results are paged
    if (response.nextPageToken) {
      options.pageToken = response.nextPageToken;
    }
  } while (response.nextPageToken);
  
  // Insert data in a spreadsheet
  var ss = SpreadsheetApp.getActive();
  var susp = ss.getSheetByName("Suspended");
  var main = ss.getSheetByName("Main List");
  var background = ss.getSheetByName("Background");
  susp.clearContents();
  susp.getRange(1,1,users.length, users[0].length).setValues(users);
  
  var mData = main.getDataRange().getValues();
  var sData = susp.getDataRange().getValues();
  
  background.getRange("B2:B").clearContent();
  
  var num = 2;
  
  for (m = 1; m < mData.length; m++) {
    
    for (s = 0; s < sData.length; s++) {
      
      if (mData[m][3] != "" && mData[m][3] == sData[s][1]) {
        
        background.getRange("B" + num).setValue(mData[m][3]);
        
        var num = num + 1;
        
      }
      
    }
    
  }
  
  var bData = background.getDataRange().getValues();
  
  var num = 1;
  
  for (b = 1; b < bData.length; b++) {
    
    if (bData[b][1] != "") {
      
      for (m = 1; m < mData.length; m++) {
        
        if (bData[b][1] == mData[m][3]) {
          
          main.deleteRow(m + num);
          
          var num = num - 1;
          
        }
        
      }
      
      background.getRange("B" + (b + 1)).clearContent();
      
    }    
    
  }
    
}