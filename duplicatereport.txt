/*Script to check for duplicate email submissions by email value and submit a report
Author: Shawn Hodgson
Last Update: 2/22/2018
http://shawnjhodgson.com/2014/08/forms-sending-a-report-when-a-duplicate-submission-is-made/
*/
// menu added on open
function onOpen() {
  FormApp.getUi() // Or DocumentApp or FormApp.
  .createMenu('Settings')    
  .addItem('Authorize', 'authorize')
  .addItem('Set Email', 'setEmail')
  .addToUi();
}

//easily authorize the script to run from the menu
function authorize(){  
  var respEmail = Session.getActiveUser().getEmail();  
  MailApp.sendEmail(respEmail,"Form Authorizer", "Your form has now been authorized to send you emails");
}

//set the email of the form respondant
function setEmail(){
  var ui = FormApp.getUi(); // Same variations.  
  var result = ui.prompt(
    'Duplicate Report',
    'Please enter email recipient:',
    ui.ButtonSet.OK_CANCEL);  
  // Process the user's response.
  var button = result.getSelectedButton();
  var emailAddress = result.getResponseText();
  if (button == ui.Button.OK) {
    setProperty('EMAIL_ADDRESS',emailAddress);   
  } 
}


//get user email/utility function
function getUserEmail(response){
  var itemRes = response.getItemResponses();
  for (var i = 0; i < itemRes.length; i++){
    var respQuestion = itemRes[i].getItem().getTitle();
    var index = itemRes[i].getItem().getIndex();  
    var email = respQuestion.toLowerCase();
      var regex = /.*email.*/;
    if(regex.test(email) == true){
      email = itemRes[i].getResponse();  
      return [email, index];     
    }        
  }    
}

//setting script properties
function setProperty(key,property){
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(key,property);
}

//getting script properties
function getProperty(property){
  var scriptProperties = PropertiesService.getScriptProperties();
  var savedProp = scriptProperties.getProperty(property);  
  return  savedProp;
}

//controller function
function onFormSubmit(e){
  Logger.log("Submitted");
  var response = e.response;
  var subEmail = getUserEmail(response); 
  //Wait 10 seconds for values ot hit spreadsheet
  Utilities.sleep(10000);
  //CHANGE Sheet Name to your submission sheet name
  var values = getValues("Sheet Name");
  var dubArray = checkForDups(subEmail[0],subEmail[1],values);
  //If more than a single submission is found built and submit email report
  Logger.log("dubArray Length " + dubArray.length);
  if(dubArray.length > 1){  
    var headers = getHeaders(values);
    var emailBody = mailFormat(dubArray,headers);
    var emailAddress =  getProperty('EMAIL_ADDRESS');
    Logger.log("Sending Email");
    MailApp.sendEmail(emailAddress,"Duplicate Report","", {htmlBody: emailBody});
  }  
}

 //grabbing general settings from spreadsheet sheet Settings
function getValues(sheet) {
  var form = FormApp.getActiveForm();
  var ssID = form.getDestinationId();
  var ss = SpreadsheetApp.openById(ssID);
  var sheet = ss.getSheetByName(sheet);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(1,1,lastRow,lastColumn);
  var values = range.getValues();
  return values;  
}

//takes the email value, email position, and submitted form values and checks if the entered email has submitted more than once
function checkForDups(subEmail,emailPos,values){
  var dubArray = [];
  var pos = parseInt(emailPos)-1;
  for(i in values){    
    if (subEmail == values[i][pos]){   
      Logger.log("True " + subEmail + " " + values[i][pos]);
      dubArray.push(values[i]);      
    }
  }
  return dubArray;
}

//function to get headers
function getHeaders(values){
  var headers = values[0];
  return headers;
}


//function to build email message
function mailFormat(dubArray,headers){
  var tableStart = "<html><body><table border=\"1\">";
  var tableEnd = "</table></body></html>";  
  dubArray = addHTML(dubArray,false); 
  headers =  addHTML(headers,true);// using a function that works on multi dimension array to work on a 1 dimension array. not gonna work  
  dubArray.unshift(headers);  
  dubArray = tableStart + dubArray.join('') + tableEnd;
  return dubArray;
}

//function to add html table elements to headers and rows/cells
function addHTML(plainMSG,ifHeader){
  var rowStart = "<tr>";
  var rowEnd = "</tr>";
  var cellStart = "<td>";
  var cellEnd = "</td>";
  if (ifHeader == false){
    for (i in plainMSG){   
      for (j in plainMSG[i]){
        plainMSG[i][j] = cellStart + plainMSG[i][j] + cellEnd;       
      }
      plainMSG[i] = rowStart + plainMSG[i].join('') + rowEnd;
    }
    return plainMSG;
  }
  else if(ifHeader == true){
    for (i in plainMSG){ 
      plainMSG[i] = cellStart + plainMSG[i] + cellEnd;
    }
    plainMSG  = rowStart + plainMSG.join('') + rowEnd;
    return plainMSG;
  }
}

//function to send out mail
function mailSendSimple(emailAddr,emailSub, emailMsg){  
  MailApp.sendEmail(emailAddr,emailSub,"", {htmlBody: emailMsg}); 
}