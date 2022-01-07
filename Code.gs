/**
 * 
 * Author: Nick Young - @techupover | techupover.com
 * 
 * 
 * CREDITS
 * Kudos to https://www.pbainbridge.co.uk/2020/06/bulk-create-google-drive-folders-20.html
 * and various stackoverflow threads for filling in some gaps about how things work
 */

/**
 * 
 * CHANGE THESE VARIABLES PER YOUR PREFERENCES
 * 
 */

const CUSTOM_MENU_NAME = 'Neat Features'; // the name of the new menu in your google sheet
const CUSTOM_MENU_ITEM = 'Generate Folders Now'; // the menu item you'll click to run the script
const DATA_TAB_NAME = 'folder_data'; // name for the sheet tab that contains the folder info that will be created/processed
const LOG_TAB_NAME = 'log'; // name of the sheet tab that will store the log messages from this script
const DATE_FORMAT = 'yyyy-MM-dd HH:mm:ss'; // date format to use for the log entries. Probably dont change this unless you really really want to.

/**
 * 
 * DO NOT CHANGE ANYTHING UNDER THIS LINE
 * 
 * ONLY CHANGE THINGS IN THE CONFIG.GS FILE
 * 
 */

/**
 * When the spreadsheet is open, add a custom menu
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var customMenuItems = [
    {name: CUSTOM_MENU_ITEM, functionName: 'processGoogleSheetData_'}
  ];
  spreadsheet.addMenu(CUSTOM_MENU_NAME, customMenuItems);
}

/**
* Bulk create Google Folders from data within a Google Sheet.
* This is the function to call from the custom menu item.
* Others are referenced by this one.
*/
function processGoogleSheetData_() {
  
  // get current spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Log starting of the script
  logEventToGoogleSheet('Script has started');

  // display Toast notification
  ss.toast('Starting!!!', 'Yep, we are starting...now.');
  
  // get TimeZone
  var timeZone = ss.getSpreadsheetTimeZone();
  
  // get Data sheet
  var dataSheet = ss.getSheetByName(DATA_TAB_NAME);
  
  // get all data as a 2-D array
  var data = dataSheet.getDataRange().getValues();
  
  // create a name:value pair array to send the data to the next Function
  var spreadsheetData = {ss:ss, timeZone:timeZone, dataSheet:dataSheet, data:data};
  
  // run Function to create Google Folders
  var doCreateFolders = createFolders(spreadsheetData);
  
  // check success status
  if (doCreateFolders) {
    // display Toast notification
    ss.toast('Script complete', 'Finished');
  }
  else {
    // script completed with error
    // display Toast notification
    ss.toast('Got some errors. Check the logs', 'Finished');
  }
  
  // Log starting of the script
  logEventToGoogleSheet('Script finished');
  
  
}


/**
* Loop through each row and create folders, set permissions
*/
function createFolders(spreadsheetData) {
  
  // extract data from name:value pair array
  var ss = spreadsheetData['ss'];
  var timeZone = spreadsheetData['timeZone'];
  var dataSheet = spreadsheetData['dataSheet']; 
  var data = spreadsheetData['data'];

  // get last row number so we know when to end the loop
  var lastRow = dataSheet.getLastRow();

  var folderIdMap = new Object();

  // start of loop to go through each row iteratively
  for (var i=1; i<lastRow; i++) {
    
    // extract values from row of data for easier reference below
    var rootFolderId = data[i][0];
    var templateDocId = data[i][1]
    var newFolderName = data[i][2];
    var parentFolderName = data[i][3];
    var folderId = data[i][4];
    var permissionEmailViewer = data[i][5];
    var permissionEmailEditor = data[i][6];
    
    // only perform this row if the folder ID is blank
    if(folderId == '') {

      // if the sheet doesn't have a specified parent folder name, then it goes in the root.
      // if the parent folder name is supplied then it had to have been created by this script, 
      // so get it from script properties (where the ID is saved during an earlier loop)
      if(parentFolderName == '') {
        destinationFolderId = rootFolderId;
        
      } else {
        
        var thisMapKey = createMapString(rootFolderId + '___' + parentFolderName);
        var destinationFolderId = folderIdMap[thisMapKey];
        
      }
      
      // display Toast notification
      ss.toast(newFolderName, 'Creating New Folder');

      // run Function to create Google Folder and return its URL/ID
      var folderDetails = createFolder(newFolderName, destinationFolderId);

      // check new Folder created successfully
      if (folderDetails) {

        // extract Url/Id for easier reference later
        var newFolderUrl = folderDetails['newFolderUrl'];
        var newFolderId = folderDetails['newFolderId'];

        //push the key/folder id to the array map so we can use it later in the loop
        var thisMapKey = createMapString(destinationFolderId + '___' + newFolderName);
        folderIdMap[thisMapKey] = newFolderId;

        // copy the template doc into the new directory (if specified)
        if(templateDocId != '') {
          makeCopy(templateDocId, newFolderId, newFolderName + ' Template Document');
        } 
                
        // set the Folder ID value in the google sheet, inserting it as a link
        var newFolderLink = '=HYPERLINK("' + newFolderUrl + '","' + newFolderName + '")';
        dataSheet.getRange(i+1, 5).setFormula(newFolderLink);
        
        // check if Viewer Permissions need adding - if there are emails in the column for this row
        if (permissionEmailViewer != '') { 
          
          // run Function to add Folder permissions
          var currentRow = i+1;
          var addPermissionsFlag = addPermissions('VIEWER', timeZone, dataSheet, permissionEmailViewer,
                                                  newFolderId, currentRow, 8);
          
          // if problem adding Permissions return for status message
          if (addPermissionsFlag == false) {
            // display Toast notification and return false flag
            ss.toast('Error when adding Viewer Permissions to: ' + newFolderName, 'Error');
            return false;
          }
          
        }

        // check if Editor Permissions need adding - if there are emails in the column for this row
        if (permissionEmailEditor != '') { 
          
          // run Function to add Folder permissions
          var currentRow = i+1;
          var addPermissionsFlag = addPermissions('EDITOR', timeZone, dataSheet, permissionEmailEditor,
                                                  newFolderId, currentRow, 9);
          
          // if problem adding Permissions return for status message
          if (addPermissionsFlag == false) {
            // display Toast notification and return false flag
            ss.toast('Error when adding EDITOR Permissions to: ' + newFolderName, 'Error');
            return false;
          }
          
        }
        
        // write all pending updates to the google sheet using flush() method
        SpreadsheetApp.flush();
        
      } else {
        // write error into 'Permission Added?' cell and return false value
        dataSheet.getRange(i+1, 4).setValue('Error creating folder. Please see Logs');
        // new Folder not created successfully
        return false;
      }

    } else {

      ss.toast('Skipping Row - Folder ID already set', 'Moving onto the next row to process');

    }
    
  } // end of loop to go through each row in turn **********************************
  
  // completed successfully
  return true;
  
  
}

function makeCopy(sourceDocumentId, destinationFolderId, destinationFileName) {

  var destinationFolder = DriveApp.getFolderById(destinationFolderId);
  
  return DriveApp.getFileById(sourceDocumentId).makeCopy(destinationFileName,destinationFolder);

}


/**
 * Function to create new Google Drive Folder and return details (url, id)
*/

function createFolder(folderName, destinationFolderId) {
  
  try {
    // get destination Folder
    var destinationFolder = DriveApp.getFolderById(destinationFolderId);
  }
  catch(e) {
    logEventToGoogleSheet('Error getting destination folder: ' + e + e.stack);
    var destinationFolder = false;
  }
  
  
  // proceed if successfully got destination folder
  if (destinationFolder) {
    var documentProperties = PropertiesService.getDocumentProperties();
    
    try {
      // create new Folder in destination
      var newFolder = destinationFolder.createFolder(folderName);
      // get new Drive Folder Url/Id and return to Parent Function
      var newFolderUrl = newFolder.getUrl();
      var newFolderId = newFolder.getId();
      var folderDetails = {newFolderUrl:newFolderUrl, newFolderId:newFolderId};
      
      return folderDetails;
    }
    catch(e) {
      logEventToGoogleSheet('Error creating new Folder: ' + e + e.stack);
      return false;
    }
  }
  else {
    // return false as unable to get destination folder
    return false;
  }
}


/**
 * Function to add permissions to each Folder using the provided email address(es).
 * 
 * role var can be either VIEWER or EDITOR
*/

function addPermissions(role, timeZone, dataSheet, permissionEmail, newFolderId, currentRow, permAddedCol) {
  
  // split up email address array to be able to loop through them separately
  var emailAddresses = permissionEmail.split(',');
  logEventToGoogleSheet(role + ' emailAddress array is: ' + emailAddresses);
  
  // get length of array for loop
  var emailAddressesLength = emailAddresses.length;
  
  
  try {
    // get Google Drive Folder
    var newFolder = DriveApp.getFolderById(newFolderId);
  }
  catch(e) {
    logEventToGoogleSheet('Error getting destination folder: ' + e + e.stack);
    var newFolder = false;
  }
  
  
  // proceed if successfully got destination folder
  if (newFolder) {
    
    // loop through each email address and add as 'Editor' *******************
    for (var i=0; i<emailAddressesLength; i++) {
      
      var emailAddress = emailAddresses[i].trim();
      logEventToGoogleSheet(role + ' emailAddress for adding permission is: ' + emailAddress);
      
      try {
        
        if(role == 'VIEWER') {

          logEventToGoogleSheet('Adding ' + emailAddress + ' as ' + role);
          newFolder.addViewer(emailAddress);
          var success = true;

        } 

        if(role == 'EDITOR') {

          logEventToGoogleSheet('Adding ' + emailAddress + ' as ' + role);
          newFolder.addEditor(emailAddress);
          var success = true;

        }
        // add 'Edit' permission using email address

        if (success) {
          // write timestamp into 'Permission Added?' cell
          var date = new Date;
          var timeStamp = Utilities.formatDate(date, timeZone, DATE_FORMAT);
          dataSheet.getRange(currentRow, permAddedCol).setValue(timeStamp);
        }
        else {
          // write error into 'Permission Added?' cell and return false value
          dataSheet.getRange(currentRow, permAddedCol).setValue('Error adding ' + role + ' permission. Please see Logs');
          return false;
        }
        
      }
      catch(e) {
        logEventToGoogleSheet('Error adding ' + role + ' permission: ' + e + e.stack);
      }
      
    }
    
  }
  else {
    // write error into cell and return false value
    dataSheet.getRange(currentRow, permAddedCol).setValue('Error getting folder. Please see Logs');
    // return false as unable to get Google Drive Folder
    return false;
  }
  
  
  // return true as all permissions added successfully
  return true;
  
  
}


/** 
* Write log message to Google Sheet
*/

function logEventToGoogleSheet(text_to_log) {
  
  // get the user running the script
  var activeUserEmail = Session.getActiveUser().getEmail();
  
  // get the relevant spreadsheet to output log details
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var googleSheet = ss.getSheetByName(LOG_TAB_NAME);
  
  // create and format a timestamp
  var now = new Date();
  var timeZone = ss.getSpreadsheetTimeZone();
  var niceDateTime = Utilities.formatDate(now, timeZone, DATE_FORMAT);
  
  // create array of data for pasting into log sheet
  var logData = [niceDateTime, activeUserEmail, text_to_log];
  
  // append details into next row of log sheet
  googleSheet.appendRow(logData);
  
}

/**
 * Create a string that can easily be used as an object/array key for reference
 * during the create folder loop. (Helps with identifying the parent folder dynamically)
 */
function createMapString(input_string) {

  return input_string.replace(/\W/g, '');

}
