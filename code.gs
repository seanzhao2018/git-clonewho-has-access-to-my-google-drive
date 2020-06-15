var colFolderName = 2;

var colScanSubFolder = 3;

var colDateScanned = 4;

var rowToBegin = 6;



function onOpen() {

  var ui = SpreadsheetApp.getUi();

  ui.createMenu('File-Access')

      .addItem('File Access Information','scanFolder')

      .addItem('List Subfolder Tree','listFolderTree')

      .addItem('List Subfolders','listSubFolders')

      .addToUi();

}



function scanFolder() {

  

  // declare this sheet

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');

  // clear any existing contents

  var range = sheet.getRange("A2:Z5000");

  range.clearContent();



  // append a header row

  // sheet.appendRow(["Folder","Name", "Date Last Updated", "Size", "URL", "ID", "Type", "Owner", "Sharing Access", "Sharing Permission", "Editor",  "Viewer"]);

  

  for (var i = rowToBegin; i < 100; i++) {

    var folderName = SpreadsheetApp.getActiveSpreadsheet()

      .getSheetByName('Input').getRange(i, colFolderName)

      .getValue();

    var scanSubFolder = SpreadsheetApp.getActiveSpreadsheet()

      .getSheetByName('Input').getRange(i, colScanSubFolder)

      .getValue();

    var lastScan = SpreadsheetApp.getActiveSpreadsheet()

      .getSheetByName('Input').getRange(i, colDateScanned)

      .getValue();

    if (folderName == '') continue;    

    if (lastScan == '') {

      ListNamedFilesandFolders(folderName, sheet, scanSubFolder);

      SpreadsheetApp.getActiveSpreadsheet()

      .getSheetByName('Input').getRange(i, colDateScanned)

      .setValue(Date()); 

    }

  }

}





function ListNamedFilesandFolders(foldername, sheet, scanSubFolder) {



  foldername = foldername.trim();

  

  if (foldername == '') {

    sheet.getRange(2, 1).setValue("folder name is empty");

    return;

  }

  

  var foldersnext;

  

  try {

    if (foldername.toUpperCase() == "ROOT")

      foldersnext = DriveApp.getRootFolder();

    else

      foldersnext = DriveApp.getFoldersByName(foldername).next();

  }

  catch(e) {

    sheet.getRange(2, 1).setValue("folder not found");

    return;

  }



  scanFolderFile(foldersnext, sheet)



 if (scanSubFolder.toUpperCase() != 'NO') {

  var subfolders = foldersnext.getFolders();



  // now start a loop on the SubFolder list

  while (subfolders.hasNext()) {

      var mysubfolder = subfolders.next();

      scanFolderFile(mysubfolder, sheet);

      scanChildFolders(mysubfolder, sheet);

    }

  }

}



function getAllViewers(myfile) {

    var view = [];

    var viewers = myfile.getViewers();

    for (var v=0; v<viewers.length; v++) {

      view.push(viewers[v].getEmail());

    }

    view = view.join(", ");

    return view;

}



function getAllEditors(myfile) {

    var edit = [];

    var editors = myfile.getEditors();

    for (var ed=0; ed<editors.length; ed++) {

      edit.push(editors[ed].getEmail());

    }

    edit = edit.join(", ");

    return edit;

}



function getSharingAccess(myfile) {



      var privacy;

      

      var access = myfile.getSharingAccess();

      switch(access) {

        case DriveApp.Access.PRIVATE:

          privacy = "Private";

          break;

        case DriveApp.Access.ANYONE:

          privacy = "Anyone";

          break;

        case DriveApp.Access.ANYONE_WITH_LINK:

          privacy = "Anyone with a link";

          break;

        case DriveApp.Access.DOMAIN:

          privacy = "Anyone inside domain";

          break;

        case DriveApp.Access.DOMAIN_WITH_LINK:

          privacy = "Anyone inside domain who has the link";

          break;

        default:

          privacy = "Unknown";

      }

      

      return privacy;

}



function getSharingPermission(myfile) {



      var permission = myfile.getSharingPermission();

      

      switch(permission) {

        case DriveApp.Permission.COMMENT:

          permission = "can comment";

          break;

        case DriveApp.Permission.VIEW:

          permission = "can view";

          break;

        case DriveApp.Permission.EDIT:

          permission = "can edit";

          break;

        default:

          permission = "";

      }

      

      return permission;



}



function scanFolderFile (myFolder, sheet) {

  var myfiles = myFolder.getFiles();

  var data = [];

  // loop through files in this folder

  while (myfiles.hasNext()) {

    var myfile = myfiles.next();

    var fname = myfile.getName();

    var fdate = myfile.getLastUpdated(); 

    var fsize = myfile.getSize();

    var furl = myfile.getUrl();

    var fid = myfile.getId();

    var ftype = myfile.getMimeType();

    var fowner = myfile.getOwner().getEmail();

    var fsharingaccess = getSharingAccess(myfile);

    var fsharingpermission = getSharingPermission(myfile);

    var viewers = getAllViewers(myfile);

    var editors = getAllEditors(myfile);



      // Populate the array for this file

    data = [ 

        getFullFolderPath(myfile),

        fname,

        fdate,

        fsize,

        furl,

        fid,

        ftype,

        fowner,

        fsharingaccess,

        fsharingpermission,

        editors,

        viewers

      ];

    //Logger.log("data = "+data); //DEBUG

    sheet.appendRow(data);

  } // Completes listing of the files in the named folder



}





function scanChildFolders(aFolder, sheet) {



  var childFolders = aFolder.getFolders();



  while( childFolders.hasNext()){

    childFolder = childFolders.next();

    scanFolderFile(childFolder, sheet);

    scanChildFolders(childFolder, sheet);

  }

}





function getFullFolderPath(file) {

    var folders = [],

      parent = file.getParents();



    while (parent.hasNext()) {

      parent = parent.next();

      folders.push(parent.getName());

      parent = parent.getParents();

    }



    if (folders.length) {

      // Display the full folder path

      return folders.reverse().join("/");

    }

}





function getSubFolders(parent, sheet, level) {

  parent = parent.getName();

  var childFolder = DriveApp.getFoldersByName(parent).next();

  var childFolders = childFolder.getFolders();

  while(childFolders.hasNext()) {

    var child = childFolders.next();

    sheet.getRange(sheet.getLastRow() + 1,3+level).setValue(child.getName());

    getSubFolders(child, sheet, level+1);

  }

  return;

}



function listFolderTree() {

  var sheet =  SpreadsheetApp.getActiveSpreadsheet()

      .getSheetByName('Folders');

  var folderName = sheet.getRange(1, 2)

      .getValue();

      

  var range = sheet.getRange("A2:Z5000");

  range.clearContent();

  

  var parentFolder;

  

  try {

    if (folderName == "" || folderName.toUpperCase() == "ROOT") {

     parentFolder = DriveApp.getRootFolder();

     sheet.getRange(1, 2).setValue("root")

    }

    else 

     parentFolder = DriveApp.getFoldersByName(folderName).next();

    

    var childFolders = parentFolder.getFolders();

    while(childFolders.hasNext()) {

      var child = childFolders.next();

      sheet.getRange(sheet.getLastRow() + 1,3).setValue(child.getName());

      getSubFolders(child, sheet, 1);

    }

  }

  catch(e) {

    sheet.getRange(1, 2).setValue("folder not found");

    return;

  }

}



function listSubFolders() {

  var sheet =  SpreadsheetApp.getActiveSpreadsheet()

      .getSheetByName('Folders');

  var folderName = sheet.getRange(1, 2)

      .getValue();

      

  var range = sheet.getRange("A2:Z5000");

  range.clearContent();

  

  var parentFolder;

  

  try {

    if (folderName == "" || folderName.toUpperCase() == "ROOT") {

     parentFolder = DriveApp.getRootFolder();

     sheet.getRange(1, 2).setValue("root")

    }

    else 

     parentFolder = DriveApp.getFoldersByName(folderName).next();

    

    var childFolders = parentFolder.getFolders();

    while(childFolders.hasNext()) {

      var child = childFolders.next();

      sheet.getRange(sheet.getLastRow() + 1,3).setValue(child.getName());

      // getSubFolders(child, sheet, 1);

    }

  }

  catch(e) {

    sheet.getRange(1, 2).setValue("folder not found");

    return;

  }

}



