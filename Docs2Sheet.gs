// Code to replace Docs with Data from Sheets -- thanks @https://jeffreyeverhart.com/ // 

// Function to create the menu when opening the SpreadSheet 
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // Defining first Menu and Item inside the Menu
  const menu = ui.createMenu('Additional Tools');
  menu.addItem('Create and Replace', 'createNewGoogleDocs');
  menu.addToUi();
  }
// Function on clicking the Button defined Above in The Menu

function createNewGoogleDocs () {
  // Define the Docs Template, The folder to store and the Sheet where data is stored
const PsoTemplate = DriveApp.getFileById('16f4S2CXFocIdOmc');
const destinationFolder = DriveApp.getFolderById('1P6uBhC1xT0Vnpo');
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customer_Data');
const rows = sheet.getDataRange().getValues();
//Logger.log(rows[0][1]);

  // Make a copy of the Docs template, open the new document and load the body in a variable
const copy = PsoTemplate.makeCopy(`${rows[3][1]}_Name`, destinationFolder);
const doc = DocumentApp.openById(copy.getId());
const body = doc.getBody();

// Replace in the Body of the document the following words with the values in rows[row][column]
body.replaceText('{{CustomerName}}', rows[3][1]);


// Save the doc , get the URL of the doc and paste the url in the cell (row 1, column 4)
doc.saveAndClose();
const url = doc.getUrl();
var linktodoc = SpreadsheetApp.newRichTextValue().setText('Link to Doc').setLinkUrl(url).build();
sheet.getRange(3,3).setRichTextValue(linktodoc);  

}
