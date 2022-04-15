/**
 * Reserved name function - called when Google Sheets are opened.
 * Create a new menu button on the toolbar for NoPixel Phonebook.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("NoPixel Phonebook")
  .addItem("Select Phonebook", "selectPhonebook")
  .addItem("Replace Phone Numbers", "replacePhoneNumbers")
  .addToUi();

  setSelectedPhonebook(EMPTY_PHONEBOOK);
}


const EMPTY_PHONEBOOK = -1;


function selectPhonebook() {
  let ui = SpreadsheetApp.getUi();
  let phonebookSheetName = ui.prompt("Enter phonebook sheet name.").getResponseText();

  let sheet = SpreadsheetApp.getActive().getSheetByName(phonebookSheetName);

  if (sheet == null) {
    setSelectedPhonebook(EMPTY_PHONEBOOK);
    ui.alert(`InvalidSheetName: Sheet with name ${phonebookSheetName} was not found.`);
  }
  else{
    setSelectedPhonebook(new Phonebook(sheet));
    ui.alert(`Phonebook "${phonebookSheetName}" was created successfully.`);
  }
}


function replacePhoneNumbers() {
  let ui = SpreadsheetApp.getUi();
  let targetSheetName = ui.prompt("Enter target sheet name.").getResponseText();

  let sheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);

  if (sheet == null) {
    ui.alert(`InvalidSheetName: Sheet with name ${targetSheetName} was not found.`);
  }
  else if (getSelectedPhonebook() == EMPTY_PHONEBOOK){
    ui.alert("EmptyPhonebookError: Phonebook is currently empty, please select one of the sheets as phonebook.");
  }
  else{
    let phonebook = new Phonebook();
    phonebook.fromJSON(getSelectedPhonebook());
    phonebook.replacePhoneNumbers(sheet)
  }
}
