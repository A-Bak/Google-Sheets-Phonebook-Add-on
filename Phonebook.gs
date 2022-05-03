class Phonebook {

  constructor(sheetInstance = null) {

    this.phoneNumberDict = {};

    if (sheetInstance) {
      this.fillPhoneBook(sheetInstance);
    }
  }

  /**
   * Method fills the phonebook with pairs of (PHONE_NUMBER, NAME)
   * from given Google Sheets sheet. The pairs of values must be
   * in the first and second column respectively.
   */ 
  fillPhoneBook(phonebookSheet) {
    if (!phonebookSheet) {
      throw `Invalid phonebook sheet ${phonebookSheet}.`
    }

    let data = phonebookSheet.getDataRange().getValues();
    data.forEach((row) => this.addPhoneNumber(row[0], row[1]));
  }

  /**
   * Method adds a single phone number to the phonebook.
   */
  addPhoneNumber(phoneNumber, name) {
    if (typeof phoneNumber === 'number' && typeof name === 'string') {
      this.phoneNumberDict[phoneNumber] = name;
    }
  }

  /**
   * Method replaces all of the phone numbers in first and second
   * column of the target sheet with names from the phone book. 
   */
  replacePhoneNumbers(targetSheet) {
    if (!targetSheet) {
      throw `Invalid target sheet ${targetSheet}.`;
    }

    let range = targetSheet.getDataRange(); 
    let values = range.getValues();

    for (let i = 0; i < values.length; i++) {
      values[i][0] = this.replace(values[i][0]);
      values[i][1] = this.replace(values[i][1]);
    }

    range.setValues(values);
  }
  
  /**
   * Method replaces a single phone number entry with correspoinding
   * name in the phonebook if possible. 
   */
  replace(number) {
    if (this.phoneNumberDict[number]) {
      return this.phoneNumberDict[number];
    }
    else {
      return number;
    }
  }

  /**
   * Method serialized Phonebook object to a JSON string.
   */
  toJSON() {
    return JSON.stringify(this.phoneNumberDict);
  }

  /**
   * Method deserializes Phonebook object from a JSON string.
   */
  fromJSON(jsonString) {
    this.phoneNumberDict = jsonString;
  }
}


/*******************************************************************************
*                       Save/Load Phonebook Between Calls
********************************************************************************/


/**
 * Function sets the script property 'selectedPhoneBook' to provided
 * phonebook. Used to store the phone book between script calls.
 */
function setSelectedPhonebook(phonebook) {
  let scriptProperties = PropertiesService.getScriptProperties();
  
  if (phonebook != EMPTY_PHONEBOOK) {
    scriptProperties.setProperty('selectedPhoneBook', phonebook.toJSON());
  }
  else {
    scriptProperties.setProperty('selectedPhoneBook', EMPTY_PHONEBOOK);
  }
}

/**
 * Function returns the selected phonebook if one was created.
 * Otherwise it returns EMPTY_PHONEBOOK value.
 */
function getSelectedPhonebook() {
  let scriptProperties = PropertiesService.getScriptProperties();
  let selectedPhoneBookValue = scriptProperties.getProperty('selectedPhoneBook');

  if (selectedPhoneBookValue != EMPTY_PHONEBOOK) {
    let phonebook = new Phonebook();
    phonebook.fromJSON(JSON.parse(selectedPhoneBookValue));
    return phonebook;
  }
  else {
    return EMPTY_PHONEBOOK;
  }
}


/*******************************************************************************
*                               Google Sheets UI
********************************************************************************/


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
    let phonebook = getSelectedPhonebook();
    phonebook.replacePhoneNumbers(sheet);
  }
}