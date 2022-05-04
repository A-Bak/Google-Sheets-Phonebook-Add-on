class Phonebook {

  constructor(sheetInstance = null) {

    this.phoneNumberDict = {};
    this.styleDict = {};

    if (sheetInstance) {
      this.fillPhoneBook(sheetInstance);
    }
  }

  /**
   * Method fills the phonebook from given Google Sheets sheet.
   * The pairs of values must be in the first and second column
   * respectively.
   * 
   * The method fills phoneNumberDict with pairs [number, name]
   * and styleDict with pairs [number, phoneNameStyle], where
   * phoneNameStyle contains fontSize, fontColor, backgroundColor.
   */ 
  fillPhoneBook(phonebookSheet) {
    if (!phonebookSheet) {
      throw `Invalid phonebook sheet ${phonebookSheet}.`;
    }
    let range = phonebookSheet.getDataRange();

    let values = range.getValues();

    let fontSizes = range.getFontSizes();
    let fontColors = range.getFontColorObjects();
    let backgrounds = range.getBackgrounds();

    for (let i = 0; i < values.length; i++) {
      this.addPhoneNumber(values[i][0], values[i][1]);

      let phoneNameStyle = {
        fontSize: fontSizes[i][1],
        fontColor: fontColors[i][1].asRgbColor().asHexString(),
        backgroundColor: backgrounds[i][1]
      };

      this.styleDict[values[i][0]] = phoneNameStyle;
    }
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

    values.forEach((row) => {
      row[0] = this.lookUpName(row[0]);
      row[1] = this.lookUpName(row[1])
    })

    range.setValues(values)
  }

  /**
   * Method text style (fontSize, fontColor, backgroundColor) of
   * a cell with corresponding style stored in the phone book.
   */
  replaceTextStyles(targetSheet) {
    if (!targetSheet) {
      throw `Invalid target sheet ${targetSheet}.`;
    }

    let range = targetSheet.getDataRange(); 
    let values = range.getValues();

    let fontSizes = range.getFontSizes();
    let fontColors = range.getFontColorObjects();
    let backgrounds = range.getBackgrounds();

    values.forEach( (values_row, i) => {
      let textStyle = this.lookUpStyle(values_row[0]);
      fontSizes[i][0] = textStyle.fontSize;
      fontColors[i][0] = textStyle.fontColor;
      backgrounds[i][0] = textStyle.backgroundColor;

      textStyle = this.lookUpStyle(values_row[1]);
      fontSizes[i][1] = textStyle.fontSize;
      fontColors[i][1] = textStyle.fontColor;
      backgrounds[i][1] = textStyle.backgroundColor;
    });

    range.setFontSizes(fontSizes);
    range.setFontColors(fontColors);
    range.setBackgrounds(backgrounds);
  }
  
  /**
   * Method returns the name associated with a phone number.
   */
  lookUpName(number) {
    if (this.phoneNumberDict[number]) {
      return this.phoneNumberDict[number];
    }
    else {
      return number;
    }
  }

  /**
   * Method returns the text style associated with a phone number.
   */
  lookUpStyle(number) {
    if (this.styleDict[number]) {
      return this.styleDict[number];
    }
    else {
      return {
        fontSize: 10,
        fontColor: '#000000',
        backgroundColor: '#ffffff'
      }
    }
  }

  /**
   * Method serialized Phonebook object to a JSON string.
   */
  toJSON() {
    let data = {
      phoneNumberDict: this.phoneNumberDict,
      styleDict: this.styleDict
    }

    return JSON.stringify(data);
  }

  /**
   * Method deserializes Phonebook object from a JSON string.
   */
  fromJSON(jsonString) {
    this.phoneNumberDict = jsonString.phoneNumberDict
    this.styleDict = jsonString.styleDict
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

  if (selectedPhoneBookValue != null && selectedPhoneBookValue != EMPTY_PHONEBOOK) {
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
  .addItem("Replace Phone Numbers (Apply Text Styles)", "replacePhoneNumbersApplyTextStyles")
  .addToUi();

  setSelectedPhonebook(EMPTY_PHONEBOOK);
}


const EMPTY_PHONEBOOK = -1;


/**
 * Function prompts the user for a name of a sheet in the Google Sheets spreadsheet.
 * The function returns selected sheet if given name exists, null otherwise.
 */
function selectSheetPrompt(ui, message) {
  let targetSheetName = ui.prompt(message).getResponseText();
  let sheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);

  if (sheet == null) {
    ui.alert(`InvalidSheetName: Sheet with name ${targetSheetName} was not found.`);
  }

  return sheet;
}

/**
 * Function prompts the user for a name of a sheet in the Google Sheets spreadsheet.
 * The function returns selected sheet if given name exists, null otherwise.
 */
function selectPhonebook() {
  let ui = SpreadsheetApp.getUi();
  let sheet = selectSheetPrompt(ui, "Enter phonebook sheet name.");

  if (sheet == null) {
    return;
  }

  setSelectedPhonebook(new Phonebook(sheet));
  ui.alert(`Phonebook was created successfully.`);
}

/**
 * Function replaces phone numbers with corresponding names from the phonebook
 * in the target sheet.
 */
function replacePhoneNumbers() {
  let ui = SpreadsheetApp.getUi();
  let sheet = selectSheetPrompt(ui, "Enter target sheet name.");

  if (sheet == null) {
    return;
  }
  
  if (getSelectedPhonebook() == EMPTY_PHONEBOOK){
    ui.alert("EmptyPhonebookError: Phonebook is currently empty, please select one of the sheets as phonebook.");
  }
  else {
    let phonebook = getSelectedPhonebook();
    phonebook.replacePhoneNumbers(sheet);
  }
}

/**
 * Function replaces phone numbers with corresponding names from the phonebook
 * in the target sheet and applies stored fontSize, fontColor, backgroundColor.
 */
function replacePhoneNumbersApplyTextStyles() {
    let ui = SpreadsheetApp.getUi();
  let sheet = selectSheetPrompt(ui, "Enter target sheet name.");
  
  if (sheet == null){
    return;
  }

  if (getSelectedPhonebook() == EMPTY_PHONEBOOK){
    ui.alert("EmptyPhonebookError: Phonebook is currently empty, please select one of the sheets as phonebook.");
  }
  else{
    let phonebook = getSelectedPhonebook();
    phonebook.replaceTextStyles(sheet);
    phonebook.replacePhoneNumbers(sheet);
  }
}