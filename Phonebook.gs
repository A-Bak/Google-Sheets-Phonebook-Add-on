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
      values[i][0] = this.replace(values[i][0])
      values[i][1] = this.replace(values[i][1])
    }

    range.setValues(values)
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
    return JSON.stringify(this.phoneNumberDict)
  }

  /**
   * Method deserializes Phonebook object from a JSON string.
   */
  fromJSON(jsonString) {
    this.phoneNumberDict = jsonString
  }
}

/**
 * Function sets the script property 'selectedPhoneBook' to provided
 * phonebook. Used to store the phone book between script calls.
 */
function setSelectedPhonebook(phonebook) {
  let scriptProperties = PropertiesService.getScriptProperties();
  let jsonString = phonebook.toJSON();
  Logger.log(jsonString);
  scriptProperties.setProperty('selectedPhoneBook', jsonString);
}

/**
 * Function returns the selected phonebook if one was created.
 * Otherwise it returns EMPTY_PHONEBOOK value.
 */
function getSelectedPhonebook() {
  let scriptProperties = PropertiesService.getScriptProperties();
  let jsonString = scriptProperties.getProperty('selectedPhoneBook');
  return JSON.parse(jsonString)
}
