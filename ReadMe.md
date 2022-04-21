# NoPixel Phonebook Add-on
---

This add-on allows the users to more easily substitute phone numbers with corresponding names in a Google Sheets document. Its main purpouse is to allow users to receive a readable form of a phone subpoena utilizing a pre-existent set of phone numbers and names.

## Installation
---

The add-on is currently under review in the Google Marketplace and it will most likely be under review for the foreseeable future.
Therefore, it is not possible to install it through the marketplace

![Add-on is under review](/resources/installation/under_review.png)

Alternative installation:

1. (Optional) Download a release version of the add-on
2. Open Google Sheets
3. Open Google Apps Scripts (located in the `Extensions` menu on the toolbar)
4. Copy the entire contents of `Phonebook.gs`
5. Replace the `function myFunction()` in `Code.gs` with the contents of `Phonebook.gs`
6. Reload Google Sheets (it might take more than one try)
7. Done

Note: The script will only work for the specific Google Sheet. You will have to repeat this process for different sheets.

![Step 4](/resources/installation/step_4.png)
![Step 5-1](/resources/installation/step_5_1.png)
![Step 5-2](/resources/installation/step_5_2.png)
![Step 6](/resources/installation/step_6.png)


## Usage
---
In order to use this add-on you need to create a separate sheet containing pairs of phone numbers and names that will act as a 'phonebook'. This sheet will be used by the add-on as a reference to know which phone numbers need to be replaced in a subpoena. The pairs of `[Number, Name]` must be located in the first two columns of this sheet without any empty rows.

![Screenshot_1](/resources/screenshots/screenshot_1.png)

To replace phone numbers in a subpoena you need to first read a phonebook by using the `Select Phonebook` button. After you have selected a phonebook you can use the `Replace Phone Numbers` button to select a sheet in which you want to replace phone numbers with corresponding names.

![Screenshot_1](/resources/screenshots/screenshot_2.png)

When selecting a sheets to be used as phonebook, type in the sheet name exactly as it appears in the tabs at the bottom. 

![Screenshot_1](/resources/screenshots/screenshot_3.png)

Similarly, you need to input the sheet name correctly when choosing a target sheet where the phone numbers will be replaced.

![Screenshot_1](/resources/screenshots/screenshot_4.png)

Afterwards, each number that appears in the selected phonebook should be replaced with a corresponding name in the target sheet which containing the information from a subpoena. 

![Screenshot_1](/resources/screenshots/screenshot_5.png)