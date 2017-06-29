# Google Apps Script Work for David Hestrin

### BEIGoogleFormEmail (features/email)

Created:	 Jun 28 2017  
Edited:	 Jun 29 2017

-

### Summary

The code in this repo is for use with Google Apps Script to process Google Sheets for Better Earth Institute. This README file should be used for high-level notes and instructions regarding the files in this repo and how to use this script.

### Branch

#### features/email

This branch contains just the code necessary to send a quick email straight from Google Sheets.

-

### How To Use

1. Select `"Email"` >> `"Email comments..."` from the menu bar in Google Sheets.

2. Enter the column index where email addresses are stored. Column A = 1, column B = 2, etc.

3. Enter the column index where email bodies are stored. Column A = 1, column B = 2, etc.

4. Enter the subject line for your emails. All emails will use the same subject line.

5. A confirmation alert will appear once all emails have been sent.
