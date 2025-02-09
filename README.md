# Batch Emailer

## Rationale
Out of necessity, I created an Apps Script for a Google Sheet so that I could pick and choose which parents to send a batch email to. You can use this for emailing parents or emailing anyone else for that matter. It's handy to have when you want to send specific messages to groups of people for sports teams, certain groups, family, or anyone else. Of course you can create message lists in your email program, but this is meant to be a temporary list of people and their emails which you can just select and send messages to. Only works with Google Sheets and Gmail.

## Usage
Copy the Google Sheet available here: [Batch Emailer Google Sheet](https://docs.google.com/spreadsheets/d/1_GBEL4F2JmbVhUSW21HPu8fMpxNp51gkV8lLS2ZgyGQ/copy)
![Animation on how to use the Google Sheet](https://secretgoldfish.weebly.com/uploads/3/6/5/7/365767/batch-emailer-v1-3-birdy-demo_orig.gif)

## Changelog
### Batch Emailer v1.3.2 (birdy) [Nov. 10, 2024]
* Re-wrote documentation on top of Apps Script code to detail the Google Sheets requirements in terms of columns and rows
* Added .trim() function to name and email fields to trim extra spaces possibly added to cells

### Batch Emailer v1.3.1 (birdy) [Nov. 9, 2024]
* Bug fix - bcc emails count not considering empty addresses and duplicates
* Emails will now be filtered so that duplicates and empty email addresses will be removed before being counted and added to the Bcc field in Gmail

### Batch Emailer v1.3 (birdy) [Oct. 27, 2024]
* Google limits each message to 50 emails currently that can be used by Apps Script
* Added a feature that will chunk the email addresses if over 50 and divide them over multiple drafts (the same email draft will be created with batches of up to 50 email addresses each)
* Added a special message when the limit is exceeded about the Google limitations with a batch number eg. "Draft 1 of 5"
* Added up to 800 checkboxes on the Google Sheet

### Batch Emailer v1.2 (allin) [Oct. 21, 2024]
* Added an option to "Select All Rows"
* Added an option to "Deselect All Rows"
* Added a email verification which will highlight cells (in yellow) that are missing emails
* Cleaned up the default message in the Draft email

### Batch Emailer v1.1 (saranwrap) [Oct. 15, 2024]
* Added a 2nd email column for a secondary email
* Alternating row colour for visibility
* First column and first row frozen so Sheet can be scrolled without losing track of headings
* Colours added to tabs and header row

### Batch Emailer v1.0 (kungfu) [Oct. 14, 2024]
* Google Sheet created and formatted
* Menu created with "Email Tools"
