# Training expiry notifier for Google Apps Script

A small utility that reads a sheet named Training, keeps only the newest record per person and course, then emails a weekly summary of items that are expired, expiring within thirty days, and expiring within sixty days. It also color codes each line by department inside the email body.

## What it expects in your sheet

Columns in the Training sheet must be laid out as follows

A Training
B First Name
C Last Name
D Status
E Due Date
F Completion Date
G Expiration Date
H Department

The script uses zero based array indices under the hood and the current field picks are
training name at index one
first name at index two
last name at index three
expiration date at index seven
department at index eight

If your layout differs, adjust the indices in Code.gs.

## Quick start

1 Create a new project in Google Apps Script or use an existing one.
2 Add the files from this repo. Keep the folder called src and the manifest file appsscript.json at the project root.
3 In the Apps Script editor, open File then Project properties then Script properties and add the keys below
   SHEET_ID
   SHEET_NAME
   RECIPIENT_EMAIL
   CC_EMAIL
   SUBJECT_LINE

   You can leave any of them blank to use the defaults inside the code. Defaults are only for development, set real values before you run in production.

4 Run the function sendWeeklyExpiryEmail once to authorize scopes.
5 Run the function createTimeDrivenTrigger to schedule a weekly run. It schedules Monday at seven in the morning. Change the hour in the helper if you prefer a different time.

## The email buckets

Using only the latest record per person and course
expired
expiring within thirty days
expiring within sixty days

## Department colors

Defaults inside the code
Production red
Estimating orange
Human Resources purple
Service green
Sales blue

Edit the map inside formatList if you want different colors.

## Sheet level helper formula

Place this inside any summary table where columns A and B hold first name and last name, and the column header in row two contains the course name before any parenthesis. The formula returns the newest expiration date for that person and course with a status label

```
=IFERROR(
  LET(
    person, $A3 & $B3,
    hdr, INDEX($2:$2, COLUMN()),
    course, TRIM(REGEXEXTRACT(hdr & "", "^[^(\n]+")),
    exp, MAX(
      FILTER(
        Training!$H:$H,
        Training!$A:$A = person,
        LOWER(TRIM(REGEXEXTRACT(Training!$B:$B & "", "^[^(\n]+"))) = LOWER(course),
        Training!$H:$H <> ""
      )
    ),
    label, IF(exp < TODAY(), "Expired",
           IF(exp <= TODAY()+30, "Expiring in 30 days",
           IF(exp <= TODAY()+60, "Expiring in 60 days", ""))),
    dt, TEXT(exp, "yyyy/mm/dd"),
    IF(label = "", dt, dt & " " & CHAR(45) & " " & label)
  ),
  "Not in the system"
)
```

## Scopes used

spreadsheets read only
send mail
create triggers

These are declared in appsscript.json.

## Notes

This project uses V8 runtime.
Dates are formatted with the script time zone.
If you prefer to bind the script to the sheet, you can simplify by reading SpreadsheetApp.getActive instead of openById and you may store configuration in Script properties the same way.
