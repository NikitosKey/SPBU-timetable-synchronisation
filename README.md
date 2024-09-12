SPBU-timetable-synchronisation
---
Synchronization of SPBU Timetable with Google calendar via Google Drive and Google Apps Scripts.


# Getting started

- Create a folder on your GDrive for saving and converting .xlsx from [site](https://timetable.spbu.ru/MATH/).
- Create a Google Apps Scripts project. 
- Create a Google calendar.
- Copy code from script.gs to your project.
- Go to project settings and turn on appsscript.json.
- Copy all from appsscript.json from this repo to your json in your project.
- Add Folder_ID and CALENDAR_ID in the script properties.
- Change variable TimeTableFile_URL_withoutDate in the script to the link to download the schedule file for your group. This is the same as in the example, except for the studentGroupId=XXXXXX.
- Add script triggers to run script automaticly.

 Enjoy :)
