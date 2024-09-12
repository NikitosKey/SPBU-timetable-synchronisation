/* Global var Section */
const scriptProperties = PropertiesService.getScriptProperties();
var Folder_ID = scriptProperties.getProperty('Folder_ID');
var CALENDAR_ID = scriptProperties.getProperty('CALENDAR_ID');
var TimeTableFile_URL_withoutDate = 'https://timetable.spbu.ru/StudentGroupEvents/ExcelWeek?studentGroupId=394790&weekMonday='
var calendar = CalendarApp.getCalendarById(CALENDAR_ID);



/* Main Section */
function main_script(){
  var startDate = new Date();  // Начальная дата // Можно поставить "2024-09-02"" либо "2025-02-01" для рефреша всего первого, либо второго сема.

  if (1 <= startDate.getMonth() && startDate.getMonth() <= 6)
    endDate = new Date(startDate.getFullYear(), 6, 20);

  else if ((8 <= startDate.getMonth() && startDate.getMonth()<= 11) || startDate.getMonth() == 0)
    endDate = new Date(startDate.getFullYear()+1, 0, 31);

  else {
    Logger.log("Hollyday, ending the program...");
    return;
  } 

  //endDate = new Date("2024-09-15");
  GetDataFromSPBUTimetable(startDate, endDate);
  convert(Folder_ID);
  //clearingPeriod(startDate, endDate);
  UploadDataToCalendar(startDate, endDate);
}


function GetDataFromSPBUTimetable(startDate, endDate) {
  
  // Uploading fresh TimeTable
  let mondays = getMondays(startDate, endDate);
  
  mondays.forEach(function(monday) {
    let year = monday.getFullYear();
    let month = (monday.getMonth() + 1).toString().padStart(2, '0');
    let day = monday.getDate().toString().padStart(2, '0');


    let formattedDate = year + '-' + month + '-' + day;
    writeFile(formattedDate);
  });
}


function UploadDataToCalendar(startDate, endDate) {
  startDate = findLastMonday(startDate);
  let mapOfFiles = getAllFileNames();

  mapOfFiles.forEach(function(Id, FileName) {
    let fileDate = new Date(FileName).getTime();
    if (fileDate >= startDate.getTime() && fileDate <= endDate.getTime()) 
      SetCalendar(Id);
  });
}


/* Getting Data Section */

// Download file from SPBU Timetable and save it in folder.
function writeFile(CurrentWeak){
  let file_url = TimeTableFile_URL_withoutDate + CurrentWeak;
  let response = UrlFetchApp.fetch(file_url);


  let blob = response.getBlob();
  let content = response.getContent();
  let fileBlob = Utilities.newBlob(content, blob.getContentType(), CurrentWeak); 
  let folder = DriveApp.getFolderById(Folder_ID);
  let fileNamesAndIds = getAllFileNames();

  while(fileNamesAndIds.has(CurrentWeak)) {
    DriveApp.getFileById(fileNamesAndIds.get(CurrentWeak)).setTrashed(true);
    fileNamesAndIds = getAllFileNames();
  }
  folder.createFile(fileBlob);
}

// Convert all .xlsx files to spreadsheet in folder
function convert(folder_Id) {
  let folderIncoming = DriveApp.getFolderById(folder_Id);
  let files = folderIncoming.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  while (files.hasNext()) {
    let source = files.next();
    let sourceId = source.getId();
    let fileName = source.getName().replace('.xlsx', '');
    
    let file = {
        title: fileName,
      };
      
    file = Drive.Files.copy(file, sourceId, {convert: true});

    Drive.Files.remove(sourceId);
  }
}


//Gets names and ids of all files in the folder
function getAllFileNames() {
  let folder = DriveApp.getFolderById(Folder_ID); 
  let files = folder.getFiles();
  let result = new Map()


  while (files.hasNext()) {
    let file = files.next();
    result.set(file.getName(), file.getId());
  }

  return result;
}


/* Filling Calendar Section */

// Clear all events between start and end dates
function clearingPeriod(startDate, endDate){
  const events = calendar.getEvents(startDate, endDate);
  events.forEach(function(event){ 
    event.deleteEvent();
  });
}

// Setting calendar
function SetCalendar(SPREADSHEET_ID) {
  let table = SpreadsheetApp.openById(SPREADSHEET_ID);

  tEvents = getEventsFromSpreadsheet(table);

  let tEventsMap = new Map();
  tEvents.forEach(function(event){
    let str_buff = '';
    str_buff += event.title;
    str_buff += getYearMonthDayHHMM(event.startDate);
    str_buff += getYearMonthDayHHMM(event.endDate); 
    str_buff += event.location; 
    str_buff += event.info;
    tEventsMap.set(str_buff, event);
  });


  let startDate = new Date(table.getName());
  let endDate = new Date(startDate);
  endDate.setDate(endDate.getDate() + 7);
  let cEvents = calendar.getEvents(startDate, endDate);
  

  let cEventsMap = new Map();

  cEvents.forEach(function(event){
    let str_buff = '';
    str_buff += event.getTitle();
    str_buff += getYearMonthDayHHMM(event.getStartTime());
    str_buff += getYearMonthDayHHMM(event.getEndTime());
    str_buff += event.getLocation();
    str_buff += event.getDescription();
    cEventsMap.set(str_buff, event);
  });


  cEventsMap.keys().forEach(function(cEvent){
    if(!(tEventsMap.has(cEvent))){
      cEventsMap.get(cEvent).deleteEvent();
    }
  });

  tEventsMap.keys().forEach(function(tEventKey){
    if(!(cEventsMap.has(tEventKey))){  
      let tEventValue = tEventsMap.get(tEventKey); 
      // You can add filters.
      // For example
      //if (tEventValue.title != 'English, practical class' || tEventValue.info == 'Teacher: Kononov B.'){
        calendar.createEvent (tEventValue.title, tEventValue.startDate, tEventValue.endDate, 
          { 
            location: tEventValue.location,
            description: tEventValue.info
          });
      //}
    }
  });
}

// Class for event from Spreadsheet
class tableEvent {
  constructor(title, startDate, endDate, location, info) {
    this.title = title;
    this.startDate = startDate;
    this.endDate = endDate;
    this.location = location;
    this.info = info;
  }
}

// Gets events from table
function getEventsFromSpreadsheet(table) {
  let sheet = table.getActiveSheet();
  let data = sheet.getDataRange().getValues();
  let year = data[1].toString().split(' ')[2];

  const dateCol = 0;
  const timeCol = 1;
  const titleCol = 2;
  const placeCol = 3;
  const teachersCol = 4;
  
  let tEvents = [];

  for (var i = 4; i < data.length; i++) {
    var row = data[i];

    let dateStr = row[dateCol].toString();
    let timeStr = row[timeCol].toString();
    let title = row[titleCol].toString();
    let location = row[placeCol].toString();
    let teacher = row[teachersCol].toString();

    if (dateStr != '') savedDate = dateStr;
        
    let dateParts = savedDate.split(" ");

    let month = getMonthNumber(dateParts[1]);
    let day = parseInt(dateParts[2]);
    
    let timeParts = timeStr.split("–");
    let startTimeParts = timeParts[0].split(":");
    let endTimeParts = timeParts[1].split(":");
    
    let startHours = parseInt(startTimeParts[0]);
    let startMinutes = parseInt(startTimeParts[1]);
    let endHours = parseInt(endTimeParts[0]);
    let endMinutes = parseInt(endTimeParts[1]);
    
    let startDate = new Date(year, month, day, startHours, startMinutes);
    let endDate = new Date(year, month, day, endHours, endMinutes);
    let info = "Teacher: " + teacher;

    tEvents.push(new tableEvent(title, startDate, endDate, location, info));
  }
  return tEvents;
}

/* Date functions */

// Returns Month Number by its string name
function getMonthNumber(monthName) {
  let monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return monthNames.indexOf(monthName);
}

// Find monday on the left of current date
function findLastMonday(date){
  while (date.getDay() !== 1) { 
    date.setDate(date.getDate() - 1);
  }
  return date;
}


// Gets all the beginnings of the weeks for the rest of the semester
function getMondays(startDate, endDate) {
  let result = [];
  let date = new Date(startDate);
  
  date = findLastMonday(date);
  
  while (date <= endDate) {
    result.push(new Date(date));
    date.setDate(date.getDate() + 7);
  }
  
  return result;
}

// Returns date in YYYY-MM-DD_HH:MM format
function getYearMonthDayHHMM(date) {
  const year = date.getFullYear()
  const month = (date.getMonth() + 1 < 10 ? '0' : '') + (date.getMonth() + 1)
  const day = date.getDate();
  const hour = date.getHours().toString();
  const minutes = date.getMinutes().toString();

  return year + '-' + month + '-' + day + '_' + hour + ':' + minutes;
}
