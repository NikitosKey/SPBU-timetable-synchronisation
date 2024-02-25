/*
Начало работы: 
Необходимо поменять переменные в Global var Section.
1) Указать папку для скачивания excel таблиц с timetable
2) Указать id гугл календаря, в который перенесётся расписание. (Лучше всего создать чистый календарь, чтобы расписание можно было скрывать и оно не смешивалось с другими событиями основного календаря)
3) Указать ссылку на скачивание таблицы своей группы, она идентична той, которая в образце, главное вырезать из неё всё лишнее, по идее она получится почти такая же, за исключением studentGroupId.

Краткое описание сего говнокодинга:

1) Сперва смотрим дату для того, чтобы понять в какой части учебного года находится начало выполнения скрипта, маленькая оптимизация, чтобы получать и обновлять расписание с какой-то даты по конец текущего семестра.

2) Скачиваем с сайта таблички

3) Конверитуем их в гугл таблички 

4) Считываем данные, создаём для каждого зантия событие в календаре и заполняем.

Кароч, тут всё самое важное закомментировано, читайте говнокод перед запуском. 
Самое главное, чтобы у вас была папка и календарь и их id были тут, потому что если вы запустите это с моими id то есть вероятность сломать всё мне, хотя восстановить будет нетрудно, да и защита от долбоёба там по идее какая-то есть.

Функция конвертации всегда конвертирует все xlsx в вашей папке в gsheet затем удаляет все xlsx. Будьте внимательны, если нужны xlsx,то закомментируйте там одну строку и xlsx останутся вместе с gsheet.
Функции main_script(), GetDataFromSPBUTimetable(), UploadDataToCalendar(), convert(), writeFile(), SetCalendar() можно менять под свои задачи путём комментирования и декомментирования каких-то строк, обращайте внимания на подписи.
Функции writeFile(), SetCalendar(), convert() можно сильно оптимизировать, если есть желание.

Удачи в разборе этого говнокода :)
 */



// Global var Section
var FolderID = '1gjiGOijTFkun8z3BvFzHQ5RntiJHhcUe';
var CALENDAR_ID = 'c_955fb74aa9f7de31f92723985fd153561d836954314fe0be7fe22cde3abe37d0@group.calendar.google.com'
var TimeTableFile_URL_withoutDate = 'https://timetable.spbu.ru/StudentGroupEvents/ExcelWeek?studentGroupId=366599&weekMonday='



// Main Section
function main_script(){
  // В зависимости от целей можно комментировать каждую строку
  GetDataFromSPBUTimetable();
  convert(FolderID);
  UploadDataToCalendar();
}


function GetDataFromSPBUTimetable() {
  var startDate = new Date();  // Начальная дата // Можно поставить "2023-09-01" либо "2024-01-01" для рефреша всего первого, либо второго сема.
  var endDate;
  //var count = 1;
  if (1 <= startDate.getMonth() && startDate.getMonth() <= 6)
    endDate = new Date (startDate.getFullYear(), 6, 20);

  else if ((8 <= startDate.getMonth() && startDate.getMonth()<= 11) || startDate.getMonth() == 0)
    endDate = new Date (startDate.getFullYear()+1, 0, 31);

  else {
    Logger.log("Hollyday, ending the program...");
    return;
  } 
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


function UploadDataToCalendar() {
  let mapOfFiles = getAllFileNames();
  let todayDate = new Date();
  let todayTime = todayDate.getTime();
  
  
  mapOfFiles.forEach(function(Id, FileName) {
    // Эту часть надо комментировать, если нужно обновить весь семестр.
    let fileNameParts = FileName.split('-');
    let fileDate = new Date(fileNameParts[0], fileNameParts[1]-1, fileNameParts[2]);
    let fileTime = fileDate.getTime();
    if (todayTime <= fileTime) 

    // Эту часть комментировать не надо, иначе в календаре ничего не изменится
    SetCalendar(Id);
  });
}



// Getting Data Section
function writeFile(CurrentWeak){
  let file_url = TimeTableFile_URL_withoutDate + CurrentWeak;
  let response = UrlFetchApp.fetch(file_url);


  let blob = response.getBlob();
  let content = response.getContent();
  let fileBlob = Utilities.newBlob(content, blob.getContentType(), CurrentWeak); 
  let folder = DriveApp.getFolderById(FolderID);
  let fileNamesAndIds = getAllFileNames();

  while(fileNamesAndIds.has(CurrentWeak)) {
    DriveApp.getFileById(fileNamesAndIds.get(CurrentWeak)).setTrashed(true);
    fileNamesAndIds = getAllFileNames();
  }
  folder.createFile(fileBlob);
}


function convert(folderId) {
  let folderIncoming = DriveApp.getFolderById(folderId);
  let files = folderIncoming.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  while (files.hasNext()) {
    let source = files.next();
    let sourceId = source.getId();
    let fileName = source.getName().replace('.xlsx', '');
    
    let file = {
        title: fileName,
      };
      
    file = Drive.Files.copy(file, sourceId, {convert: true});


    // Если вам не нужно удалять xlsx, то закомментируйте.
    Drive.Files.remove(sourceId);
  }
}


// Gets all the beginnings of the weeks for the rest of the semester
function getMondays(startDate, endDate) {
  let result = [];
  let date = new Date(startDate);
  
  // Переход к первому понедельнику после startDate
  while (date.getDay() !== 1) { 
    date.setDate(date.getDate() - 1);
  }
  
  // Перебираем все понедельники до endDate
  while (date <= endDate) {
    result.push(new Date(date));
    date.setDate(date.getDate() + 7);
  }
  
  return result;
}

//Gets names and ids of all files in the folder
function getAllFileNames() {
  let folder = DriveApp.getFolderById(FolderID); 
  let files = folder.getFiles();
  let result = new Map()


  while (files.hasNext()) {
    let file = files.next();
    result.set(file.getName(), file.getId());
  }

  return result;
}



// Filling Calendar Section
function SetCalendar(SPREADSHEET_ID) {
  let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  let data = sheet.getDataRange().getValues();
  let year = data[1].toString().split(' ')[2];
  
  let calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  
  let savedDate = "";
  // let saveTitle = "";
  // let saveLocations = ""; //Тут попытки в оптимизацию
  // let saveInfo = "";
  // Пропускаем первую строку таблицы (заголовки столбцов)
  for (var i = 4; i < data.length; i++) {
    var row = data[i];

    const dateCol = 0;
    const timeCol = 1;
    const titleCol = 2;
    const placeCol = 3;
    const teachersCol = 4;


    let dateStr = row[dateCol].toString();
    let timeStr = row[timeCol].toString();
    let title = row[titleCol].toString();
    let location = row[placeCol].toString();
    let teacher = row[teachersCol].toString();

    // Тут тоже я пытался, но забил болт)
    // if(title != saveTitle){
    //   saveTitle = title;
    //   saveLocations = location;
    //   saveInfo = "Teacher: " + teacher;
    // } 
    // else {
    //     saveLocations += '\n' + location;
    //     saveInfo += ', ' + teacher;
    // }

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

    // Логгер, если нужен)
    // Logger.log("Title " +title.toString());
    // Logger.log("Location " +location.toString());
    // Logger.log("Info " +info.toString());
    // Logger.log("Day " +day.toString());
    // Logger.log("Month " +month.toString());
    // Logger.log("Year " +year.toString());
    // Logger.log("startHours " +startHours.toString());
    // Logger.log("startMinutes " +startMinutes.toString());
    // Logger.log("endHours " +endHours.toString());
    // Logger.log("endMinutes " +endMinutes.toString());
    // Logger.log(startDate);
    // Logger.log(endDate);
    // Logger.log("=====================");


    let events = calendar.getEventsForDay(startDate);


    let flag = false;
    events.forEach(function(event){
      let eventTitle = event.getTitle();
      //let eventLocation = event.getLocation();
      let eventDescription = event.getDescription();
      if (eventTitle.toString() == title.toString() && eventDescription.toString() == info)
        event.deleteEvent();
    });

   if(!flag){
      calendar.createEvent (title, startDate, endDate, 
      { 
        location: location,
        description: info
      });
   }

  }
}

// Функция для определения номера месяца по его названию
function getMonthNumber(monthName) {
  let monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return monthNames.indexOf(monthName);
}
