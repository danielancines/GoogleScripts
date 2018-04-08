function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [];
  menuItems.push({name: 'Sum Hours', functionName: 'googleCalendarEventsHoursCalculator'});

  spreadsheet.addMenu('Custom Scripts', menuItems);
}

function getCalendars(){
  var calendars = [];
  var calendarNames = ['CALENDAR1','CALENDAR2'];

  for each(var calendarName in calendarNames){
    var calendar = CalendarApp.getCalendarsByName(calendarName)[0];
    if (typeof(calendar) === 'undefined')
      continue;

    calendars.push(calendar)
  }

  return calendars;
}

function googleCalendarEventsHoursCalculator() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cal = CalendarApp.getDefaultCalendar();
  var column = 1;
  var calendars = getCalendars();
  var weekTotalHours = 0;
  var today = new Date();
  var totalHoursPerDay = [];
  var myEmail = 'MYEMAIL';

  sheet.clear();

  for each(var calendar in calendars){
    if (typeof(calendar) === 'undefined')
      continue;

    var targetRow = 1;
    var dates = getDaysOfWeek(today);
    var totalHours = 0;
    var totalHoursDay = 0;

    sheet.setFrozenColumns(1);

    sheet.getRange(1, 1).setHorizontalAlignment('Center').setValue('Dates');
    for each(var date in dates){
      sheet.getRange(++targetRow, 1).setValue(date);
    }

    targetRow = 1;

    sheet.getRange(targetRow, ++column).setFontColor(calendar.getColor()).setHorizontalAlignment('Center').setValue(calendar.getTitle());

    for(var i = 0; i < dates.length; i++){
      totalHours = 0;
      targetRow++;
      var events = calendar.getEventsForDay(dates[i]);
      for (var j = 0; j < events.length; j++) {
        var dayHours = getTotalHours(events[j].getEndTime(), events[j].getStartTime());
        if (dayHours > 4)
          continue;

        var guest = events[j].getGuestByEmail(myEmail);
        if (guest == null || guest.getGuestStatus() === CalendarApp.GuestStatus.YES){
          totalHours += dayHours;
        }
      }

      totalHoursDay += totalHours;
      sheet.getRange(targetRow, column).setFontColor(calendar.getColor()).setValue(totalHours);
      totalHoursPerDay.push({row:targetRow,totalHours:totalHours});
    }

    sheet.getRange(targetRow + 2, column).setFontColor(calendar.getColor()).setValue(totalHoursDay);

    sheet.setColumnWidth(column + 1, 30);
    column += 1;
  }

  sheet.getRange(targetRow + 2, 1).setValue('Total');

  sheet.getRange(1, ++column).setValue('Total Hours');

  var dayTotalHours = 0;
  for(var i = 2;i<9;i++){
    dayTotalHours = 0;
    for each(var item in getHoursByRow(totalHoursPerDay, i)){
      dayTotalHours += item.totalHours;
    }

    weekTotalHours += dayTotalHours;
    sheet.getRange(i, column).setValue(dayTotalHours);
  }

  sheet.setColumnWidth(column + 1, 30);
  sheet.getRange(1, column+2).setValue("Weekly Hours");
  sheet.getRange(2, column+2).setValue(weekTotalHours);
  sheet.autoResizeColumns(1, calendars.length * 3);
}

function getHoursByRow(arr, row){
  return arr.filter(function(item){return item.row === row;} );
}

function getTotalHours(endTime, startTime){
  return (endTime/1000/60/60) - (startTime/1000/60/60);
}

function getDaysOfWeek(date){
  var dayNum = date.getDay() - 1;
  if (dayNum < 0)
    dayNum = 6;
  var monday = new Date(date);
  monday.setDate(date.getDate() - dayNum);

  var dates = [];
  dates.push(monday);
  var lastDate = monday.getDate();

  var newDate = new Date(monday);
  for(var i = 1;i<7;i++){
    newDate.setDate(lastDate + 1);
    lastDate = newDate.getDate();
    dates.push(new Date(newDate));
  }

  return dates;
}
