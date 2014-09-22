// * don't include all-day events
// * recognize overlapping events?
// * find tagging for R+D hours


// name of the calendar
var calname = "DTLJ_hours_eva";

// name of person (it'll be printed in a column)
var yourname = "Eva";

// specify calendar year you want to compute
var year = 2014;

var company = "DAILY TOUS LES JOURS";
var company_short = "DTLJ";


/*
 * Main function to compute the timesheet
 *
 */
function calculate_timesheet(){

  // MOST IMPORTANT = GET ACCESS TO CALENDAR

  // alternative 1: 
  // get calendar by name
  var calendars = CalendarApp.getCalendarsByName(calname);
  Logger.log('Found %s matching calendars.', calendars.length);
  var cal = calendars[0];

  // alternative 2:
  // get main calendar of gmail address
  // var mycal = "dailytouslesjours@gmail.com";
  // var cal = CalendarApp.getCalendarById(mycal);

  // alternative 3:
  // get calendar by ID
  // var cal = CalendarApp.getCalendarById('oe85pujkdk0o81ofo8kg98dbq8%40group.calendar.google.com');
  // Logger.log('The calendar is named "%s".', cal.getName());

  
  // get all the events within the specified calendar year
  var yearStr = year.toString();
  var events = cal.getEvents(new Date("January 1, "+yearStr+" 00:00:00 CST"), new Date("December 31, "+yearStr+" 23:59:59 CST"));
  
  // the timesheet will be inserted into the active spreadsheet
  // it will create (or overwrite) a sheet with the name of the calendar year
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(yearStr);
  if (sheet == null) {
    sheet = ss.insertSheet(yearStr);
  } else {
    sheet.clearContents(); 
  }
  

  // header text for the timesheet
  var head = sheet.getRange(1,1,3,1);
  head.setValues([[company],[ "TIMESHEET FOR : "+yourname],[ "YEAR : "+year]]);
  head.setFontWeight("bold");
  
  // summing up the total hours
  var range = sheet.getRange(4,7,1,1);
  range.setFormula('=SUM(G8:G1000)');
  


  // array to collect all project names
  var allProjects = [];
  
  // let's do a first loop over all events, to extract the project names
  // reminder, project names are placed between [] square brackets!
  for (var i=0;i<events.length;i++) {

    var eventText = events[i].getTitle();
    var splitText = eventText.split("]"); 
    var project = "-";
    
    if (splitText.length > 1) {
      project = mapProject(splitText[0].substr(1)); 
    } else {
      project = company_short;
    }
        
    if (allProjects.indexOf(project) == -1) {
      allProjects.push(project);
    }
  }
  
  // summing up of each project's hours
  for (p=0;p<allProjects.length;p++) {
    var range = sheet.getRange(4,10+p,1,1);
    var letter = String.fromCharCode(p+10 + 65-1);
    //range.setFormula('=SUM(G8:G1000)');
    range.setFormula('=SUM('+letter+'8:'+letter+'1000)');
  }
  
  
  // Create the header
  var header = [["Wk","Name","Log description","General Task", "Date", "Project", "Total Hours", "Weekly Total", ""]]
  var range = sheet.getRange(5,1,1,9);
  range.setValues(header);
  range.setFontWeight("bold");
  
  // same variables needed for computing hours
  var newWeek = "";
  var off = 7;
  var sumWeek = 0;
  var row = 0;
  
  // fill in the project names
  range = sheet.getRange(5,10,1,allProjects.length);
  range.setValues([allProjects]);
  range.setFontWeight("bold");
  
  
  // loop through all calendar events again
  // to print them into the timesheet
  // and to calculate hours
  for (var i=0;i<events.length;i++) {
    
    // exclude all-day events, we only want clearly specified time-periods
    if (!events[i].isAllDayEvent()) {
    
      // time
      var startTime = events[i].getStartTime();
      var endTime = events[i].getEndTime();
      var duration = (endTime - startTime) / (3600*1000);
      var weekno = getWeek(startTime);
      sumWeek += duration;

      // text
      var eventText = events[i].getTitle();
      var splitText = eventText.split("]"); 
      var project = "";
      var description = "";
      if (splitText.length > 1) {
        project = mapProject(splitText[0].substr(1)); 
        description = splitText[1].substr(1);
      } else {
        // no project name defined
        project = "-";
        description = eventText;
      }
      
      // week number, and week hours sum
      row=i+off;
      var writeWeek = "";
      if (weekno != newWeek) {
        // write old week sum
        if (newWeek != "") {
          var sumCell=sheet.getRange(row-1,8,1,1);
          sumCell.setValue(sumWeek);
        }
        sumWeek = 0;
        newWeek = weekno;
        writeWeek = weekno;
        off++;
      }
      row=i+off;
      
      // write out event details
      var details=[[writeWeek, yourname,description, "", new Date(startTime.getFullYear(), startTime.getMonth(), startTime.getDate()), project,  duration]];
      var range=sheet.getRange(row,1,1,7);
      range.setValues(details);
      var fontStyles = [[ "bold","normal","normal","normal","normal","normal","bold" ]];
      var formats = [[ "w0","","","","DDD, MMM d, yyyy","","0.00" ]];
      range.setNumberFormats(formats);
      range.setFontWeights(fontStyles);
      
      // write hours into correct project column
      for (p=0;p<allProjects.length;p++) {
        if (project == allProjects[p]) {
          var projectCell = sheet.getRange(row, 10+p, 1,1);
          projectCell.setValue(duration);
        }
      }
    
    }
  }
  
  // write week number
  if (newWeek != "") {
    var sumCell=sheet.getRange(row,8,1,1);
    sumCell.setValue(sumWeek);
    sumCell.setFontWeight("bold");
  }
}

/*
 * To cleanup / refine project names, if needed
 *
 */
function mapProject( p ) {

  p = p.trim();
  p = p.toUpperCase(); 

  if (p == "UNIOND") {
    return "UD";
  }
  if (p == "UNION DEPOT") {
    return "UD";
  }
  else if (p == "R+D") {
    return "DTLJ";
  }
  else if (p == "21 O") {
    return "21O";
  }
  else if (p == "TURLUTTE") {
    return "TURLUTE";
  }
  else if (p == "SDCV") {
    return "MEMORAMA";
  }
  else if (p == "PLANETARIUM") {
    return "PLANE";
  }
  else if (p == "DAILY") {
    return "DTLJ";
  }
  else if (p == "OFFICE") {
    return "DTLJ";
  }
  else if (p == "KICKSTART") {
    return "DTLJ";
  }
  else if (p == "1%") {
    return "ONE%";
  }
  else if (p == "NFB") {
    return "MCL";
  }
  else if (p == "MCLAREN") {
    return "MCL";
  }
  else if (p == "MCLARENA") {
    return "MCL";
  }
  else if (p == "T-A-D") {
    return "TAD";
  }
  else return p;
}


/*
 * To cleanup / refine project names, if needed
 *
 */
function getWeek( d ){ 
  var onejan = new Date(d.getFullYear(),0,1); 
  return Math.ceil((((d - onejan) / 86400000) + onejan.getDay()+1)/7); 
} 


//function onOpen() {
//  Browser.msgBox('App Instructions - Please Read This Message', '1) Click Tools then Script Editor\\n2) Read/update the code with your desired values.\\n3) Then when ready click Run export_gcal_to_gsheet from the script editor.', Browser.Buttons.OK);
//
//}
