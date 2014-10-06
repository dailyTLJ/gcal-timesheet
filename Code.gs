// 1. change calname
// 2. change yourname
// 3. change year
// 4. compute with RUN > calculate_timesheet
// 5. go back to the spreadsheet and look at the resultss
// 6. if projects haven't been named/grouped properly, edit the mapProject() function


// name of the calendar
// var calname = "DTLJ_hours_eva";
// var calname = "DTLJ_hours_mouna";
var calname = "DTLJ_hours_pierre";

// name of person (it'll be printed in a column, and be included in the sheet-name)
// var yourname = "Eva";
// var yourname = "Mouna";
var yourname = "Pierre";

// specify calendar year you want to compute
var year = 2014;

var company = "DAILY TOUS LES JOURS";
var company_short = "DTLJ";//


/*
 * Main function to compute the timesheet
 *
 */
function calculate_timesheet(){

  // MOST IMPORTANT = GET ACCESS TO CALENDAR

  // alternative 1: 
  // get calendar by name, you gmail account needs to have/give access to the calendar
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
  var sheet = ss.getSheetByName(yourname+"_"+yearStr);
  if (sheet == null) {
    sheet = ss.insertSheet(yourname+"_"+yearStr);
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
  
  var project_head = sheet.getRange(1,10,3,1);
  project_head.setValues([["PROJECTS"],[ ""],[ "INDIVIDUAL PROJECT HOURS"]]);
  project_head.setFontWeight("bold");
  


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
    var letter = numToCol(p);
    range.setFormula('=SUM('+letter+'8:'+letter+'10000)');
  }
  
  // Create the header
  var header = [["Wk","Name","Log description","(Category)", "Date", "Project", "Total Hours", "Weekly Total", ""]]
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
          sumCell.setValue(sumWeek.toFixed(2));
        }
        sumWeek = 0;
        newWeek = weekno;
        writeWeek = weekno;
        off++;
      }
      row=i+off;
      
      // write out event details
      var details=[[writeWeek, yourname,description, "", new Date(startTime.getFullYear(), startTime.getMonth(), startTime.getDate()), project,  duration.toFixed(2)]];
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
          projectCell.setValue(duration.toFixed(2));
        }
      }
    
    }
  }
  
  // write week number
  if (newWeek != "") {
    var sumCell=sheet.getRange(row,8,1,1);
    sumCell.setValue(sumWeek.toFixed(2));
    sumCell.setFontWeight("bold");
  }
}


/*
 * clean-up of project names
 * if you ended up using different naming-conventions
 * for projects, but still want to group the hours together
 * 
 */
function mapProject( p ) {
  // get rid of whitespace characters before and after
  p = p.trim();
  p = p.toUpperCase(); 
  
  switch (p) {
      
    case "UNIOND":  
    case "UNION DEPOT":
      return "UD";
      
    case "IT":
      return "ITM";
      
    case "R+D":
    case "DAILY":
    case "OFFICE":
    case "KICKSTART":
      return company_short;
      
    case "21 O":
      return "21O";
      
    case "TURLUTTE":
      return "TURLUTE";
      
    case "SDCV":
      return "MEMORAMA";
      
    case "PLANETARIUM":
      return "PLANE";
      
    case "1%":
      return "ONE%";
      
    case "NFB":
    case "MCLAREN":
    case "MCLARENA":
      return "MCL";
      
    case "T-A-D":
      return "TAD";
  }
  
  return p;
}


/*
 * define the week number
 * 
 */
function getWeek( d ){ 
  var onejan = new Date(d.getFullYear(),0,1); 
  return Math.ceil((((d - onejan) / 86400000) + onejan.getDay()+1)/7); 
} 


function numToCol( p ){
  if (p<17) {
    // return String.fromCharCode(p+10 + 65-1);
    return String.fromCharCode(p+10+64);
  } else {
    var firstChar = String.fromCharCode( Math.floor((p+10)/26) +64 );
    var secondChar = String.fromCharCode( (p+10)%26 + 64 );
    return firstChar+''+secondChar;
  }
}


/*
 * some info that pops up when opening the spreadsheet 
 * connected to the gscript
 * 
 */
function onOpen() {
  Browser.msgBox('Time to crunch numbers!', '1) Go to the script via TOOLS > Script Editor\\n2) Change the calendar-name, person-name and calendar year\\n3) Then when ready click RUN > calculate_timesheet', Browser.Buttons.OK);
}
