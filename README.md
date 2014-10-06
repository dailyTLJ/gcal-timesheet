How to create timesheets from your google-calendar events
====
After battling the painful timesheet hours-entry in basecamp for too long, we decided to jump ship and use a combination of google calendar and google script instead. The google app script we now use, takes the events in your google calendar, parses their event name to extract the project name, lists all events in a spreadsheet and sums up your workhours per week and per project. 



How to enter events into your calendar
----
Enter all your work-blocks as separate events into google calendar. Don't have events overlapping. Start each event name with the project name between square brackets, followed by a more detailed description:  "[abc] brainstorming ... "

<p align="center">
	<img src="https://raw.githubusercontent.com/dailyTLJ/gcal-timesheet/master/calendar2.png"/>
</p>



How to compile your timesheet
----
1. Create a new spreadsheet in google drive
2. Open the script editor via > Tools > Script Editor
3. Paste in the script (Code.gs)
4. Adjust the necessary variables
	- calname
	- yourname
	- year
	- company, company_short
5. Compute your hours and generate your timesheet with > Run > calculate_timesheet
6. Go back to the spreadsheet, select the correct sheet and verify the results

<p align="center">
	<img src="https://raw.githubusercontent.com/dailyTLJ/gcal-timesheet/master/spreadsheet.png"/>
</p>



Helpful tips and tricks
----
* Don't make any changes in the spreadsheet that would be deleted when you rerun the script. Rather clean up your calendar events!
* Use your calendar for future-event planning as well (instead of doing that in another calendar), that way you save some time
* The script can't change the width of your columns, but the good thing is, once you layout your columns and rows to your liking, it will keep that layout even when you rerun the script
* Always look at your weekly hours sumup first, to see if the numbers make sense. If these numbers are too high, check if you don't have overlapping event blocks



This script is based on one by Justin Gale, see http://blog.cloudbakers.com/blog/export-google-calendar-entries-to-a-google-spreadsheet. 