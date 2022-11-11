function getTrackedEvents() {

  var lastEventsTrackedCount = 0;

  //Get data from the info sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Info");
  var firstDay = sheet.getRange(1,2).getValue();
  var lastDay = sheet.getRange(2,2).getValue();

  //Get the email alias if defined
  var emailAlias = sheet.getRange(4,2).getValue();

  if (!emailAlias) {
    emailAlias = "nabeel.saad+";
  }

  var events = CalendarApp.getDefaultCalendar().getEvents(firstDay, lastDay, );
  var eventsTrackedCount = 0;
  var eventsTracked = [];

  var rangeToClear = ss.getSheetByName("CalendarEvents").getRange("A2:F1000");
  rangeToClear.clearContent();
  
  Logger.log("First day : " + firstDay + "\nLast day : " + lastDay + "\nProcessing " + events.length + " calendar events.");
  
  for (var i = 0; i < events.length; i++) {

    var guestEmail = "";
    if (!events[i].isAllDayEvent()) {
      var guestList=events[i].getGuestList();
      var trackingGuestEmails = "";
      var manualLabel = "";

      //TESTING LINE
      /*
      if (events[i].getTitle() === "Ankerstore PM feedback/roadmap session") {
        Logger.log("test");
      }
      */

      for (var j = 0; j < guestList.length; j++) {

        var guestEmail = guestList[j].getEmail();

        //Checking for manually added labels using email alias
        if (guestEmail.includes(emailAlias)) {
          //Logger.log("found added alias");
          guestEmail = (guestEmail.match(/(?<=\+).*?(?=@)/));
          if (guestEmail.length > 0) {
            guestEmail = guestEmail[0];
          }
          if (manualLabel == "") {
            manualLabel += guestEmail;
          } else if (manualLabel.includes(guestEmail)) {
              //Logger.log("already found this domain, so moving on");
            } else {
            manualLabel += ", " + guestEmail;
          }
        } else {

          //Checking for customer email domains
          if (!guestEmail.includes("@grafana.com")) {
            guestEmail = guestEmail.match(/(?<=@).*/);
            if (guestEmail.length > 0) {
              guestEmail = guestEmail[0];
            }
            if (trackingGuestEmails == "") {
              trackingGuestEmails += guestEmail;
            } else if (trackingGuestEmails.includes(guestEmail)) {
              //Logger.log("already found this domain, so moving on");
            } else {
              trackingGuestEmails += ", " + guestEmail;
            }
          }
        }
      }

      var eventDuration = getEventDuration(events[i]);
      if (trackingGuestEmails != "") {
        //Logger.log("datetime " + events[i].getStartTime() + " -- event name: " + events[i].getTitle() + " -- attendees: " + trackingGuestEmails + " -- duration " + eventDuration.hoursAndMinutesHR + " -- duration " + eventDuration.timeInMinutes);
        eventsTracked.push([events[i].getStartTime(), events[i].getTitle(), trackingGuestEmails, manualLabel, eventDuration.hoursAndMinutesHR, eventDuration.timeInMinutes ]);
        eventsTrackedCount++;
      }

      if (eventsTrackedCount % 25 == 0 && lastEventsTrackedCount != eventsTrackedCount) {
        Logger.log(eventsTrackedCount + " out of " + i);
        lastEventsTrackedCount = eventsTrackedCount;
      }
    }
    
  }
  //Logger.log("eventsTracked.length " + eventsTracked.length);

  var sheet = ss.getSheetByName("CalendarEvents");
  for (var x = 0; x < eventsTracked.length; x++) {
    for (var y = 0; y < eventsTracked[x].length; y++) {
      sheet.getRange(x+2,y+1).setValue(eventsTracked[x][y]);
    }
    //SpreadsheetApp.flush();
  }

  Logger.log(eventsTrackedCount + " tracked event(s) out of " + events.length);
}


function getEventDuration(event) {
  //Logger.log("end: " + event.getEndTime());
  //Logger.log("start: "+ event.getStartTime());
  var timeInMinutes = (event.getEndTime() - event.getStartTime()) / 1000 / 60; //1000 for milliseconds and 60 for seconds --> get the result in minutes

  var obj = {
    "hoursAndMinutesHR": toHoursAndMinutes(timeInMinutes), 
    "timeInMinutes": timeInMinutes
  };

  return obj;

}

//Convert totalMinutes to xhyym format
function toHoursAndMinutes(totalMinutes) {
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;

  var result = "";
  //return { hours, minutes };
  if (hours != 0) {
    result += hours + "h";
  }

  if (minutes < 10 && minutes > 0) {
    result += minutes + "m";
  } else if (minutes != 0) {
    result += "0" + minutes + "m";
  }

  return result;
}

//Menu item to run from sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Time tracker')
      .addItem('Extract from calendar', 'getTrackedEvents')
      //.addSeparator()
      //.addSubMenu(ui.createMenu('Sub-menu')
      //    .addItem('Second item', 'menuItem2'))
      .addToUi();
}