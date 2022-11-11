# Google Calendar Time Tracker

This is a simple Google App Script attached to a Google sheet.

## Functionality

This script will run through you calendar invites and pull all invites with:
- either an external domain to your company domain
- your email alias followed by + and any categorisation text you'd like to have
  - e.g john.doe+marketing@mydomain.com

## Usage
In the "Info" sheet:
- Insert a start and end date for the script to extra events in that date range
- Insert your email alias
  - e.g john.doe
  The script will then look for anything with john.doe+{categorisation text here}@mydomain.com

- Once you've added the required info, you can run the script using the Custom Menu in Google Sheet:
  - Time Tracker -> Extract from Calendar
- The events will be extracted into the "CalendarEvents" sheet. If you want to tweak the content, duplicate it into a separate sheet.

- The script can be run as many times as desired, it will always clear out "CalendarEvents" and re-extract the events into that sheet.
