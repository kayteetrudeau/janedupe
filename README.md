# janedupe
Automatically copies subscribed Jane calendar to Google calendar so time gets blocked.

The use case is this: I have a google calendar that syncs with jane, so anytime I put an event in that (appointment through private practice, for example) it blocks that time out in Jane. However, the inverse doesn't happen, even if you're subscribed to the Jane calendar - it just adds another calendar view to your current Google one. So you need to add things to that main Google calendar manually when you add a Jane appointment to properly block off the time. 

**The fix:**

1. Go to [https://script.google.com/](https://script.google.com/) from the Google account with the calendar you're using.

Click "New Project" and paste this in the window:

```
function syncIcsToGoogleCalendar() {
  // 1. CONFIGURATION: Replace these with your actual details
  const ICS_URL = "YOUR ICS URL"; 
  const TARGET_CALENDAR_ID = "primary"; // Use "primary" or a specific Calendar ID (e.g., 'abc123@group.calendar.google.com')
  const DAYS_TO_SYNC = 30; // How many days ahead to pull events
  
  const targetCal = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);
  if (!targetCal) {
    Logger.log("Target calendar not found. Check your CALENDAR_ID.");
    return;
  }

  // 2. Fetch the raw ICS file directly
  let response;
  try {
    response = UrlFetchApp.fetch(ICS_URL);
  } catch (e) {
    Logger.log("Failed to fetch ICS feed: " + e.message);
    return;
  }

  const icsData = response.getContentText();
  
  // 3. Simple ICS Parser for VEVENT blocks
  const events = parseIcsEvents(icsData);
  
  const now = new Date();
  const futureCutoff = new Date(now.getTime() + (DAYS_TO_SYNC * 24 * 60 * 60 * 1000));
  
  // Get existing target events to prevent creating duplicates
  const existingEvents = targetCal.getEvents(now, futureCutoff);

  let createdCount = 0;

  events.forEach(event => {
    // Filter to only sync future events within our cutoff window
    if (event.startDate >= now && event.startDate <= futureCutoff) {
      
      // Check if an event with the same title and start time already exists
      const isDuplicate = existingEvents.some(existing => {
        return existing.getTitle() === event.summary && 
               existing.getStartTime().getTime() === event.startDate.getTime();
      });

      if (!isDuplicate) {
        if (event.isAllDay) {
          targetCal.createAllDayEvent(event.summary, event.startDate, {
            description: event.description,
            location: event.location
          });
        } else {
          targetCal.createEvent(event.summary, event.startDate, event.endDate, {
            description: event.description,
            location: event.location
          });
        }
        createdCount++;
      }
    }
  });

  Logger.log(`Sync complete. Created ${createdCount} new events.`);
}

// Helper function to extract event fields from raw ICS text
function parseIcsEvents(icsText) {
  const events = [];
  // Unfold lines split by line breaks and spaces (ICS standard folding)
  const unfolded = icsText.replace(/\r\n[ \t]/g, '');
  const vevents = unfolded.split("BEGIN:VEVENT");

  for (let i = 1; i < vevents.length; i++) {
    const block = vevents[i].split("END:VEVENT")[0];
    
    const summary = getIcsValue(block, "SUMMARY");
    const description = getIcsValue(block, "DESCRIPTION");
    const location = getIcsValue(block, "LOCATION");
    
    const dtStartRaw = getIcsRawLine(block, "DTSTART");
    const dtEndRaw = getIcsRawLine(block, "DTEND");

    if (!dtStartRaw || !summary) continue;

    const startParsed = parseIcsDate(dtStartRaw);
    const endParsed = dtEndRaw ? parseIcsDate(dtEndRaw) : null;

    if (startParsed) {
      events.push({
        summary: summary,
        description: description,
        location: location,
        startDate: startParsed.date,
        endDate: endParsed ? endParsed.date : startParsed.date,
        isAllDay: startParsed.isAllDay
      });
    }
  }
  return events;
}

function getIcsValue(block, key) {
  const match = block.match(new RegExp(`^${key}(?:;[^:]*)?:(.*)$`, "m"));
  return match ? match[1].trim().replace(/\\,/g, ',').replace(/\\;/g, ';').replace(/\\n/g, '\n') : "";
}

function getIcsRawLine(block, key) {
  const match = block.match(new RegExp(`^${key}(.*)$`, "m"));
  return match ? match[0] : null;
}

function parseIcsDate(line) {
  const parts = line.split(":");
  if (parts.length < 2) return null;
  
  const meta = parts[0];
  const value = parts[1].trim();
  const isAllDay = meta.includes("VALUE=DATE") || value.length === 8;

  if (isAllDay) {
    const year = parseInt(value.substring(0, 4), 10);
    const month = parseInt(value.substring(4, 6), 10) - 1;
    const day = parseInt(value.substring(6, 8), 10);
    return { date: new Date(year, month, day), isAllDay: true };
  } else {
    // Basic ISO 8601 UTC parse (YYYYMMDDTHHMMSSZ)
    const year = parseInt(value.substring(0, 4), 10);
    const month = parseInt(value.substring(4, 6), 10) - 1;
    const day = parseInt(value.substring(6, 8), 10);
    const hour = parseInt(value.substring(9, 11), 10) || 0;
    const min = parseInt(value.substring(11, 13), 10) || 0;
    const sec = parseInt(value.substring(13, 15), 10) || 0;
    
    // Assume UTC if ends in Z
    const date = value.endsWith("Z") 
      ? new Date(Date.UTC(year, month, day, hour, min, sec))
      : new Date(year, month, day, hour, min, sec);

    return { date: date, isAllDay: false };
  }
}
```

2. On the `const ICS_URL = "YOUR ICS URL"; ` line, this ICS URL refers to URL you'll see when you click "Settings" from the three dots on the subscribed calendar:
<img width="393" height="353" alt="Screenshot 2026-09-10 at 12 10 36" src="https://github.com/user-attachments/assets/2fb58867-6b43-4b94-88ef-b626019c0bcf" />

<img width="701" height="276" alt="Screenshot 2026-09-10 at 12 11 18" src="https://github.com/user-attachments/assets/045ab34e-d39e-4748-978c-42d76ebacdf2" />

Replace the text in the quotation marks with your actual URL.

If you have multiple Jane accounts, you'll need to make a new Google Scripts project for each one you're subscribed to and it will work (it's been tested)

3. Hover over the alarm clock icon on the left sidebar, then click "Triggers" on the popup menu. It will ask you to save your script: please do so now.

   Configure it as you see fit - I have mine set to every 15 mins.

5. Test by hovering over the `< >` icon in the sidebar and going to "Editor" and clicking the "Run" button on the top bar. If it works, it will tell you how many instances it added to your calendar.


Note: It will not remove cancelled events, so you'll still need to manually remove those from your calendar.
