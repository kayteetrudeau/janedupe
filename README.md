# janedupe
Automatically copies subscribed Jane calendar to Google calendar so time gets blocked.

The use case is this: I have a google calendar that syncs with jane, so anytime I put an event in that (appointment through private practice, for example) it blocks that time out in Jane. However, the inverse doesn't happen, even if you're subscribed to the Jane calendar - it just adds another calendar view to your current Google one. So you need to add things to that main Google calendar manually when you add a Jane appointment to properly block off the time. 

Thanks to all my pals who've helped me test and try other things to get us to a place where multiple clinics and cancellations work (thanks Brayden, Beth, & Jo!)

You will sometimes get errors emailed to you - I do every couple of days but mine're always time-based; when I go to check them, they have eventually worked out. 

Don't forget that there are a few steps below the code.

**The fix:**

1. Go to [https://script.google.com/](https://script.google.com/) from the Google account with the calendar you're using.

Click "New Project" and paste this in the window:

```
/*
 * MULTI-JANE TO GOOGLE CALENDAR SYNC
 *
 * Remember to sub out your ICS URL (the calendar URL you're subscribed to) for the generic line under your `const JANE_ACCOUNTS`,
and edit the name of the worksite under the "name" and "MARKER_PREFIX"
 * Remember to add your gmail email address linked to your calendar under this: `const TARGET_CALENDAR_ID`
 *
 * Imports appointments from multiple Jane calendars into a single Google Calendar.
 *
 * Instructions are here: https://github.com/kayteetrudeau/janedupe/edit/main/README.md
// ==========================================

const TARGET_CALENDAR_ID = 'GOOGLE CALENDAR ACCOUNT'S EMAIL ADDRESS'; // Where all appointments should land

const JANE_ACCOUNTS = [
  {
    name: 'Clinic 1',
    url: 'ICS URL',
    markerPrefix: 'WORKSITE NAME: ',
    key: 'clinic_one'
  }
];


// ==========================================
// MAIN SYNC FUNCTION
// ==========================================

function syncJaneAppointments() {

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    Logger.log('Another sync is already running. Skipping.');
    return;
  }

  try {

    const now = new Date();
    const properties = PropertiesService.getScriptProperties();

    // Validate and load target calendar
    const calendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);
    if (!calendar) {
      Logger.log("Target calendar not found. Check your TARGET_CALENDAR_ID.");
      return;
    }

    // Read destination calendar once for all accounts up to 2035
    const calendarEvents = calendar.getEvents(now, new Date(2035, 0, 1));

    // Index existing imported events by their unique markers
    const eventsByMarker = new Map();

    calendarEvents.forEach(event => {
      const description = event.getDescription() || '';
      
      JANE_ACCOUNTS.forEach(account => {
        const position = description.indexOf(account.markerPrefix);
        if (position === -1) return;

        const uid = description.substring(position + account.markerPrefix.length);
        if (!uid) return;

        const marker = account.markerPrefix + uid;
        if (!eventsByMarker.has(marker)) {
          eventsByMarker.set(marker, []);
        }
        eventsByMarker.get(marker).push(event);
      });
    });

    // Loop through every account configured in JANE_ACCOUNTS
    JANE_ACCOUNTS.forEach(account => {
      
      Logger.log('------------------------------');
      Logger.log('Processing: ' + account.name);

      if (!account.url || (!account.url.startsWith('https://') && !account.url.startsWith('http://')) || account.url.includes('PASTE_')) {
        Logger.log('Skipping ' + account.name + ': URL is missing or invalid.');
        return;
      }

      let response;
      try {
        response = UrlFetchApp.fetch(account.url);
      } catch (e) {
        Logger.log('Failed to fetch feed for ' + account.name + ': ' + e.message);
        return;
      }

      if (response.getResponseCode() !== 200) {
        Logger.log('Failed to fetch ' + account.name + ', status code: ' + response.getResponseCode());
        return;
      }

      const ics = response.getContentText();
      if (!ics.includes('BEGIN:VCALENDAR') || !ics.includes('END:VCALENDAR')) {
        Logger.log('Invalid calendar format for ' + account.name);
        return;
      }

      const unfolded = ics.replace(/\r?\n[ \t]/g, '');
      const rawEvents = unfolded.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g) || [];

      Logger.log('Found ' + rawEvents.length + ' appointments in ' + account.name + ' feed.');

      let created = 0;
      let updated = 0;
      let skipped = 0;
      let deleted = 0;

      const presentUids = new Set();
      const missingKey = uid => 'missing_' + account.key + '_' + uid;

      rawEvents.forEach(eventText => {
        const uid = getIcalValue(eventText, 'UID');
        if (!uid) return;

        presentUids.add(uid);
        properties.deleteProperty(missingKey(uid));

        const summary = getIcalValue(eventText, 'SUMMARY') || 'Appointment';
        const startInfo = getIcalProperty(eventText, 'DTSTART');
        const endInfo = getIcalProperty(eventText, 'DTEND');

        if (!startInfo || !endInfo) return;

        const startDate = parseIcalDate(startInfo);
        const endDate = parseIcalDate(endInfo);

        if (!startDate || !endDate) return;

        const marker = account.markerPrefix + uid;
        const propertyKey = account.key + '_' + uid;
        const existingEvents = eventsByMarker.get(marker) || [];

        // Keep past appointments untouched
        if (startDate <= now) {
          skipped++;
          return;
        }

        // Update existing appointment if times or title changed
        if (existingEvents.length > 0) {
          const existing = existingEvents[0];

          const needsTimeUpdate = (
            existing.getStartTime().getTime() !== startDate.getTime() ||
            existing.getEndTime().getTime() !== endDate.getTime()
          );

          const needsTitleUpdate = existing.getTitle() !== summary;

          if (needsTimeUpdate || needsTitleUpdate) {
            if (needsTimeUpdate) {
              existing.setTime(startDate, endDate);
            }
            if (needsTitleUpdate) {
              existing.setTitle(summary);
            }
            updated++;
          } else {
            skipped++;
          }

          // Clean up duplicate copies with the same UID
          for (let i = 1; i < existingEvents.length; i++) {
            existingEvents[i].deleteEvent();
            deleted++;
          }

          properties.setProperty(propertyKey, 'copied');
          return;
        }

        // Respect manually deleted events
        if (properties.getProperty(propertyKey)) {
          skipped++;
          return;
        }

       // Create new appointment
        const newEvent = calendar.createEvent(
          summary,
          startDate,
          endDate,
          { description: marker }
        );

        newEvent.addPopupReminder(30); // Sets a notification 30 minutes before
        properties.setProperty(propertyKey, 'copied');
        eventsByMarker.set(marker, [newEvent]);
        created++;
      });

      // Check for cancelled / missing events for this specific account
      if (rawEvents.length > 0) {
        eventsByMarker.forEach((matchingEvents, marker) => {
          if (!marker.startsWith(account.markerPrefix)) return;

          const uid = marker.substring(account.markerPrefix.length);
          if (!uid || presentUids.has(uid)) return;

          const propertyKey = account.key + '_' + uid;
          const countKey = missingKey(uid);

          const previousCount = Number(properties.getProperty(countKey) || 0);
          const newCount = previousCount + 1;

          Logger.log(account.name + ' - Missing UID: ' + uid + ', consecutive checks: ' + newCount);

          if (newCount >= 3) {
            matchingEvents.forEach(event => {
              if (event.getStartTime() > now) {
                event.deleteEvent();
                deleted++;
              }
            });

            properties.deleteProperty(propertyKey);
            properties.deleteProperty(countKey);
            eventsByMarker.delete(marker);
          } else {
            properties.setProperty(countKey, String(newCount));
          }
        });
      }

      Logger.log(account.name + ' Sync Results -> Created: ' + created + ', Updated: ' + updated + ', Skipped: ' + skipped + ', Deleted: ' + deleted);

    });

  } finally {
    lock.releaseLock();
  }

}


// ==========================================
// ICAL HELPER FUNCTIONS
// ==========================================

function getIcalProperty(eventText, propertyName) {
  const regex = new RegExp('^' + propertyName + '((?:;[^:]*)?):(.*)$', 'm');
  const match = eventText.match(regex);
  return match ? {
    parameters: match[1],
    value: match[2].trim().replace(/\\,/g, ',').replace(/\\;/g, ';').replace(/\\n/g, '\n')
  } : null;
}

function getIcalValue(eventText, propertyName) {
  const property = getIcalProperty(eventText, propertyName);
  return property ? property.value : null;
}

function parseIcalDate(info) {
  const value = info.value;

  if (/^\d{8}$/.test(value)) {
    return Utilities.parseDate(value, 'UTC', 'yyyyMMdd');
  }

  if (value.endsWith('Z')) {
    return Utilities.parseDate(value, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  }

  const tzMatch = info.parameters.match(/TZID=([^;:]+)/);
  const timezone = tzMatch ? tzMatch[1] : Session.getScriptTimeZone();

  return Utilities.parseDate(value, timezone, "yyyyMMdd'T'HHmmss");
}
```

2. Configure your Target Calendar:
On the const TARGET_CALENDAR_ID = '...'; line at the top, replace the placeholder with the actual Google email address under which your calendar lives (or leave it as your primary calendar email).

3. Add your Jane iCal Feed URL(s):
Look at the const JANE_ACCOUNTS section. For each Jane account you want to sync, update the following fields:

url: Paste your Jane iCal feed URL inside the quotation marks. (Note: You can grab your iCal feed URL from your Jane practitioner profile under your settings).

<img width="393" height="353" alt="Screenshot 2026-09-10 at 12 10 36" src="https://github.com/user-attachments/assets/2fb58867-6b43-4b94-88ef-b626019c0bcf" />

<img width="701" height="276" alt="Screenshot 2026-09-10 at 12 11 18" src="https://github.com/user-attachments/assets/045ab34e-d39e-4748-978c-42d76ebacdf2" />

markerPrefix: Give it a unique work site label (e.g., 'Clinic One: ' or 'Private Practice: '). This is what the script writes into the description to track reschedules and cancellations.

If you only have one clinic, you can use the code as is. If you want to add more clinics, add more blocks under `const JANE_ACCOUNTS = [` so it'll look like this:
```
const JANE_ACCOUNTS = [
  {
    name: 'Clinic One',
    url: 'https://...',
    markerPrefix: 'Jane UID C1: ',
    key: 'clinic_one'
  },
  {
    name: 'Clinic Two',
    url: 'https://...',
    markerPrefix: 'Jane UID C2: ',
    key: 'clinic_two'
  },
  {
    name: 'Clinic Three', // <--- Just add new ones here!
    url: 'https://...',
    markerPrefix: 'Jane UID C3: ',
    key: 'clinic_three'
  }
];
```
4. Set up the Automatic Trigger:

Click on the alarm clock icon on the left sidebar of the Apps Script editor, then click "Triggers".

Click "+ Add Trigger" in the bottom right.

Choose `syncJaneAppointments` to run, select "Time-driven", and set it to run as frequently as you like (e.g., Every 15 minutes or Every hour).

It will prompt you to authorize the script permissions—follow the prompts to allow it to access your calendar.

5. Test Your Script:

Go back to your code by clicking the < > icon (Editor) on the left sidebar.

Click the "Run" button on the top toolbar.

Check the Execution Log at the bottom; if it works successfully, it will print out how many appointments it found, created, or updated!

Note on Cancellations & Reschedules: this version automatically handles changes. If an appointment is rescheduled in Jane, the script updates the time in Google Calendar. If an appointment is canceled in Jane, the script will automatically remove it from your Google Calendar after 3 consecutive successful checks (to prevent accidental deletions from brief feed glitches).
