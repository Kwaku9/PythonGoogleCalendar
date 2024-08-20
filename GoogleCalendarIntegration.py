import datetime
import os.path
import calendar
from datetime import timedelta

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("calendar", "v3", credentials=creds)

        def get_upcoming_events_for_day(service, day_of_week, time_of_day):
            print(f"Fetching events for {day_of_week} during {time_of_day}...")

            # Manually enforce the correct offset
            correct_offset = timedelta(hours=-4)
            now = datetime.datetime.now(datetime.timezone(correct_offset))

            # Calculate the start and end times for the day
            current_weekday = now.weekday()
            target_weekday = list(calendar.day_name).index(day_of_week.capitalize())
            days_ahead = (target_weekday - current_weekday) % 7

            start_date = (now + datetime.timedelta(days=days_ahead)).replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = start_date + datetime.timedelta(days=8)

            start_time = start_date.isoformat()
            end_time = end_date.isoformat()

            print(f"Fetching events between {start_time} and {end_time}...")

            events_result = service.events().list(calendarId='primary', timeMin=start_time, timeMax=end_time,
                                                  singleEvents=True, orderBy='startTime').execute()
            events = events_result.get('items', [])

            # Filter events for the specific day and time of day
            filtered_events = [event for event in events if matches_user_input(event, day_of_week, time_of_day)]
            print(f"Filtered events: {filtered_events}")
            return filtered_events

        def matches_user_input(event, day_of_week, time_of_day):
            # Check if the event matches the day of the week and time of day provided by the user
            event_time = datetime.datetime.fromisoformat(event['start']['dateTime']).astimezone(datetime.timezone(timedelta(hours=-4)))
            target_day = event_time.strftime('%A').lower() == day_of_week.lower()

            print(f"Checking event {event_time} for day {day_of_week} and time {time_of_day}...")

            if time_of_day == "morning":
                match = target_day and event_time.hour in [10, 11]
            elif time_of_day == "afternoon":
                match = target_day and event_time.hour in [12, 13, 14, 15]
            elif time_of_day == "evening":
                match = target_day and event_time.hour in [17, 18, 19, 20]
            else:
                match = False

            print(f"Match result: {match}")
            return match

        def is_conflict(slot, events):
            correct_offset = timedelta(hours=-4)
            slot_start = datetime.datetime.fromisoformat(slot).replace(tzinfo=datetime.timezone(correct_offset))
            slot_end = slot_start + datetime.timedelta(hours=1)

            print(f"Checking slot from {slot_start} to {slot_end} for conflicts...")

            for event in events:
                event_start = datetime.datetime.fromisoformat(event['start']['dateTime']).astimezone(datetime.timezone(correct_offset))
                event_end = datetime.datetime.fromisoformat(event['end']['dateTime']).astimezone(datetime.timezone(correct_offset))

                print(f"Against event from {event_start} to {event_end}...")

                # Check if there is any overlap between the slot and event
                if max(slot_start, event_start) < min(slot_end, event_end):
                    print("Conflict found!")
                    return True

            print("No conflict found.")
            return False

        def get_first_available_slot(service, day_of_week, time_of_day):
            timezone = datetime.timezone(datetime.timedelta(hours=-4))
            now = datetime.datetime.now(timezone)

            print(f"Current time: {now}")
            print(f"Current timezone offset: {now.utcoffset()}")

            # Fetch events for the current week
            events = get_upcoming_events_for_day(service, day_of_week, time_of_day)

            # Determine the possible slots based on time of day
            if time_of_day == "morning":
                slots = [10, 11]
            elif time_of_day == "afternoon":
                slots = [12, 13, 14, 15]
            elif time_of_day == "evening":
                slots = [17, 18, 19, 20]
            else:
                print("Invalid time of day provided.")
                return None

            current_weekday = now.weekday()
            target_weekday = list(calendar.day_name).index(day_of_week.capitalize())
            days_ahead = (target_weekday - current_weekday) % 7
            if days_ahead == 0 and now.hour >= max(slots):
                days_ahead = 7

            target_date = now + datetime.timedelta(days=days_ahead)

            # Check slots for the current week
            for hour in slots:
                slot_time = datetime.datetime(target_date.year, target_date.month, target_date.day, hour,
                                              tzinfo=timezone).isoformat()
                print(f"Checking slot: {slot_time}")
                if not is_conflict(slot_time, events):
                    print(f"Available slot found: {slot_time}")
                    return slot_time

            # If no slot is available, move to the next week and re-fetch events
            target_date += datetime.timedelta(days=7)

            # Calculate the new start and end of the next week
            start_date = target_date - datetime.timedelta(days=target_weekday)
            end_date = start_date + datetime.timedelta(days=7)

            # Fetch events for the entire next week
            start_time = start_date.isoformat()
            end_time = end_date.isoformat()

            print(f"Fetching events for next week between {start_time} and {end_time}...")
            events = service.events().list(calendarId='primary', timeMin=start_time, timeMax=end_time,
                                           singleEvents=True, orderBy='startTime').execute().get('items', [])

            for hour in slots:
                slot_time = datetime.datetime(target_date.year, target_date.month, target_date.day, hour,
                                              tzinfo=timezone).isoformat()
                print(f"Checking slot for next week: {slot_time}")
                if not is_conflict(slot_time, events):
                    print(f"Available slot found: {slot_time}")
                    return slot_time

            print("No available slot found within the next two weeks.")
            return None

        user_day = input("Enter the day of the week (e.g., Tuesday): ")
        user_time = input("Enter the time of day (morning, afternoon, evening): ")

        first_slot = get_first_available_slot(service, user_day, user_time)

        if first_slot:
            confirmation = input(
                f"Your appointment is available on {first_slot[:10]} at {first_slot[11:16]} (UTC). "
                f"Would you like me to confirm this for you? (yes/no): "
            )
            if first_slot:
                confirmation = input(
                    f"Your appointment is available on {first_slot[:10]} at {first_slot[11:16]} (UTC). "
                    f"Would you like me to confirm this for you? (yes/no): "
                )
                if confirmation.lower() == 'yes':
                    # Create the event
                    event = {
                        'summary': 'Booking with Travel Advisor',
                        'location': 'Zoom Call',
                        'description': 'Planning your trip to Puerto Rico, Q&A',
                        'start': {
                            'dateTime': first_slot,
                            'timeZone': 'UTC',
                        },
                        'end': {
                            'dateTime': (datetime.datetime.fromisoformat(first_slot) + datetime.timedelta(
                                hours=1)).isoformat(),
                            'timeZone': 'UTC',
                        },
                        'recurrence': [
                            'RRULE:FREQ=DAILY;COUNT=1'
                        ],
                        'attendees': [
                            {'email': 'lpage@example.com'},
                            {'email': 'sbrin@example.com'},
                        ],
                        'reminders': {
                            'useDefault': False,
                            'overrides': [
                                {'method': 'email', 'minutes': 24 * 60},
                                {'method': 'popup', 'minutes': 10},
                            ],
                        },
                    }

                    event = service.events().insert(calendarId='primary', body=event).execute()
                    print(
                        "Your appointment has been successfully booked! A confirmation email has been sent, and you will receive a reminder closer to the time.")
                else:
                    # If the user says no, restart the process
                    print("No appointment was made. Let's try again.")
                    main()  # Restart the process
            else:
                print("No available slot found for the given criteria.")
    except HttpError as error:
        print(f"An error occurred: {error}")

#    else:
#        print(f"Your appointment has been successfully booked! A confirmation email has been sent, and you can view your meeting details here: {event.get('htmlLink')}.")

if __name__ == "__main__":
    main()