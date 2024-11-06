#!/usr/bin/python3

import sys
import os.path
from icalendar import Calendar
import csv

filename = sys.argv[1]
# TODO: use regex to get file extension (chars after last period), in case it's not exactly 3 chars.
file_extension = str(sys.argv[1])[-3:]
headers = ('UID', 'Summary', 'Organizer', 'Attendee', 'Description', 'Location', 'Start Time', 'End Time', 'URL')

class CalendarEvent:
    """Calendar event class"""
    uid = ''
    summary = ''
    organizer = ''
    attendee = ''
    description = ''
    location = ''
    start = ''
    end = ''
    url = ''

    def __init__(self, name):
        self.name = name

events = []


def open_cal():
    if os.path.isfile(filename):
        if file_extension == 'ics':
            print("Extracting events from file:", filename, "\n")
            f = open(sys.argv[1], 'rb')
            gcal = Calendar.from_ical(f.read())

            for component in gcal.walk():
                event = CalendarEvent("event")
                event.uid = component.get('UID')
                # if component.get('TRANSP') == 'TRANSPARENT': continue #skip event that have not been accepted
                if component.get('SUMMARY') == None: continue # skip events without a summary
                event.summary = component.get('SUMMARY')
                if component.get('ORGANIZER') != None:
                    event.organizer = component.get('ORGANIZER').params['CN'] + " <" + component.get('ORGANIZER') + ">"
                if component.get('ATTENDEE') != None:
                    event.attendee = component.get('ATTENDEE')
                if component.get('DESCRIPTION') != None:
                    event.description = component.get('DESCRIPTION')
                event.location = component.get('LOCATION')
                if hasattr(component.get('dtstart'), 'dt'):
                    event.start = component.get('dtstart').dt
                if hasattr(component.get('dtend'), 'dt'):
                    event.end = component.get('dtend').dt
                event.url = component.get('URL')
                events.append(event)
            f.close()
        else:
            print("You entered ", filename, ". ")
            print(file_extension.upper(), " is not a valid file format. Looking for an ICS file.")
            exit(0)
    else:
        print("I can't find the file ", filename, ".")
        print("Please enter an ics file located in the same folder as this script.")
        exit(0)


def csv_write(icsfile):
    csvfile = icsfile[:-3] + "csv"
    try:
        with open(csvfile, 'w') as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
            wr.writerow(headers)
            for event in sortedevents:
                values = (
                    event.uid, 
                    event.summary.encode('utf8').decode(), 
                    event.organizer, 
                    event.attendee, 
                    event.description.encode('utf8').decode(), 
                    event.location, 
                    event.start, 
                    event.end, 
                    event.url
                )
                wr.writerow(values)
            print("Wrote to ", csvfile, "\n")
    except IOError:
        print("Could not open file! Please close Excel!")
        exit(0)


def debug_event(class_name):
    print("Contents of ", class_name.name, ":")
    print(class_name.summary)
    print(class_name.uid)
    print(class_name.description)
    print(class_name.location)
    print(class_name.start)
    print(class_name.end)
    print(class_name.url, "\n")

open_cal()
# sortedevents=sorted(events, key=lambda obj: obj.start) # Needed to sort events. They are not fully chronological in a Google Calendard export ...
sortedevents = events
csv_write(filename)
#debug_event(event)
