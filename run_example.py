#!/usr/bin/env python3

from O365 import Account
from icalendar import Calendar, Event
from html2text import html2text
import datetime as dt

credentials = ('client_id', 'client_secret')
credentials = ('7f41f8bf-b0d2-4f84-9f57-c1f813ab157e', 'ydBXJS90;[|csehnWLL650|')

account = Account(credentials)

cal = Calendar()
cal.add('prodid', '-//Oj D&B Calendar//mxm.dk//')
cal.add('version', '2.0')

scopes = ['offline_access', 'calendar']

account = Account(credentials)

if not account.is_authenticated:  # will check if there is a token and has not expired
    # ask for a login
    account.authenticate(scopes=scopes)

schedule = account.schedule()

calendar = schedule.get_default_calendar()

q = calendar.new_query('start').greater_equal(dt.datetime(2018, 5, 20))
q.chain('and').on_attribute('end').less(dt.datetime(2019, 7, 1))

events = calendar.get_events(query=q, include_recurring=True)
for event_in in events:
    body = event_in.body
    # print(dir(event_in))
    # exit()
    event_out = Event()
    event_out.add('summary', event_in.subject)
    event_out.add('dtstart', event_in.start)
    event_out.add('dtend', event_in.end)
    event_out.add('location', event_in.location)
    event_out.add('description', html2text(event_in.body))
#    event_out.add('X-ALT-DESC', event_in.get_body_text())
    event_out['uid'] = event_in.ical_uid
    cal.add_component(event_out)

    # print(event.location)
    # print(event.get_body_text())
    # print()



import tempfile, os
with open(os.path.join('.', 'example.ics'), 'wb') as f:
    f.write(cal.to_ical())
    f.close()

#    pprint(event)

#pprint(events)
