import re
from datetime import datetime
from pytz import timezone
import pandas as pd
from enum import Enum
from json import loads, dumps
from ics import Calendar, Event

class ExcelColumns(str, Enum):
    START_DAY = 'startDay'
    END_DAY = 'endDay'
    START_TIME = 'startTime'
    END_TIME = 'endTime'
    NAME = 'name'
    DESCRIPTION = 'description'
    PLACE = 'place'
    GROUP = 'group'

    def __str__(self):
        return '%s' % self.value

firstAppointmentRow = 7

# read by default 1st sheet of an excel file
df = pd.read_excel('input.xlsx', 
                            skiprows=firstAppointmentRow-1, 
                            nrows=80, 
                            header=None,
                            usecols="E,G,H,I,K,L,N",
                            names=[
                                ExcelColumns.START_DAY, 
                                ExcelColumns.END_DAY, 
                                ExcelColumns.NAME,
                                ExcelColumns.START_TIME,
                                ExcelColumns.END_TIME,
                                ExcelColumns.PLACE,
                                ExcelColumns.GROUP
                            ]
                           )

# remove empty rows
df = df.dropna(subset=[ExcelColumns.NAME])
# filter out all appointments which have no date set
df = df.dropna(subset=[ExcelColumns.START_DAY])

def set_name(row):
    split_name = parseName(row, 0)
    return split_name if split_name else row[ExcelColumns.NAME]

def set_description(row):
    return parseName(row, 1)

def parseName(row, i):
    if ' - ' in row[ExcelColumns.NAME]:
        return row[ExcelColumns.NAME].split(' - ')[i]
    elif '(' in row[ExcelColumns.NAME]:
        return row[ExcelColumns.NAME].split('(')[i].replace(')', '') 
    elif ';' in row[ExcelColumns.NAME]:
        return row[ExcelColumns.NAME].split(';')[i]
    else:
        return None

def set_place(row):
    if pd.notna(row[ExcelColumns.PLACE]):
        return re.sub(r"Wassera[lfingen]*\.", "Wasseralfingen", row[ExcelColumns.PLACE])
    return None

def set_end_day(row):
    if pd.notna(row[ExcelColumns.END_DAY]):
        return row[ExcelColumns.END_DAY]
    return row[ExcelColumns.START_DAY]

df[ExcelColumns.DESCRIPTION] = df.apply(set_description, axis=1)
df[ExcelColumns.NAME] = df.apply(set_name, axis=1)
df[ExcelColumns.PLACE] = df.apply(set_place, axis=1)
df[ExcelColumns.END_DAY] = df.apply(set_end_day, axis=1)

def set_start(row):
    if pd.notna(row[ExcelColumns.START_TIME]):
        return pd.to_datetime(row[ExcelColumns.START_DAY]) + pd.to_timedelta(row[ExcelColumns.START_TIME].strftime("%H:%M:%S"))
    return row[ExcelColumns.START_DAY]

def set_end(row):
    if pd.notna(row[ExcelColumns.END_TIME]):
        return pd.to_datetime(row[ExcelColumns.END_DAY]) + pd.to_timedelta(row[ExcelColumns.END_TIME].strftime("%H:%M:%S"))
    return row[ExcelColumns.END_DAY]

df['start'] = df.apply(set_start, axis=1)
df['end'] = df.apply(set_end, axis=1)

df['start'] = pd.to_datetime(df['start']).dt.strftime('%Y-%m-%d%H:%M:%S')
df['end'] = pd.to_datetime(df['end']).dt.strftime('%Y-%m-%d%H:%M:%S')

calendar = Calendar()
tz = timezone('Europe/Berlin')
for index, row in df.iterrows():
    event = Event()
    event.name = row[ExcelColumns.NAME]
    event.description = row[ExcelColumns.DESCRIPTION]
    event.location = row[ExcelColumns.PLACE]
    event.begin = datetime.strptime(row['start'],'%Y-%m-%d%H:%M:%S').astimezone(tz)
    if row['start'].endswith('00:00:00') and row['end'].endswith('00:00:00'):
        event.make_all_day()
    try:
        event.end = datetime.strptime(row['end'],'%Y-%m-%d%H:%M:%S').astimezone(tz)
    except ValueError:
        event.duration = { "days": 0, "hours": 2 }
    calendar.events.add(event)

print(calendar.events)
with open('output.ics', 'w') as ics_file:
    ics_file.writelines(calendar.serialize_iter())

