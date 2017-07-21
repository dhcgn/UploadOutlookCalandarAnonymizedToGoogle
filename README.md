# UploadOutlookCalanderAnonymizedToGoogle

This program reads your events from Outlook for the next 3 months and uploads a very small subset of event information to the employees personal calendar.

The idea is, that the employee can benefit from BYOD but with no confidential data is leaked.

## Inhouse Exchange Calendar

Subject: Meeting with John Miller  
Time: 01.01.2017 10:00 - 12:00  
Location: Room 23  
Participant: CIO John Wick, James Smith
Content: Confidential Information XZY  
Attachements: confidential.pdf  

## Employee Calendar

Subject: MWJM  
Time: 01.01.2017 10:00 - 12:00  
Location: Room 23  
Participant: *null*  
Content: *null*  
Attachements: *null*  

## Details

**Only** these fields are uploaded:

| Name     | comment             |
|-|-|
| Subject  | **one Word one letter** (so the employee can derive thie origin event)|
| Time     | 1:1                 |
| Location | 1:1                 |

## Build this program

In order to build this program you need to create your own client secret from Google.

## Usage

Get all of your calender ids

`file.exe GetCalender`

Sync the events from Outlook to this calander (will drop existing events, so you really should use a seperate calander instance)

`file.exe Sync <Id>`
