## Background
I am the coach and admin of the Madison Middle School ultimate frisbee team.  We have between 50-100 players in 6th to 8th grade.

You can get more context on this season from https://docs.google.com/document/d/1A2F7ThHtcMm23bxk8-30rMT2svaqT3gMbRWeSR_QXXY/view

I need your help to create my team's roster based off of some different data sources.  I am experienced at software engineering, but I am not proficient and banging out python pandas scripts, spreadsheet merges, or web applications on my own.

## Terminology
I will use "player" and "student" interchangeably.
I will use "parent" and "caretaker" interchangeably.

### Requirements for Roster Synthesis from Data Sources
I need functionality that will ultimate build or update the team roster, which is stored in the "Roster" sheet of https://docs.google.com/spreadsheets/d/1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8/edit?gid=267530468#gid=267530468

This will use similar/same data sources as Stage 1.

The roster sheet uses the first 5 rows to specify the intended data for the column.  This metadata includes:
- Column Name: the name of the column for the roster sheet.
- Type: the type of the data (e.g., string, email address)
- Source: what data source this should be populated from.  It usually specifies the data source name and the data source column.  If no data source colum is provided, it can be assumed to share the same name as the "column name"
- Additional note: provides extra notes on how the field should be set.
Study these columns since 

There should be one row per player in the roster sheet, and the list of players is derived from Final Forms.

You can either edit the sheet inline or generate a new CSV or Spreadsheet.  I am partial to Google Sheets, but I'm flexible.

You could also create a spreadsheet that imports all of these datasources and then does a join within the spreadssheet.  I think an important decision is "where do you join" the data sources?  Do you do that outside of the spreadsheet and stick it into the roster or do you import all the data into the spreadsheet and then use Spreadheet formulas to join.  

You are going to have join across these data sources.
Please analyze your "join keys", but it will often be player name or email address.
Unforunately though because this is human-entered data, it's possible that there is misspellings between the data sources and so you may need to do fuzzy matching or create a translation layer.
Be lenient on misspelling and do fuzzy matching if it helps.  

## Data Sources
### SPS Final Forms
Exports of this data is https://drive.google.com/drive/folders/1SnWCxDIn3FxJCvd1JcWyoeoOMscEsQcW?usp=drive_link
Use the most recent export.  I can also give you a direct link if that helps.
The export time is in the filename as a ISO8601 datetime (e.g., 2025-09-05T05:15:38Z).
There is one row per player that signed up.  
There is contact info for at least one caretaker, but likely two.
I am only interested in some of the columns from this data source.
In the roster metadata, I usually refer to this data source as "FinalForms".  

### Additional Questionaire for Coaches
Responses are in this sheet: https://docs.google.com/spreadsheets/d/1f_PPULjdg-5q2Gi0cXvWvGz1RbwYmUtADChLqwsHuNs/edit?usp=sharing .  
This updates as new entries come in.
There should be one entry per player.
I usually just want to bring all of the columns from this data source into the roster.
In the roster metadata, I usually refer to this data source as "AdditionalInfoForm".  

### Team Mailing List
The team mailing list is madisonultimatefall25@googlegroups.com.
You can see exports of who is on the mailing list in https://drive.google.com/drive/folders/1pAeQMEqiA9QdK9G5yRXsqgbNVzEU7R1E?usp=drive_link
Use the most recent export.  I can also give you a direct link if that helps.
The export time is in the filename as a ISO8601 datetime (e.g., 2025-09-05T05:15:38Z).
There is one row per player caretaker that signed up. 
The "Email address" and "Posting permissions" column are most pertinent.
In the roster metadata, I usually refer to this data source as "MailingList".

## Conclusion
Does the goal make sense?  I want to understand and review your plan before you execute.


# Links:
Roster: https://docs.google.com/spreadsheets/d/1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8/edit?gid=267530468
FinalForms: https://drive.google.com/file/d/1pWUIw2rM0MfNWnaC3Ltsz6Wj8_PGFHrH/view?usp=drive_link
AdditionalInfoForm: https://docs.google.com/spreadsheets/d/1f_PPULjdg-5q2Gi0cXvWvGz1RbwYmUtADChLqwsHuNs
MailingList: https://drive.google.com/file/d/1n0h81l31lsGvvSPrZUT5SOuS6jXT4h6E/view?usp=drive_link
Service account email: stevel@cedar-scene-471205-t3.iam.gserviceaccount.com