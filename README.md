# classroom-automation
Manage common Google Classroom actions with Google Apps Script

## Features

- Script is container-bound to a Google Sheets spreadsheet
- Use spreadsheet to show outupt such as class and assignment listing
- Use spreadsheet as an interface to select desired courses aand assignments
- Use spreadsheet to set up batch assignment creation

## Install

### From existing spreadsheet

Open the spreadsheet [Classroom automation 2.0](https://docs.google.com/spreadsheets/d/1BdS3FRFQoaiZOdvSqowEK6FDRJNdk53JlJeWEFca2LU/template/preview) and make a copy of it (along with the attached Apps Script project) in your MyDrive folder (*not* to a shared folder). (Use the same Google account that you use for Google Classroom.)

### From scratch

Use the same Google account that you use for Google Classroom. Create a Google spreadsheet and create the following sheets and columns with the given names:
  - **courses**: `select` | `id` | `name` | `section` | `courseState` | `alternateLink`
  - **assignments**: `select` | `id` | `title` | `alternateLink` | `state` | `courseId` | `topicId` | `description` | `materials` | `maxPoints` | `due`
  - **submissions**: `select` | `id` | `courseId` | `courseWorkId` | `assignment` | `userId` | `email` | `title` | `url` | `state` | `late` | `maxPoints` | `draftGrade` | `assignedGrade` |
  - **batch**: `select` | `courseId` | `topic` | `title` | `sch_date` | `sch_time` | `due_date` | `due_time` | `points` | `material` | `description` |
      - Format columns `sch_date` and `due_date` as a date type, and columns `sch_time` and `due_time` as a time type. This helps to ensure that scripts can read these values as a Javascript Date type.

For convenience, format the `select` column in each sheet as a checkbox. 

Select *Extensions > Apps Script*. It will open a new project. Copy and paste app.js into it. Optionally you can copy and paste tests.js and uncomment tests you want to run. Refresh the spreadsheet window and a new menu named *Classroom* should appear. Start with *Refresh course list*. You will be asked to authorize the script to access your Google account; if you have multiple

## Scripts available

- **Refresh course list**
- **Refresh assignments list** for selected course
- **Refresh submissions list** for selected assignment
- **Merge submissions** of Google Docs for selected assignment into a single Doc
- **Batch assign** multiple assignments for entire course
- ... or write your own using existing classes and functions (documentation coming soon)

## Change log

### 2.1 (2024-08-29)

Allow optional fields in batch assign. Allow creation of assignments that are missing fields such as description, due date, schedule date, material, or topic. Title still required.

### 2.0 (2024-08-17)

Refactor codebase. No new features. Reimplement current feature set with class-based logic. This will make debugging and testing more reliable and systematic, especially when Google introduces subtle changes in API behavior. The new classes (especially DateTime, SheetTable, Course, MergeDoc) will enable more rapid build out of new features (e.g. collecting comments and generating grade reports), and are potentially reusable to automate other Google products such as mail, calendar, and drive.

### 1.5 (2023-06-14)

Use datetimes in batch assignment. In previous versions of batch assignment, you specified the scheduled date using separate columns for year, month, day, and hour. This made the date info not human-friendly to read, and it was difficult to do date arithmetic. Update the sheet ""batch"" so it has a single column to specify the scheduled date as a date object and scheduled time as a time object; same for the due date and due time. Update the function "spec_journal" so it accepts dates and times in this format and converts to the formats needed by Classroom API. Also added sheet named "convert-to-datetime" to show how previous users can easily convert old assignment specifications to use date and time objects.

Removed timely completion function. This never worked reliably, and given the potential to create user confusion over grading, it seemed prudent to remove this.

### 1.4 (2022-11-22)

Fix time zone dependence. In previous versions, the assignments list showed all due dates in Eastern Daylight Time, even those after Nov 06. Update the function "list_assignments_all" so it shows each date in appropriate local time for that date.

### 1.3 (2022-11-21)

Fix time zone error. In previous versions, the function that sets an assignment's due date hard coded the time zone to Eastern Daylight Time. Update "format_due_time" so it computes this from the user's local time.

### 1.2 (2022-11-03)

Fix merge submissions. Google imposed limits on how many changes an API call can make on a Google Doc, resulting in failed merge. Now 'merge_submissions' explicitly saves the merge doc after copying content from each submission, and reopens for the next content copy.

### 1.1 (2022-09-03)

Public beta. Initial feature set includes: viewing of courses, assignments, and submissions; merging submissions for an assignment into a single Google Doc; and batch assignment of journal-type Classwork.


## Update procedure

### In spreadsheet

1. Update **about** and **versions**
2. Set name in version history
3. Copy spreadsheet to public folder, and redact student info in **submissions**
4. Get url of the copy, append `/template/preview`

### In repo

1. Update **Change log** in `README.md`
2. In **Installation** of `README.md`, update the link to the template
3. Copy and paste app.js and tests.js
