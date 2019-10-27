# Google-Calendar-Event-Adder

Finds the days you've been scheduled from an excel file and automatically creates Google Calendar Events with the appropriate event name and date


## Getting Started

* Go to this url: https://developers.google.com/calendar/quickstart/python?refresh=1
* Click 'Enable The Google Calendar API'
* Accept
* Click Download client configurations
* Copy the credentials.json file you just downloaded to this directory


### Setting up the Program 
This will only happen the first time you run the program. The program will ask for access to your google calendar, and create a token.pickle:
* Run program: ```python main.py```
* Sign in to google
* QuickStart wants access to google account - Allow

### Running the Program
Copy an excel file with the month's schedule to this folder

Run the program with:
```bash
python main.py
```

When you run the program it will ask for 3 things:
1. Name of the excel file: Enter the name exactly (capital letters, small letters and spaces should be taken care of)
2. Sheet number: Each Excel file has multiple sheets. Enter the sheet number you want to scan. (eg:- 1, 2, 3, ...)
3. Name: Enter the name which you want to search the file for


##### Example:
```
$ python main.py
Enter the name of the excel file: May Schedule.xlsx
Which sheet number of the excel file do you want to scan [1, 2, 3, ...]: 1
Enter the name you want to search in the excel file: Dave
```

* Run the program each time you want to scan a new excel sheet
* By the end of it, all the google calendar events will be created for you.
* Check [calendar.google.com](https://calendar.google.com/) or the calendar app on your phone to see the events.
* Each event is set with a reminder notification for 6 hours, 1 day, 2 days, and 3 days before the event.
