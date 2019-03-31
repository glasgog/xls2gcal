# xls2gcal
Get a work shift from an excel file link and add it to a given google calendar


Actually work on excel file only. Open source format (ods) doesn't work.
You could use ods converting it to xlsx file, using Open/Libre Office
or doing it from terminal:
```
libreoffice --convert-to xlsx example.ods
```

Need the following libraries, that could be installed with pip:
```
sudo pacman -S python-pip
sudo pip install python-dateutil pytz xlrd
sudo pip install google-api-python-client oauth2client
```
Note: if you have both python 2.x and 3.x, use:
```
sudo python3 -m pip install libraries_list
```

Follow the intruction at [Google Calendar API Python Quickstart](https://developers.google.com/calendar/quickstart/python), and save the client configuraton credential file in the project folter at data/client_secret.json.

Some prior configuration files are needed:
* **xls.url**, simply containing the url of the xlsx file
* **conf.ini**, containing some configuration about excel structure and calendar.

conf.ini have the following structure:
```
[EXCEL]
sheet_number = 1
first_useful_row = 4
date_column = C
worker_column = AD

[CALENDAR]
calendar_name = Shift
days_to_read = 60
```

U can find further examples in example.xlsx, example_conf.ini and example_xls.url under data/example folder.
