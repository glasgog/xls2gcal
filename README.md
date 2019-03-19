# xls2gcal
Get a work shift from an excel file and add it to a given google calendar


Actually work on excel file only. Open source format (ods) doesn't work.
You could use ods converting it to xlsx file, using Open/Libre Office
or doing it from terminal: libreoffice --convert-to xlsx example.ods

Need the following libraries, that could be installed with pip:
...
sudo pacman -S python-pip
sudo pip install python-dateutil pytz xlrd
sudo pip install google-api-python-client oauth2client
...
Note: if you have both python 2.x and 3.x, use:
...
sudo python3 -m pip install libraries_list
...

See [Google Calendar API Python Quickstart](https://developers.google.com/calendar/quickstart/python) for more information.
