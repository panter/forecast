Revenue Forecast Spreadsheet
========

This script will load all project (except projects with the state "lost" or "closed") from the controllr
and generates a revenue forecast for a period of max. 2 years.

## Installation:

1. Set the AUTH_TOKEN in the RevenueForecast.gs script. This is the authentication token from a controllr user with API privileges.
2. Set the API_HOST_URL in the RevenueForecast.gs script. This is the root url to your controllr server.
3. Create a new spreadsheet "Revenue Forecast" on your Google Drive.
4. Create a new script in the spreadsheet from the menu "Tools > Script editor"
5. Paste in the contents of the file "RevenueForecast.gs" and save
6. Go to your spreadsheet and reload. Load the Revenue Forecast Data from the Menu "Controllr > Aktuellen Revenue Forecast laden"
