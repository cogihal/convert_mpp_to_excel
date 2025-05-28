# Convert Microsoft Project .mpp file to Excel Gantt Chart

## Description

This is a Python script that converts a Microsoft Project .mpp file into an Excel Gantt chart.

## Description of config.json

```
{
  "font_name"  : font name that you want to use
  "tab_title"  : tab title string
  "start_date" : start date for gantt chart in format YYYY/MM/DD
  "end_date"   : end date for gantt chart in format YYYY/MM/DD

  "holidays": [
    list of holidays in format YYYY/MM/DD
  ]
}
```

## Pre-requisites

- Java Vartual Machine (JVM).

## Developing environments

- Python 3.13.3
- jpype1==1.5.2
- mpxj==14.0.0
- openpyxl==3.1.5
- pyreadline3==3.5.4

