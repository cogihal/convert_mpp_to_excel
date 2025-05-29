# Convert Microsoft Project .mpp file to Excel Gantt Chart

## Description

This is a Python script that converts a Microsoft Project .mpp file into an Excel Gantt chart.

## Description of config.json

```
{
  "path_to_jvm" : the path of jvm.dll (Java VM)
  "font_name"   : font name that you want to use
  "tab_title"   : tab title string
  "start_date"  : start date for gantt chart in format YYYY/MM/DD
  "end_date"    : end date for gantt chart in format YYYY/MM/DD

  "holidays": [
    list of holidays in format YYYY/MM/DD
  ]
}
```

If the `JAVA_HOME` environment variable has been set already, "path_to_jvm" is not necessary.
If the `JAVA_HOME` environment variable is set, "path_to_jvm" will not be used even if it is set.

## Pre-requisites

- Java Vartual Machine (JVM).  
  mpxj requires JVM to run.

## Developing environments

- Python 3.13.3
- jpype1==1.5.2
- mpxj==14.0.0
- openpyxl==3.1.5
- pyreadline3==3.5.4

