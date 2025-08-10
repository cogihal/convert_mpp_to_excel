# Convert Microsoft Project .mpp file to Excel task list and gantt chart

## Description

This is a Python script that converts a Microsoft Project .mpp file into an Excel task list and gantt chart.

## How to use

1. Modify the 'config.toml' file refering the sample TOML file and to set the parameters for the gantt chart that you want to genarate.
1. Run the Python script.
1. The script will prompt you to input the path to the .mpp file you want to convert.
1. The script will convert a gantt chart in Excel format based on the .mpp file.
1. After converting the gantt chart, the script asks you to input the excel base file name to save.  
   The extention of the file name is used as '.xlsx' automatically.

## Description of config.toml

Prepare 'config.toml' by referring the sample TOML file. The configuration file name must be 'config.toml'.

```toml
# config.toml

path_to_jvm = full path file name of jvm.dll (Java VM) : ex. "C:\\jdk\\bin\\server\\jvm.dll"
font_name   = font name that you want to use to excel : ex. "Meiryo UI"
tab_title   = tab title string : ex. "Project Blue"
start_date  = start date for gantt chart in format "YYYY/MM/DD"
end_date    = end date for gantt chart in format "YYYY/MM/DD"

holidays = [
  list of holidays in format "YYYY/MM/DD", ...
]
```

If the `JAVA_HOME` environment variable has been set already, "path_to_jvm" is not necessary.
If the `JAVA_HOME` environment variable is set, "path_to_jvm" will not be used even if it is set.

## Pre-requisites

- Java Vartual Machine (JVM).  
  Because mpxj requires JVM to run.

