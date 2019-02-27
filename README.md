# Contents

* port_usage.ps1
  * PowerShell script which will dump the process usage bucketed into counts
* port_usage_2018-10-19T00.19.53.4636467+02.00.csv
  * Example output
* Process Usage.qvf
  * Qlik app to consume the results

# Config steps for Script:

* Place the PowerShell script somewhere on the server, ideally the C drive because that’s what I tested
  * Run it manually to confirm that it is working
    * Expected location is `C:\TEMP\ProcessUsage`
* Configure a Windows Scheduled Task to execute this regularly:
  * Start > Task Scheduler
  * Highlight Task Scheduler Library
  * Create Task
* General Tab Config
  * Name: Arbitrary
  * Account: This needs to be a local admin account
  * Run regardless of whether the user is logged in
  * Example screenshot:
![Screenshot of setting up a task](https://i.imgur.com/wmMvuyF.png)
* Triggers config:
  * 15 minutes is an ideal frequency
  * Be sure to enable the trigger
  * Example Screenshot:
![Screenshot of setting up a task](https://i.imgur.com/UE4Aekz.png)
* Actions Tab Config:
  * Start a program
  * Path: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`
  * Arguments: `-ExecutionPolicy Bypass C:\Temp\port_usage.ps1`
  * Example Screenshot:
![Screenshot of setting up a task](https://i.imgur.com/Xm2soVK.png)
* Manually run the task & make sure it outputs a CSV

# Config for Qlik App:
* Import app
* Reload
  * The Data Connection pointing to the CSVs is embedded to C:\TEMP\ProcessUsage
  * This may need adjusting to a UNC path
  * Feel free to adjust things 
