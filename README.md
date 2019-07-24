# Windows Port Allocation Per Process

This project contains a simple Powershell script that collect process and port allocation data from Windows, and stores the results to CSV files. By runnig the script at scheduled intervals trends over time can be analysed through the attahced Qlik Sense app.  

## Configuration and setup

### Powershell script

1. Copy the script Windows-Port-Usage.ps1 to local storage
1. Run the script manually to confirm it works
1. Confirm script output in `C:\TEMP\ProcessUsage` <br />
Note, the output path can be altered in script variabel `$OutputFolder` 

### Windows Task

1. Configure a Windows Scheduled Task to execute this regularly:
    * Start > Task Scheduler
    * Highlight Task Scheduler Library
    * Create Task
1. General Tab Config
    * Name: Arbitrary
    * Account: This needs to be a local admin account
    * Run regardless of whether the user is logged in
    * Example screenshot: 
    ![Screenshot of setting up a task](https://i.imgur.com/wmMvuyF.png)
1. Triggers config:
    * 15 minutes is an ideal frequency
    * Be sure to enable the trigger
    * Example Screenshot:
    ![Screenshot of setting up a task](https://i.imgur.com/UE4Aekz.png)
1. Actions Tab Config:
    * Start a program
    * Path: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`
    * Arguments: `-ExecutionPolicy Bypass C:\Temp\Windows-Port-Usage.ps1`
    * Example Screenshot:
    ![Screenshot of setting up a task](https://i.imgur.com/Xm2soVK.png)
1. Manually run the task & make sure it outputs a CSV

### Qlik Sense App

1. Import app
1. Configure data connection
    * Default conenciton points to C:\TEMP\ProcessUsage
    * Adjust to desired location of collected CSV files
1. Reload app  

## Original Author

[levi-turner](https://github.com/levi-turner)

## License

This project is provided "AS IS", without any warranty, under the MIT License - see the [LICENSE](LICENSE) file for details