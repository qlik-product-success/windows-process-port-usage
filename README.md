# Windows Port Allocation Per Process

This project contains a Powershell script that collect process and port allocation data from Windows, and stores the results to CSV files. By runnig the script at scheduled intervals through a Windows task, port allocation trends over time can be analysed through the attached Qlik Sense app.

## Configuration and setup

### Powershell script

1. Create folder `c:/temp` unless it already exists
1. Clone *or* download this GitHub project to `c:/temp`
    - Clone with Git <BR /> `git clone https://github.com/tonikautto/windows-process-port-usage` 
    - Download as ZIP <BR /> https://github.com/tonikautto/windows-process-port-usage/archive/master.zip

1. Open Powershell and run the script manually to confirm it works
    - Trace only TCP and write results to same folder as PS1 file
      <br />`Windows-Port-Usage.ps1`
    - Trace only TCP and write results to custom location
      <br />`Windows-Port-Usage.ps1 -OutputFolder "\\MyFileServer\PortTraces\"`
    - Trace TCP and UDP and write results to same folder as PS1 file
      <br />`Windows-Port-Usage.ps1 -InlcudeUDP`
    - Trace TCP and UDP and write results to custom location
      <br />`Windows-Port-Usage.ps1 -OutputFolder "\\MyFileServer\PortTraces\" -InlcudeUDP`

1. Confirm script output based selected output location

### Windows Task

5. Configure a Windows Scheduled Task to execute this regularly:
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
1. Actions Tab Config:<BR />
   *NOTE: Adjust argument based on prefernce option from step 3.*
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
