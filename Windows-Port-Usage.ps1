# MIT License
#
# Copyright (c) 2019 Toni Kautto
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

<#
.SYNOPSIS
    Get snapshot of current port usage for all current Windows processes, and store 
    the result in CSV files for furhter analysis.  
.DESCRIPTION
    This script gets current running processes and open connections for each process. 
    Additionally some basic protocol details are collected, like dynamic port range. 
    The result is stored into a CSV files, which can be consumed into Qlik Sense for 
    analysis. Files are named with Windows hostname and timestamp, to enable collecting 
    information over time and from multiple hosts in parallel. For space efficiency the 
    CSV files are compressed to ZIP archive by default. 
    Recommendation is to execute script on regular interval through Windows scheduled 
    task to get a view of port usage over time.      
.PARAMETER OutputFolder
    This paramater allows to define a custom output folder location, for example useful 
    if the output from multiple nodes should be collected in a central file share.     
    Default output is to a subfolder named "PortTraces" in the same location as this 
    script file. 
.PARAMETER IncludeUDP
    This flag enables collection of UDP port allocation.     
    By default only TCP is traced. 
.PARAMETER NoZip
    Generated CSV files are by default added in a ZIP archive to save storage space. 
    This flag leaves the genrated CSV files without compressing them in a ZIP archive.
.EXAMPLE 
    ./Windows-Port-Usage.ps1

    The default option collects information for TCP connections, and excludes UDP 
    connections. The generates CSV files are automatically compressed to a ZIP archive 
    for storage space efficiency. 
.EXAMPLE 
    ./Windows-Port-Usage.ps1 -OutputFolder "\\MyFileServer\PortTraces\" -IncludeUDP -NoZip

    This execution writes traces to fileserver, which means that logs from multiple 
    nodes can be collected to the same central location. Also UDP traces are included. 
    The genearted CSV files are not compressed to ZIP archive, which enables a direct 
    load into Qlik Sense app
.EXAMPLE 
    ./Windows-Port-Usage.ps1 -OutputFolder "\\MyFileServer\PortTraces\" -NoZip

    Write trace to file share, which means that logs from multiple nodes can be 
    collected to the same central location. Only collects TCP port consumption.
.EXAMPLE 
    ./Windows-Port-Usage.ps1 -IncludeUDP

    This execution writes traces to same folder as script. Also UDP traces are included.
#>

param (
    [string] $OutputFolder = ".\PortTraces\", 
    [switch] $IncludeUDP   = $false,
    [switch] $NoZip        = $false     
)

# Define desired output location 
# Note1: This folder will be created if it does not already exists
# Note2: Qlik Sense analysis app must be updated to target same folder 
 
# Create folder
New-Item -ItemType Directory -Force -Path "$OutputFolder" | Out-Null

# Get date and time of script execution in format YYYYMMDDThhmmss+ZZZZ
$ExecutionTimeStamp = Get-Date -Format o | ForEach-Object {$_ -replace "[-:]|(\.[0-9]{7})"}
$ExecutionDate      = Get-Date -Format "yyyyMMdd"

# Generate name for CSV output files
$CsvProcessList    = "$OutputFolder$env:computername`_Processes_$ExecutionTimeStamp.csv"
$CsvTcpConnections = "$OutputFolder$env:computername`_TcpConnections_$ExecutionTimeStamp.csv"
$CsvUdpConnections = "$OutputFolder$env:computername`_UdpConnections_$ExecutionTimeStamp.csv"
$CsvTcpSettings    = "$OutputFolder$env:computername`_TcpSettings_$ExecutionTimeStamp.csv"
$CsvUdpSettings    = "$OutputFolder$env:computername`_UdpSettings_$ExecutionTimeStamp.csv"
$ZipOutputArchive  = "$OutputFolder$env:computername`_PortUsage_$ExecutionDate.zip"

# Get running processes
# Store result into CSV file, including hostname and execution time

Get-Process | Select-Object Id, ProcessName, Product | `
Select-Object @{Name='Timestamp';Expression={$ExecutionTimeStamp}}, `
              @{Name='HostName'; Expression={$env:computername}}, `
              * | `
Export-Csv -Path "$CsvProcessList" -NoTypeInformation

# Get currently open TCP and UDP connections
# Only get UDP conneciotns if flagged for inclusion
# Store result into CSV files, including hostname and execution time

Get-NetTCPConnection | Select-Object State,CreationTime,LocalAddress,LocalPort,RemoteAddress,RemotePort,OwningProcess | `
Select-Object @{Name='Timestamp';Expression={$ExecutionTimeStamp}}, `
              @{Name='HostName'; Expression={$env:computername}}, `
              @{Name='Protocol'; Expression={"TCP"}}, `
              * | `
Export-Csv -Path "$CsvTcpConnections" -NoTypeInformation

If($IncludeUDP) {

    Get-NetUDPEndpoint | Select-Object CreationTime,LocalAddress,LocalPort, OwningProcess | `
    Select-Object @{Name='Timestamp';Expression={$ExecutionTimeStamp}}, `
                @{Name='HostName'; Expression={$env:computername}}, `
                @{Name='Protocol'; Expression={"UDP"}}, `
                * | `
    Export-Csv -Path "$CsvUdpConnections" -NoTypeInformation

}

# Get current dynamic port range
# Only get UDP range if flagged for inclusion
# Store result into CSV files, including hostname and execution time

Get-NetTCPSetting | Select-Object PolicyRuleName, DynamicPortRangeStartPort, DynamicPortRangeNumberOfPorts | `
Select-Object @{Name='Timestamp';Expression={$ExecutionTimeStamp}}, `
              @{Name='HostName'; Expression={$env:computername}}, `
              @{Name='Protocol'; Expression={"TCP"}}, `
              * | `
Export-Csv -Path "$CsvTcpSettings" -NoTypeInformation

If($IncludeUDP) {

    Get-NetUDPSetting | Select-Object PolicyRuleName, DynamicPortRangeStartPort, DynamicPortRangeNumberOfPorts | `
    Select-Object @{Name='Timestamp';Expression={$ExecutionTimeStamp}}, `
                    @{Name='HostName'; Expression={$env:computername}}, `
                    @{Name='Protocol'; Expression={"UDP"}}, `
                    * | `
    Export-Csv -Path "$CsvUdpSettings" -NoTypeInformation

}

# Append CSV files to ZIP 
# Remove the al collected files that come from this execution
# Leave all files uncompressed if NoZip flag was used
If (! $NoZip) {
    Compress-Archive -Path "$OutputFolder*$ExecutionTimeStamp*.csv" -Update -DestinationPath "$ZipOutputArchive"
    Remove-Item "$OutputFolder*$ExecutionTimeStamp*.csv"
}
