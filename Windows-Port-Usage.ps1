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
    Get snapshot of Windows IPv4 and IPv6 port usage for all current Windows processes. 
.DESCRIPTION
    This script gets current running processes and open connections for each process. 
    The result is stored into a CSV files, which can be consumed into Qlik Sense for 
    analysis. Files are named with Windows hostname and timestamp, to enable collecting 
    information over time and from multiple hosts in parallel. 
    Recommendation is to execute script on regular interval through Windows scheduled 
    task to get a view of port usage over time.      
.PARAMETER OutputFolder
    Target folder for trace output. 
    By default output is to same folder as script file. 
.PARAMETER IncludeUDP
    Apply this flag to also trace UDP port consumption. 
    By default only TCP is traced. 
.EXAMPLE 
    ./Windows-Port-Usage.ps1 -OutputFolder "\\MyFileServer\PortTraces\" -IncludeUDP

    This execution writes traces to fileserver, which means that logs from multiple 
    nodes can be collected to the same central location. Also UDP traces are included.
.EXAMPLE 
    ./Windows-Port-Usage.ps1 -OutputFolder "\\MyFileServer\PortTraces\"

    Write trace to file share, which means that logs from multiple nodes can be 
    collected to the same central location. Only collects TCP port consumption.
.EXAMPLE 
    ./Windows-Port-Usage.ps1 -IncludeUDP

    This execution writes traces to same folder as script. Also UDP traces are included.
#>

param (
    [string] $OutputFolder = ".\PortTraces\", 
    [switch] $IncludeUDP   = $false       
)

# Define desired output location 
# Note1: This folder will be created if it does not already exists
# Note2: Qlik Sense analysis app must be updated to target same folder 
 
# Create folder
New-Item -ItemType Directory -Force -Path "$OutputFolder" | Out-Null

# Get time of script execution in format YYYYMMDDThhmmss+ZZZZ
# For example 20190717T121751+1000 for 17 July 2019 12:17:51 PM in GTM+10
$ExecutionTimeStamp = Get-Date -Format o | ForEach-Object {$_ -replace "[-:]|(\.[0-9]{7})"}

# Generate name for CSV output files
$CsvProcessList    = "$OutputFolder$env:computername`_Processes_$ExecutionTimeStamp.csv"
$CsvTcpConnections = "$OutputFolder$env:computername`_TcpConnections_$ExecutionTimeStamp.csv"
$CsvUdpConnections = "$OutputFolder$env:computername`_UdpConnections_$ExecutionTimeStamp.csv"

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