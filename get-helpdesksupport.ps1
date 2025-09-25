<#
.Synopsis
    This is a script to gather information for Help Desk Support calls
.Description
    This is a basic script designed to gather user and computer inforamtion for help desk suport calls.
    Information gathered includes:
    DNS Name & IP Address
    DNS Server
    Name of Operating System
    Amount of Memory in target computer
    Amount of free space on disk
    Last Reboot of System
.Example
    Example of how to use this cmdlet
.Example
    Another example of how to use this cmdlet
#>

#Get-Helpdesksupport.ps1
#A1C Velder, Ian
#Created: October 18, 2023
#Updated
#References: PowerShell: Getting Started (Michael Bender)

##Parameters for Computername & UserName
Param (
[parameter(Mandatory=$true)][string]$ComputerName
)
#Variables
$Credential = Get-Credential
$CimSession = New-CimSession -ComputerName $ComputerName -Credential $Credential
$Analyst = $Credential.UserName
#Commands

#OS Description
    $OS= (Get-CimInstance Win32_OperatingSystem -ComputerName $ComputerName).caption

#Disk Freespace on OS Drive
    $drive = Get-WmiObject -class Win32_logicaldisk | Where-Object DeviceID -EQ 'C:'
    $Freespace = (($drive.Freespace)/1gb)

#Amount of System Memory
    $MemoryInGB = ((((Get-CimInstance Win32_PhysicalMemory -ComputerName $ComputerName).Capacity|measure -Sum).Sum)/1gb)
     $MemoryInGB

#Last Reboot of System
    $LastReboot = (Get-CimInstance -Class Win32_OperatingSystem -ComputerName $ComputerName).LastBootUpTime
     $LastReboot

#IP Address & DNS Name
    $DNS = Resolve-DnsName -Name $ComputerName | Where-Object Type -EQ "A"
    $DNSName = $DNS.Name
    $DNSIP = $DNS.IPaddress
    $IPInfo = Get-CimInstance Win32_NetworkAdapterConfiguration

#DNS Server of Target
    $DNSServer = (Get-DnsClientServerAddress -CimSession $CimSession -InterfaceAlias "ethernet" -AddressFamily IPv4).ServerAddresses

#Write Output to Screen
#Clear-Host
Write-Output "Help Desk Support Information for $ComputerName"
Write-Output "-----------------------------------------------"
Write-Output "Support Analyst: $Analyst";""
Write-Output "Computername: $ComputerName";""
Write-Output "Last System Reboot of $computername : $LastReboot ";""
Write-Output "DNS Name of $computername : $DNSName";""
Write-Output "IP Address of #DNSName : $DNSIP";""
Write-Output "DNS Server(s) for $ComputerName : $DNSServer";""
Write-Output "Total System Ram in $ComputerName : $MemoryInGB GB";""
Write-Output "Freespace on C: $Freespace GB";""
Write-Output "Version of Operating System: $OS"

