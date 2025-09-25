<#
.Synopsis
    Pulls basic user information from Active Directory.
.Description
    Pulls basic user information from 
    users in a domain
    in Active Directory. The script gathers the following properties: 
    DisplayName, EmployeeID, Email, telephoneNumber, City, and IAtrainingdate.
    The script will also look up the Last Computer the user was on via the User_Stats
    folder.  
.Example
    After running the script you will be prompted to input the user's EmployeeID.
    The script will run and the results will be written to the screen.
#>

#Script Name: Get-ADUserInfo.ps1
#Creator: A1C Velder, Ian
#Date: 12/11/2023
#Updated: 12/12/2023

#Variables - Stores path to active directory user in a variable and generates date.
    $ID = Read-Host "Please Enter the User's Employee ID Number"
    $ADUser = Get-ADUser -Filter "EmployeeId -eq '$ID'" -SearchBase "distinguished name here" -Properties * 
    $Date = Get-Date

#Enter Tasks Below as Remarks
#Get Display Name
    $DisplayName = $ADUser.DisplayName
    $Array = $DisplayName -split "USAF"
    $Name = $Array[0]
    $Office = $Array[1]
#Get Base
    $Base = $ADUser.City
#Get Employee ID
    $EmployeeID = $ADUser.EmployeeID
#Get Email
    $Email = $ADUser.EmailAddress
#Get Phone Number
    $Phone = $ADUser.telephoneNumber
#Get IAtrainingDate
    $IATrainingDate = $ADUser.iaTrainingDate
#Get Last Computer Used
    $Name2 = $ADUser.mailNickname
    $Lines = Import-CSV -Path "T:\LogData\Logon\Stats\User_Stats\*$Name2_$ID*"
    $LastComputer = $Lines.ComputerName
    $IPAddress = Resolve-DnsName $LastComputer -Type A | Select-Object -ExpandProperty IPAddress

   

#Write Output to Screen
Clear-Host
Write-Output "User Information for $Name"
Write-Output "This report was pulled on $Date"
Write-Output "------------------------------------------------"
Write-Output "User : $Name";""
Write-Output "ID Number : $EmployeeID";""
Write-Output "Email : $Email";""
Write-Output "Phone : $Phone";""
Write-Output "Office : $Office";""
Write-Output "Duty Station : $Base";""
Write-Output "IA Training : $IATrainingDate";""
Write-Output "Last Computer : $LastComputer";""
Write-Output "IP Address : $IPAddress";""
Write-Output "------------------------------------------------"

