#Ultimate_User_and_Office_Script.ps1
#A1C Velder, Ian
#Created: March 13, 2024
<#Updates:
    -Added GUI for file selection. (3/14/2024)
    -Creates file path for output if it doesn't exist. (3/14/2024)
#>    
#References: YevsUserandOfficeScript.ps1

#Get file path to list of computers
Function Get-FileName($initialDirectory)
{  
 [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = “All files (*.*)| *.csv*”
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

# *** Entry Point to Script ***

$path = Get-FileName -initialDirectory “c:fso"

If ($path -eq "$null"){
    Write-Host "File not selected, terminating operation." -ForegroundColor Magenta
     
    exit
}

#Start timer
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

#File path to computers
$comps = Import-Csv -Path $path

#Information gathering loop
foreach ($comp in $comps) {
    $computername = $comp.ComputerName
    
#Attempt to pull data from registry
    try {
        Write-Host "$computername : Querying registry..." -ForegroundColor Cyan

        $data = Invoke-Command -ComputerName $computername -ScriptBlock {Get-ItemProperty lastloggedondisplayname -path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI'} -ErrorAction Stop
        $data2 = $data.LastLoggedOnDisplayName
        $array = $data2 -split "USAF"
        $unit = $array[1]
        $name = $array[0]
        Add-Member -InputObject $comp -MemberType NoteProperty -Name "User" -Value $name
        Add-Member -InputObject $comp -MemberType Noteproperty -Name "Unit" -Value $unit

        Write-Host "$computername : Data successfully pulled from registry." -ForegroundColor Green

        continue
    }

#Error handling if can't pull from registry
    catch {
        Write-Host "$computername : Registry query failed. Attempting to pull data from Computer_Stats folder..." -ForegroundColor Yellow        
    }

#Try to get data from stats folder if registry query failed
    try {
        $stats = Import-CSV -Path "path\to\sharefolder\with\computerinfo\here" -ErrorAction Stop
        $ID = $stats.EDIPI
        $data = Get-ADUser $ID -Properties *
        $data2 = $data.DisplayName
        $array = $data2 -split "USAF "
        $unit = $array[1]
        $name = $array[0]
        Add-Member -InputObject $comp -MemberType NoteProperty -Name "User" -Value $name
        Add-Member -InputObject $comp -MemberType NoteProperty -Name "Unit" -Value $unit

        Write-Host "$computername : Data successfully pulled from Computer_Stats folder." -ForegroundColor Green

        continue
    }

#Report that both methods failed
    catch {
        Write-Host "$computername : Data not found." -ForegroundColor Red
    }
    
}

#Test for existing export path. Create path if it doesn't exist
$testpath = Test-Path -Path "C:\Temp\User_and_Office"

If ($testpath -eq $False) {
    Write-host "Creating file path..." -ForegroundColor Gray
    $exportpath = New-Item -ItemType Directory -Path "C:\Temp\User_and_Office"
    Write-host "File path created: C:\Temp\User_and_Office" -ForegroundColor Gray
}
    
#Export data to .csv
$comps | Export-Csv -Path "C:\Temp\User_and_Office\User_and_Office_Results.csv" -NoTypeInformation -Force

#Save timestamp for completed operation
[int]$elapsedminutes = $stopwatch.Elapsed.Minutes
[int]$elapsedseconds = $stopwatch.Elapsed.Seconds

$stopwatch.Stop()

#GUI prompt asking to open .csv output once operation is complete
$title    = 'Operation Complete!'
$question = 'Open CompResults.csv?'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

$decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
if ($decision -eq 0) {
    Invoke-Item -Path "C:\Temp\User_and_Office\User_and_Office_Results.csv"
    Write-Host "----------------------------------------------" -ForegroundColor Yellow
    Write-Host "Operation Complete!" -ForegroundColor Yellow
    Write-Host "Time Elapsed: $elapsedminutes minutes $elapsedseconds seconds." -ForegroundColor Yellow
    Write-Host "Have a nice day!" -ForegroundColor Yellow
    Write-Host "----------------------------------------------" -ForegroundColor Yellow
} 

else {
    Write-Host "----------------------------------------------" -ForegroundColor Yellow
    Write-Host "Operation Complete!" -ForegroundColor Yellow
    Write-Host "Time Elapsed: $elapsedminutes minutes $elapsedseconds seconds." -ForegroundColor Yellow
    Write-Host "Have a nice day!" -ForegroundColor Yellow
    Write-Host "----------------------------------------------" -ForegroundColor Yellow
}

