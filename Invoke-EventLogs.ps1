New-Module -Name MyModule -ScriptBlock {

Function Invoke-EventLogs {

[cmdletbinding()]
Param()

#start with a clear screen
Clear-Host

$title = "Event Log Menu"
$menuwidth = 30
#calculate how much to pad eft to center the title
[int]$pad =($menuwidth/2)+($title.length/2)

#define a here string for the menu options
$menu = @"

1. Application
2. Security
3. Setup
4. System
5. Windows PowerShell
6. Quit

"@

Write-Host ($title.PadLeft($pad)) -ForegroundColor Cyan
Write-Host $menu -ForegroundColor Yellow

#Read-Host writes strings but we can specifically treat result as an integer

[int]$r = Read-Host "Select a menu choice"

#validate the value
if ((1..6) -notcontains $r ){
        Write-Warning "$r is not a valid choice"
        pause
        Invoke-EventLogs
}
elseif ((1..5) -contains $r){
    #get computername for first four menu choices
    $Computername = Read-Host "Enter a computername or press Enter to use the localhost"

    if ($Computername -notmatch "\w+"){
        $computername = $env:COMPUTERNAME
    }
}

#code to execute
Switch ($r){
    1{
        Get-EventLog -LogName Application -Newest 25 -ComputerName $Computername
    }
    2{
        Get-EventLog -LogName Security -Newest 25 -ComputerName $Computername
    }
    3{
        Get-EventLog -LogName Setup -Newest 25 -ComputerName $Computername
    }
    4{
        Get-EventLog -LogName System -Newest 25 -ComputerName $Computername
    }
    5{
        Get-EventLog -LogName "Windows PowerShell" -Newest 25 -ComputerName $Computername
    }
    6{
        Write-Host "Have a nice day" -ForegroundColor Green
        #bail out command
        Return
    }
} #switch

#insert a blank line
write-host ""
pause

#run this function again
&$MyInvocation.MyCommand

}#end function

} | Import-Module
