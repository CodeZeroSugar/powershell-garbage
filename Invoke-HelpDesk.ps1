Function Invoke-HelpDesk {

[cmdletbinding()]
Param()

#start with a clear screen
Clear-Host

$title = "Help Desk Menu"
$menuwidth = 30
#calculate how much to pad eft to center the title
[int]$pad =($menuwidth/2)+($title.length/2)

#define a here string for the menu options
$menu = @"

1. Services
2. Processes
3. Operating System
4. Computer System
5. Check free disk space (MB)
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
        Invoke-HelpDesk
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
        Get-Service -ComputerName $Computername
    }
    2{
        Get-Process -Computername $Computername
    }
    3{
        Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Caption | Format-Table -HideTableHeaders
    }
    4{
        Get-CimInstance Win32_ComputerSystem | Select-Object Name,Domain,Model,Manufacturer | Format-List 
    } 
    5{
        $c = Get-CimInstance -ClassName win32_logicaldisk -ComputerName $computername -filter "deviceid='c:'"
        $c.FreeSpace/1mb
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
