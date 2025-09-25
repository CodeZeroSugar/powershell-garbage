$import = Import-Csv -Path 'C:\Temp\CompName\Comps.csv'

$results = foreach ($comp in $import.ComputerName) { 
    $Ping = Test-Connection -ComputerName $comp -Count 1 -Quiet -ErrorAction SilentlyContinue 

    If ($Ping -ne $null) {
        Try {
            $boot = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $comp -Property * -ErrorAction Stop).LastBootUpTime
        }
        Catch {
            Write-Host "Access Denied to $comp" -ForegroundColor Yellow
            [PSCustomObject]@{
                ComputerName = $comp
                Status = "Access Denied"
                LastBootUpTime = $boot
             }
            Continue
        }

        $Date = (Get-Date).AddDays(-7)

        If ($boot -lt $Date) {
            Write-Host "$comp needs reboot" -ForegroundColor Red
            [PSCustomObject]@{
                ComputerName = $comp
                Status = "Needs Reboot"
                LastBootUpTime = $boot
            }
        }
        Else {
            [PSCustomObject]@{
                ComputerName = $comp
                Status = "No Reboot Needed"
                LastBootUpTime = $boot
            }
        }
    }
    Else {
        Write-Host "$comp is Offline" -ForegroundColor Cyan
        [PSCustomObject]@{
            ComputerName = $comp
            Status = "Offline"
            LastBootUpTime = $null
        }
    }
}

# Export results
$results | Export-Csv -Path 'C:\Temp\Results.csv' -NoTypeInformation

