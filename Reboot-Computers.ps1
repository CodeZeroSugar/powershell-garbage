#Reboot-Computers.ps1
#A1C Velder, Ian
#Created: February 27, 2024
#Updated: 
#References:

#Import list of computers
$import = Import-Csv -Path "C:\Temp\Compname\Comps.csv"
$comps = $import.ComputerName

$results = foreach ($comp in $comps) {
    
#check if online
    $ping = Test-Connection -ComputerName $comp -count 1 -Quiet -ErrorAction SilentlyContinue

#If online, get Last bootup time
    If ($ping -ne $null) {
        Try {
             $boot = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $comp -Property * -ErrorAction Stop).LastBootUpTime
            }

#If failed to obtain Last bootup time, proceed to next object in "foreach" loop
        Catch {              
               Write-Host "$comp : Access Denied." -ForegroundColor Yellow
               [PSCustomObject]@{
                ComputerName = $comp
                Status = "Access Denied"
                LastBootUpTime = $null
                }
               
               Continue
              }
                   
#Determine if reboot is required
        $date = (get-date).AddDays(-7)
        If ($boot -lt $date) {
#If reboot required           
            Try {
                 Invoke-Command -ComputerName $comp -ScriptBlock { shutdown.exe /r /t 600 /c "This device has not been restarted within the last 7 days. To remain compliant, a forced reboot will occur in 10 minutes. Please save your work." } -ErrorAction Stop
                 Write-Host "$comp : Rebooting..." -ForegroundColor Magenta
                }
#If restart fails, proceed to next object in "foreach" loop
            Catch {                  
                   Write-Host "$comp : Reboot failed :(" -ForegroundColor Red
                   [PSCustomObject]@{
                     ComputerName = $comp
                     Status = "Reboot Failed"
                     LastBootUpTime = $boot
                     }

                   Continue                 
                  }
 
 #Restart was successful                     
            Write-Host "$comp : Reboot successful!" -ForegroundColor Green
            [PSCustomObject]@{
                ComputerName = $comp
                Status = "Reboot Successful"
                LastBootUpTime = $boot
                }
           }

 #If no reboot required                        
         Elseif ($boot -gt $date) {
            Write-Host "$comp : No reboot required :)" -ForegroundColor DarkGreen
            [PSCustomObject]@{
                ComputerName = $comp
                Status = "No Reboot Required"
                LastBootUpTime = $boot
                }
           }
       }

#If offline
    Else {
          Add-Member -InputObject $comp -MemberType NoteProperty -Name "Status" -Value "Offline"
          Write-Host "$comp : Offline." -ForegroundColor Cyan
          [PSCustomObject]@{
                ComputerName = $comp
                Status = "Offline"
                LastBootUpTime = $null
                }
         }   
}

#Export results
$results | Export-Csv -Path C:\Temp\Results.csv -NoTypeInformation

              