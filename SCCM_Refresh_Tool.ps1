#SCCM_Refresh_Tool.ps1
#A1C Velder, Ian
#Created: February 19, 2025
<#Updates:
    
#>    
#References: 

#Begin Script


#Here's all the configuration manager actions and the method for running them.
$Polices = {

$SCCMActions = @( "{00000000-0000-0000-0000-000000000121}",  #AppDeployment
                  "{00000000-0000-0000-0000-000000000003}",  #DiscoveryData
                  "{00000000-0000-0000-0000-000000000010}",  #FileCollection
                  "{00000000-0000-0000-0000-000000000001}",  #HardwareInventory
                  "{00000000-0000-0000-0000-000000000021}",  #MachinePolicy                  
                  "{00000000-0000-0000-0000-000000000002}",  #SoftwareInventory
                  "{00000000-0000-0000-0000-000000000031}",  #SoftwareMetering
                  "{00000000-0000-0000-0000-000000000114}",  #SoftwareUpdateDeployment
                  "{00000000-0000-0000-0000-000000000113}",  #SoftwareUpdateScan
                  "{00000000-0000-0000-0000-000000000032}")  #WindowsInstallerSource
                                   
   $error.clear()

   foreach ($action in $SCCMActions) {
        Try{       
                       
             (Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule $action -ErrorAction Stop).ReturnValue

        }

        Catch{
            
            Write-Host "$action failed on $env:COMPUTERNAME" -ForegroundColor Red

        }
      
    }

}

#Function for select file dialog
Function Get-FileName($initialDirectory){  
 [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = “All files (*.*)| *.txt*”
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename

} 

#Function for menu
Function Invoke-Menu {

[cmdletbinding()]
Param()

Clear-Host

$title = "----------Configuration Manager Refresh Tool----------"
$menuwidth = 30
#calculate how much to pad eft to center the title
[int]$pad =($menuwidth/2)+($title.length/2)

#define a here string for the menu options
$menu = @"

1. Local Host
2. Remote Host(s)
3. Read Me
4. Exit

"@

Write-Host ($title.PadLeft($pad)) -ForegroundColor Cyan
Write-Host $menu -ForegroundColor Yellow

[int]$r = Read-Host "Welcome to the Configuration Manager Refresh Tool. Please select an option"

#validate the value
if ((1..4) -notcontains $r ){
        Write-Warning "$r is not a valid choice"
        pause
        Invoke-Menu

}

#code to execute
Switch ($r){
    1{
        Clear-Host
        RefreshSCCM-Local
    }
    2{
        Clear-Host
        RefreshSCCM-Remote
    }
    3{
        Clear-Host 
        Write-Host "This is the Configuration Manager Refresh Tool, it can be used to troubleshoot issues related to Software Center." -Foregroundcolor Yellow
        Write-Host "Using this script will automate running the Actions for the Policies in the Configuration Manager Properties applet." -Foregroundcolor Yellow
        Write-Host "Select option 1 to only run the actions on this local machine." -Foregroundcolor Yellow
        Write-Host "Selecting option 2 will prompt you to select a .txt file containing a list of computer names. The script will automate running the actions on all computers in the .txt file." -foregroundcolor Yellow
        Write-Host "Administrator privileges are required to run this script!" -ForegroundColor Cyan
        Write-Host "Hopefully you find this tool useful!" -Foregroundcolor Yellow
        
        Pause
        Invoke-Menu          
    }
    4{
        Clear-Host
        Write-Host "Have a nice day!" -ForegroundColor Green
        #bail out command
        Return
    }
} #switch

}

#Function for local
Function RefreshSCCM-Local {
    
    Invoke-Command -ScriptBlock $Polices

    If(!$Error){

    Write-Host "[$ENV:COMPUTERNAME] : Configuration Manager Actions ran successfully." -ForegroundColor Green

    }

    Else{

    Write-Host "[$ENV:COMPUTERNAME] : Something went wrong, actions may not have ran properly :(" -ForegroundColor Yellow

    }

}

#Function for Remote
Function RefreshSCCM-Remote {
    
    Write-Host "Please select a .txt file containing a list of computer names"

    $path = Get-FileName -initialDirectory “c:fso"

    If ($path -eq "$null"){

        Write-Host "File not selected, terminating operation." -ForegroundColor Magenta
     
        exit
    }

    #Path to .txt with computer names
    $targets = Get-Content -Path $path

    #Loop to run configuration manager policies

    foreach ($target in $targets){
        
        $ping = Test-Connection -ComputerName $target -Count 1 -Quiet -ErrorAction SilentlyContinue

        If($ping){
        
            Try{

                Invoke-Command -ScriptBlock $Polices -ComputerName $target -ErrorAction Stop
                Write-Host "[$Target] : Configuration Manager Actions ran successfully." -ForegroundColor Green

            }

            Catch{


                Write-Host "[$Target] : Oh no, could not execute remote commands :(" -ForegroundColor Red

            } 
           
        }
         
        Else{

            Write-Host "[$Target] : Connection failed." -ForegroundColor Red
            Continue

        }          

    }

    Write-Host "Operation complete, have a great day!" -ForegroundColor Cyan  

}


#Finally, run the thing...
Invoke-Menu

#Good bye

