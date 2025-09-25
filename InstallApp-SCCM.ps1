$ComputerName = $env:COMPUTERNAME
$AppName = "APPNAME"

foreach ($Computer in $ComputerName){

    $session = New-PSSession -ComputerName $ComputerName

    Invoke-Command -Session $session -ArgumentList $ComputerName,$AppName -ScriptBlock {

        $Application = Get-WmiObject -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -Class CCM_Application | where {$_.Name -like $AppName} | Select-Object Id, Revision, IsMachineTarget   
        $AppID = $Application.Id
        $AppRev = $Application.Revision
        $AppTarget = $Application.IsMachineTarget

        ([wmiclass]'ROOT\ccm\ClientSdk:CCM_Application').Install($AppID, $AppRev, $AppTarget, 0, 'Normal', $false) | Out-Null

    }

    Remove-PSSession $session

}


