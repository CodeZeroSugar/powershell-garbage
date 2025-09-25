#Select Active Directory Group you wish to modify
    $ADGroup = "AD_GROUP_HERE"

#Command to Add Member to AD Group
    $Comps = Get-Content -Path "C:\Temp\ADAdd.txt"
    foreach ($Comp in $Comps)
        {
        $Member = (Get-ADComputer $Comp).DistinguishedName
        Add-ADGroupMember -Identity $ADGroup -Members $Member -Confirm:$false
        }
