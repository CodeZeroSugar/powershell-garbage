#Sign-Scripts.ps1
#A1C Velder, Ian
#Created: February 20, 2025
<#Updates:
    
#>    
#References: SrA Brooks, Seth

#Run as Admin

#Set path to your scripts folder<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
$scriptPath = "C:\Temp\scripts"
$testpath = Test-Path -Path $scriptPath
If($testpath){
    Write-Host "File path is valid! :)" -ForegroundColor Green
}

Else{
    Write-Host "Invalid file path :(" -ForegroundColor Yellow
    Return
}


Set-Location -Path $scriptPath

#Get EDIPI

$userID = get-itemproperty LastLoggedOnUser -path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI'
$userID2 = $userID.LastLoggedOnUser
$EDIPI = $userID2.Substring(7)

#Check for existing cert
$check = Get-ChildItem Cert:\CurrentUser\My\ | Where-Object{$_.Subject -like "*$EDIPI.adf self-signed certificate"}

#Create a new cert if none
If($Check -eq $null){
    Write-Host "No certificate detected, creating new Self-Signed Certificate..." -ForegroundColor Yellow

    Try{
        $Cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "$EDIPI.adf self-signed certificate" -CertStoreLocation Cert:\CurrentUser\My -ErrorAction Stop
    }

    Catch{
        Write-Host "Something went wrong, terminating operating..." -ForegroundColor Red
        Return
    }

}


#Use existing cert
Else{
    Write-Host "Existing Self-Signed Certificate detected!" -ForegroundColor Green
    $cert = $check[0]
}


#Create list of scripts from provided file path
$scripts = (Get-ChildItem -Path $scriptPath -Filter *.ps1).Name

foreach ($script in $scripts){

    Set-AuthenticodeSignature -FilePath $script -Certificate $Cert

}


#Operation complete!
Write-Host "Your scripts have been signed!" -ForegroundColor Green
