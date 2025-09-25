#Define what you want to ping
$subnet = "xxx.xxx.xxx"
$range = 1..254
#Execute Ping sweep, successful pings will resolve DNS Name and write results to screen
Foreach ($number in $range) {
   $ip = "$subnet.$number"
   $ping = Test-Connection $ip -Count 1 -Quiet
   if ($ping -eq 'true') {
       $name = Resolve-DnsName $ip -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Namehost
       Write-Host "$ip is occupied $name" -ForegroundColor Yellow
  }
   Else {
       Write-Host "$ip is available" -ForegroundColor Green
 }  
#Export results to csv
   $outpath = "C:\Temp\Results\SweepResults\PingOutput.csv"
   Add-Content -LiteralPath $outpath -Value "$ip,$name,$ping"            
}


