

$PortData=Get-NetTCPConnection | Select-Object -Property *,@{'Name' = 'ProcessName';'Expression'={(Get-Process -Id $_.OwningProcess).Name}}

$FinalData=@()

$count=0;

foreach ($portinfo in $PortData)
{

$count++;

If ($count -gt 40) 
    {

      Write-Host "Just pausing a minute to avoid IP blocking from ip-api.com"
      Start-Sleep 70
      $count = 0
    }


if (($portinfo.RemotePort -ne 0) -and ($portinfo.RemoteAddress -ne '127.0.0.1'))

{

$remoteIp=$portinfo.RemoteAddress  

$ipinfo =Invoke-RestMethod -Method Get -Uri "http://ip-api.com/json/$remoteIp"

$FinalData+=[PSCustomObject]@{

    ProcessName=$portinfo.ProcessName
    InstanceID= $portinfo.InstanceID
    LocalPort=$portinfo.LocalPort
    RemotePort=$portinfo.RemotePort
    IP      = $ipinfo.query
    City    = $ipinfo.city
    Country = $ipinfo.country
    Isp     = $ipinfo.isp
  }
  
}

}


$FinalData|Out-GridView