$subscriptions = Get-ChildItem -Path HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions | select Name

foreach ($sub in $subscriptions){
    $sub = $sub."Name".split('\')[7]
    $eventsources = Get-childItem -Path HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions\$sub\EventSources | select name
    foreach ($source in $eventsources){
        $source = ($source."Name" -split '\\',2)[1]
        $regkey = Get-ItemProperty -Path HKLM:$source
        foreach ($reg in $regkey){
            $LastHeartBeatTime = $reg.LastHeartBeatTime
            $date = [DateTime]::FromFileTime($LastHeartBeatTime)
            $today = Get-Date
            $timediff = New-Timespan -Start $date -End $today
            if ($timediff.Days -gt 30){
                $wefclient = $reg.PSChildName
                write-host "$wefclient has not checked in 30 days."
                write-host "Removing $wefclient from $sub subscription.`n"
                Remove-Item $reg.PSPath -Force -Recurse
                }
            }
        }
    }
