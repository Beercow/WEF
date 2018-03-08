#----------------------------------------------------------------------------------------------------
# Author: Peter Geelen
# e-mail: peter@ffwd2.me
# Web: blog.identityunderground.be
#
# Credits: 
# A great thanks to Ed Wilson, the Scripting guy for helping out with troubleshooting on the output and display formats.
# http://blogs.technet.com/b/heyscriptingguy/
#
#----------------------------------------------------------------------------------------------------

cls

#Do you want to enable transcript for logging all output to file?
$enableLogging = $TRUE
$ExportEnabled = $TRUE

# logging
# if you need detailed logging for troubleshooting this script, you can enable the transcript
# get the script location path and use it as default location for storing logs and results 
$log = $MyInvocation.MyCommand.Definition -replace 'ps1','log'
$resultPath = $PSScriptRoot + '\'
Push-Location $resultPath

if ($enableLogging) 
{
Start-Transcript -Path $log -ErrorAction SilentlyContinue
Write-Host "Logging enabled..."
Write-Host

Write-Host "Powershell version"
$PSVersionTable.PSVersion
$Host.Version
(Get-Host).Version
Write-host
}


# Initialisation
# Source: http://www.computerperformance.co.uk/powershell/powershell_get_winevent.htm
# As of 2104, there is a PowerShell bug in cultures such as "en-GB" or "en-DE", 
# which prevents the display of the properties: 'LevelDisplayName' and 'Message'; here is a work-around
$currentCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture

# Configuration
# selected Event log sources 
# limit the feedback only to the following logs
$EventLogList = 'ForwardedEvents'
#,'Forefront Identity Manager Management Agent'

#for FIM Health Check the Security log has less signification value and is skipped
#add it to the eventlog list when you need it
#,'Security'

# if the event logs are containing too much data, 
# powershell might not be able to load all due to memory limits

# return X week to collect logs
## $startdate = ((Get-Date).AddDays(-7))
## $startdate = ((Get-Date).AddDays(-14))

#collect x month of logs
## $startdate = ((Get-Date).AddMonths(-1))

#collect x years of logs
$startdate = ((Get-Date).AddYears(-1))

## Background info 
# We only need the non-informational event
# Verbose      5                                                                                                                                  
# Warning      3                                                                                                                                  
# Error        2                                                                                                                                  
# Critical     1                                                                                                                                  

# Informational and LogAlways is discarded

#Informational 4                                                                                                                                  
#LogAlways     0


Write-host "Event logs list" 
Write-host "---------------" 
$Eventloglist

#Displaying the log settings for the selected Event log list

Write-Host 
Write-Host "Event Log Properties"
Write-Host "--------------------"

$Count = 0
$Activity = "Checking log properties"
foreach ($eventlog in $EventLogList)
{
    $count += 1
    $pct = ($Count / $EventLogList.Count * 100)
    Write-Progress -Activity $Activity -Status $EventLog -PercentComplete $pct
    
    #get all evenlogs from the list and display their properties
    
    $export = Get-WinEvent -ListLog $eventlog | Select-Object  -Property LogName,LogFilePath,LogType,RecordCount,MaximumSizeInBytes,FileSize,LastWriteTime,LastAccessTime,IsLogFull,Logmode 
    #display
    $export 

    #prep export
    $exportfile = $resultPath + "_props"+$eventlog + ".csv.txt"
    if ($exportEnabled) 
    {
        $export | Export-Csv $exportfile -NoTypeInformation
    }

    #if you want all properties of the logs, replace the named attribute list with *
    #Get-WinEvent -ListLog $eventlog | Select-Object -Property *


}

$allEvents = New-Object System.Collections.Hashtable

$Count = 0
$Activity = "Checking log details"
foreach ($eventlog in $EventLogList)
{
    $Count += 1
    $pct = ($Count / $EventLogList.Count * 100)
    $status = $EventLog + " (" + $Count + "/" + $EventLogList.Count +")"
    Write-Progress -Activity $Activity -Status $status -PercentComplete $pct    
    write-host $count"." $eventlog

    [System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"
    #if ($eventlog -eq 'System'){[System.Threading.Thread]::CurrentThread.CurrentCulture = $currentCulture}
  
        
    # query the event log 
    # store the data in a hashtable, to avoid new queries
    $allEvents = $Null
    $allEvents = Get-WinEvent -FilterHashtable @{logname=$eventlog;StartTime=$startdate;level=0,1,2,3,4,5} -ErrorAction SilentlyContinue
    #DEBUG if ($eventlog -eq "Application") { $allevents}

    if ($allEvents.count -eq 0) 
    {
        $message = "No events for " + $eventlog + "log since "+ $startdate + "."
        Write-Host $message
        Write-Host
        # no data to process, skip processing for current loop
        Continue
    }

    # Events Per Second
    Write-Host
    Write-Host "Event Per Second"
    Write-Host "----------------"
    Write-Host

    # Get Events by Day
    $AvgEventsPerDay = $allEvents | Group-Object -Property {$_.TimeCreated.toshortdatestring()} -NoElement | Select @{N='Date'; E={$_.Name}},Count,@{N='EPS'; E={[math]::Round(($_.Count / 28800), 5)}} | Sort-Object {$_."Date" -as [datetime]} -Descending
    $AvgEventsPerSecond = ($AvgEventsPerDay | Measure-Object 'EPS' -Average).Average
    $AvgEventsPerSecond = [math]::Round($AvgEventsPerSecond, 5)

    # Display evnets per second/day
    Write-Host "Average EPS : $AvgEventsPerSecond"
    Write-Host
    Write-Host "Events Per Day"
    Write-host "--------------"
    $AvgEventsPerDay | Format-Table -AutoSize

    # Group by event type
    write-host
    write-host "Group by Event type" 
    write-host "-------------------"
    
    #display all events grouped by type and sorted by count
    $export = $allEvents | Group-Object -Property {$_.LevelDisplayName} -NoElement | Select @{N='Event Level'; E={$_.Name}},Count | Sort-Object Count -Descending  
    #display
    $export |  Format-Table -AutoSize   

    #prep export
    $exportfile = $resultPath + "_EventNameStats" + $eventlog + ".csv.txt"
    if ($exportEnabled) {$export | Select-Object -Property Count,Name |Export-Csv $exportfile}

    # Group by log source
    write-host
    write-host "Group by log source"
    write-host "-------------------"

    #display all events grouped by source and sorted by count
    $export = $allEvents | Group-Object -Property {$_.ProviderName} -NoElement| Select @{N='Source'; E={$_.Name}},Count | Sort-Object Count -Descending
    #display
    $export | Format-Table -AutoSize

    #prep export
    $exportfile = $resultPath + "_EventNameStats2" + $eventlog + ".csv.txt"
    if ($exportEnabled) {$export | Select-Object -Property Count,Name |Export-Csv $exportfile}

    #detailed statistics for non-information events
    write-host
    write-host "Statistics by Event ID"
    write-host "----------------------"

    #For events detailed reporting we're only interested in error events
    #not interested in informational events (level 0 and 4)
    $evtStats = $allEvents | where -Property level -Notin -Value 0,4 | Group-Object id | Select @{N='Event_ID'; E={$_.Name}},Count | Sort-Object Count -Descending 
    $allevents = $Null

    #display stats in table format
    $export = $evtStats | Select-Object Event_ID,Count  
    #display
    #$export | Format-Table -AutoSize
    $export    

    #prep export
    $exportfile = $resultPath + "_EventIDStats"+ $eventlog + ".csv.txt"
    if ($exportEnabled) {$export | Export-Csv $exportfile  -NoTypeInformation}

    # evtStats has number and ID attribute
    # other attributes must be added:
    #  - errortype name
    #  - Source
    #  - errortype name

    [System.Collections.ArrayList]$results = @() 

    # for each event id in the event statistics
    # display the most recent event
    $Activity = "Looking up last event occurrence..."

    $i= 0
    foreach ($item in $evtStats) 
    {
        $i += 1
        $pct = ($i / $evtStats.Count * 100)
        $eventID = $item.Name
        $status = "EventID: "+ $item.Event_ID
        Write-Progress -Activity $Activity -Status $status -PercentComplete $pct
    
        $customobj = "" | select Count,TimeCreated,ErrorID,ErrorType,Source,Message
        $customobj.Count = $item.Count
        $customobj.ErrorID = $item.Event_ID
            
        #get most recent event from the eventID
        $id = $item.Event_ID.ToInt32($Null)

        [System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"
        $lastevent = get-winevent -FilterHashtable @{LogName=$eventlog;Id=$id} -MaxEvents 1 -ErrorAction SilentlyContinue

        #depending on local settings, query might fail, if it fails reset to local culture 
        if ($lastevent.LevelDisplayName.Length -eq 0) 
        {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $currentCulture
            $lastevent = get-winevent -FilterHashtable @{LogName=$eventlog;Id=$id} -MaxEvents 1
        }

        $customobj.ErrorType = $lastevent.LevelDisplayName
        $customobj.Source = $lastevent.ProviderName
        $customobj.TimeCreated = $lastevent.TimeCreated
        $customobj.Message = $lastevent.Message

        
        #prep EventID export
        $exportfile = $resultPath + $eventlog +'_EventID_' + $customobj.ErrorID + ".csv.txt"
        if ($exportEnabled) 
        {
            $customobj | Export-Csv $exportfile -NoTypeInformation
        }

        $results += $customobj
    }

    #Latest even details per event
    write-host "Latest event details per event" 
    write-host "-----------------------------"  

    #display with format
    $results | Format-Table -AutoSize

    if ($exportEnabled) 
    {
        $exportfile = $resultPath + "_lastEvents_short_" + $eventlog + ".txt"
        $results| Format-Table -AutoSize | out-file $exportfile
        $exportfile = $resultPath + "_lastEvents_detail_" + $eventlog + ".txt"
        $results | out-file $exportfile
    }

}

write-host
write-host "Script Completed." 
write-host 

# disable transcripting
if ($enableLogging) 
{Stop-Transcript -WarningAction Ignore -ErrorAction SilentlyContinue}

Pop-Location 

