#Bulk Update single day of all "after hours" call schedules

Write-Host "First we get connected to teams,  a popup will come up it sometimes takes awhile."
Import-Module MicrosoftTeams
Connect-MicrosoftTeams
Write-Host "Should now be connected"

Write-Host "Getting schedules this can take a long time.  Just let it do it's thing..."
$schedules = Get-CsOnlineSchedule | where {$_.Name -like "*after hours*"}
Write-Host "Done getting schedules."

Write-Host "----"

$targetday = Read-Host "Enter day of week"
$starttime = Read-Host "Enter Start Time (8AM is 08:00)"
$endtime = Read-Host "Enter End Time (5PM is 17:00)"
foreach($schedule in $schedules)
{
    $commandString = "";

    $daysoftheweek = "Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"
    foreach ($day in $daysoftheweek)
    {
        # Write-Host "current day=$($day)"
        
        if($day -ne $targetday){    
        # Write-Host "$day not equal target day"
            if($schedule.WeeklyRecurrentSchedule.($day+"Hours") -ne $null){
                # Write-Host "$day not equal null (ie. Saturday/Sunday)"
                $commandString = $commandString + "-$($day)Hours "+'$'+"$($day) "
                Set-Variable -Name $day -Value (New-CsOnlineTimeRange -Start ($schedule.WeeklyRecurrentSchedule.($day+"Hours").Start.ToString()) -End ($schedule.WeeklyRecurrentSchedule.($day+"Hours").End.ToString()))
            }  
        }
        else{
            # Write-Host "$day equal to target day"
            $commandString = $commandString + "-$($day)Hours "+'$'+"$($day) "
            Set-Variable -Name $day -Value (New-CsOnlineTimeRange -Start $starttime -End $endtime)
        }
    }
    
    $tempschedule = Invoke-Expression('New-CsOnlineSchedule -Name "'+($schedule.Name)+'" -WeeklyRecurrentSchedule '+$commandString)
    $schedule.WeeklyRecurrentSchedule = $tempschedule.WeeklyRecurrentSchedule
    
    # I've seen this fail, when it does it crashes the script and leaves trash behind.
    # You can run the script again but this trash is left till you manually delete it.
    # And it will slow this script down even more.
    Write-Host("")
    Write-Host("Removing tempschedule $($tempschedule.Id)")
    Remove-CsOnlineSchedule -Id $tempschedule.Id
    Write-Host("")

    Set-CsOnlineSchedule -Instance $schedule
    Write-Host("New Configuration for $($schedule.Name) with ID $($schedule.Id)")
    Write-Host ($schedule.WeeklyRecurrentSchedule | Format-List -Force | Out-String)
}
