Function Get-HolidayDateRange{

param($holiday)

$dateGroups = [regex]::Match("$holiday", "(.*),(.*),(.*),(.*)")
$startDate = $dateGroups.Groups[3].value
$endDate = $dateGroups.Groups[4].value

$regexNewStartDate = [regex]::Match("$startDate", "(.*)\/(.*)\/(.*)")
$regexNewEndDate = [regex]::Match("$endDate", "(.*)\/(.*)\/(.*)")

$convertedStartDate = $regexNewStartDate.Groups[2].Value + "/" + $regexNewStartDate.Groups[1].Value + "/" + $regexNewStartDate.Groups[3].Value
$convertedEndDate = $regexNewEndDate.Groups[2].Value + "/" + $regexNewEndDate.Groups[1].Value + "/" + $regexNewEndDate.Groups[3].Value


[hashtable]$holidayDateRange =
@{
    Start = $convertedStartDate
}

if ($convertedStartDate -ne $convertedEndDate)
{
$holidayDateRange.End = $newendDay
}

New-CsOnlineDateTimeRange @holidayDateRange

}

#Extract Holiday Name
Function Get-HolidayName{

param($holiday)

    [string]$holidayName = $holiday -replace ',.*'
    

$holidayName

}

#Extract Holiday Prompt
Function Get-HolidayPrompt{

param($holiday)
    
    $holidayGroups = [regex]::Match("$holiday", "(.*),(.*),(.*),(.*)")
    $holidayPrompt = $holidayGroups.Groups[2].value
    

$holidayPrompt

}

#Get After Hours Schedule
Function Get-Schedule{

param($data)


[string[]]$daysOfWeek = @(
    'Sun',
    'Mon',
    'Tue',
    'Wed',
    'Thu',
    'Fri',
    'Sat'
)


<#
        $tr1 = New-CsOnlineTimeRange -Start 09:00 -End 12:00
        $tr2 = New-CsOnlineTimeRange -Start 13:00 -End 17:00
        $afterHours = New-CsOnlineSchedule -Name " After Hours" -WeeklyRecurrentSchedule -MondayHours @($tr1, $tr2) -TuesdayHours @($tr1, $tr2) -WednesdayHours @($tr1, $tr2) -ThursdayHours @($tr1, $tr2) -FridayHours @($tr1, $tr2) -Complement
#>

[hashtable]$timeRange = 
@{
    Sun = New-Object -TypeName Collections.ArrayList
    Mon = New-Object -TypeName Collections.ArrayList
    Tue = New-Object -TypeName Collections.ArrayList
    Wed = New-Object -TypeName Collections.ArrayList
    Thu = New-Object -TypeName Collections.ArrayList
    Fri = New-Object -TypeName Collections.ArrayList
    Sat = New-Object -TypeName Collections.ArrayList
}

$data |
ForEach-Object -Process `
{
    [string]$string = $_
    [string]$start = $string -replace '-.*'
    [string]$end   = $string -replace '.*-'
    
    [string]$startDay = $start -replace '\..*'
    [string]$endDay   = $end   -replace '\..*'

    if ( $startDay -eq $endDay )
    {
        [string[]]$days = $startDay
    }
	# if work hours exapnd into multiple days
    else
    {
        [int]$startDayIndex = $daysOfWeek.IndexOf( $startDay )
        [int]$endDayIndex   = $daysOfWeek.IndexOf( $endDay   )

        [string[]]$days = 
        $startDayIndex .. $endDayIndex |
        ForEach-Object -Process `
        {
            $daysOfWeek[$_]
        }
    }
        
    $days |
    ForEach-Object -Process `
    {
        [string]$thisDay = $_
        if ( $thisDay -eq $startDay )
        {
            $null = $start -match '.*\.(?<hour>\d\d?):(?<minute>\d\d) (?<AmPm>[AP]M)'
        
            if ( $Matches.Hour -ne '12' -and $Matches.AmPm -eq 'PM' )
            {
                [string]$startTime = '{0}:{1}' -f (
                    [int]( $Matches.hour ) + 12
                ), $Matches.minute
            }
            else
            {
                [string]$startTime = '{0}:{1}' -f $Matches.hour, $Matches.minute
            }
        }
        else
        {
            [string]$startTime = '00:00'
        }
        
        if ( $thisDay -eq $endDay )
        {
            [bool]$fullDay = $false
            $null = $end -match '.*\.(?<hour>\d\d?):(?<minute>\d\d) (?<AmPm>[AP]M)'
        
            if ( $Matches.Hour -ne '12' -and $Matches.AmPm -eq 'PM' )
            {
                [string]$endTime = '{0}:{1}' -f (
                    [int]( $Matches.hour ) + 12
                ), $Matches.minute
            }
            else
            {
                [string]$endTime = '{0}:{1}' -f $Matches.hour, $Matches.minute
            }
        }
        else
        {
            [bool]$fullDay = $true
            [string]$endtime = '23:30'
        }
        
        # generate time range $workdayTimeRange = New-CsOnlineTimeRange -Start 09:00 -End 17:00
        #"'$thisDay', '$startTime', '$endTime'"
        $timeRange.$thisDay += New-CsOnlineTimeRange -Start $startTime -End $endTime
        
        if ( $fullDay )
        {
            # generate time range
            #"'$thisday', '$endTime', '00:00'"
            $timeRange.$thisDay += New-CsOnlineTimeRange -Start $endTime -End 1.00:00
        }
    }
}

[hashtable]$parameters =
@{
    Name                    = $data.Name + " Schedule"
    WeeklyRecurrentSchedule = $true
    Complement              = $true
}

$timeRange.Keys |
Foreach-Object -Process `
{
    if ( $timeRange.$_.Count )
    {
        $parameters.$_ = $timeRange.$_
    }
}

New-CsOnlineSchedule @parameters

}