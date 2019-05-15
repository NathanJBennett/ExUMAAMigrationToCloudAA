# Script
<# .SYNOPSIS
     Importing Auto Attendants that have been extracted using the ExportAA.PS1 script 
.DESCRIPTION
     Takes all setting from Exchange UM JSON file and attempts to recreate Cloud Auto Attendants with most of the settings configured in the Exchange UM Auto Attendants
.NOTES
     Author     : Nathan Bennett - nabennet@microsoft.com

#>

#Global Variables
$path = [IO.File]::ReadAllText($home + "\UMAAs.json")
$print = $true
$UMAAsJson = ConvertFrom-Json -InputObject $path
$online = ""
ipmo .\Schedules.psm1
$importPrompt = ""
$mapNoMenus = ""
$index = 1
$connectedOnline = Get-PSSession | ?{$_.ComputerName -like "*online.lync.com" -and $_.State -eq "Opened"}

$welcomeBanner =
@'
*********************************************************
*********************************************************
**  Welcome to the Cloud Auto Attendant Import Script  **
**    Please make sure you have read the Read Me       **
*********************************************************
*********************************************************
'@


write-Host $welcomeBanner -ForegroundColor Green -BackgroundColor Black

foreach ($UMAAJson in $UMAAsJson)
{
    Add-Member -InputObject $UMAAJson -MemberType NoteProperty -Name Index -Value $index
    $index++
}

[PSCustomObject[]]$UMAAs = @()

:quit

while ( 1 )
{
    if ($UMAAs.count)
    {
        Write-Host -ForegroundColor Green "Here are the Auto Attendants you've selected:"

        $UMAAs |
        FT Name, LineURI, Status |
        Out-String |
        Write-Host -ForegroundColor Cyan        
    }

    Write-Host -ForegroundColor Green "These are your Exchange UM Auto Attendants"
    $UMAAsJson |
    Format-Table -Property Index, Name, LineURI, Status |
    Out-String |
    Write-Host -ForegroundColor Cyan

    $importPrompt = Read-Host -Prompt "Please select the Auto Attendants to import.`nType 'All' to select all `nType 'Range' to select a range of Auto Attendants `nType the Index Number (Example '2') to select a single Auto Attendant `nType 'Quit' to exit"

    switch -Regex ($importPrompt)

    {
    '^All$' {
        $UMAAs += $UMAAsJSON
        break quit
        }

        '^Range$'{
            do {
                $start = Read-Host -Prompt "Starting Auto Attendant index #"
            } while ( !( $Start -as [int] ) -or ( $start -lt 1 ) )

            $end = Read-Host -Prompt "Ending Auto Attendant index #"

            if ( !$end) {
                $end = $start
            }

            $UMAAs += $UMAAsJSON[ ($start -1 ) .. ( $end - 1 ) ] |
            Sort-Object
            
        }

        '^\d+$' {
            if ( [int]$_ -le $UMAAsJSON.count )
            {
            $UMAAs += $UMAAsJSON[ $_ - 1 ]
            }
        }

        'quit' {
            break quit
        }
        default { write-host -ForegroundColor Red "Invalid entry"}

    }

    $index = 1
    $UMAAsJson = $UMAAsJson |
    ? Name -NotIn $UMAAs.Name |
    ForEach-Object -Process `
    {
        $_.index = $index
        $index++
        $_
    }

    if (!$UMAAsJson)
    {
        break
    }


}


# Asks user if they want to import the prompts which were formatted
while ($importPrompt -ne "yes" -and $importPrompt -ne "no")
    {
    $importPrompt = Read-Host "Would you like to import your audio prompts to your you Auto Attendants? If you did not choose to have your prompts converted please select 'No' here as this will fail if prompts are in an incorrect format. [Yes|No]"
    switch($importPrompt)
        {
            yes{"We will import audio files"}
            no{"We will not import audio files"}
            default {"Invalid entry"}
        }
    }

$menuInstructions = 
 @' 
 We will now look at the menus you have created in Exchange Online Unified Messaging. Cloud Auto Attendant Menus work differently. Cloud AA Menus have 3 options: Disconnect Call, Transfer to Operator, and Transfer to a Target. Cloud Auto attendant does not use Extensions, you may transfer to a user's Sip Address, Call Queue, or another Auto Attendant. The Sip Address assigned to your Extension was exported to your JSON file (if it was available). If you choose to to have this script build your menus, it will only assign the Menu Options that has an action to transfer to an Extension. We will transfer to the Extension's sip address. We will also upload your custom menu options prompts (this may need to be changed if you have menu options that are not available in Cloud Auto Attendant).
 
 If you choose to skip building menus we will build each auto attendant to automatically disconnect. 
 You can then go configure these menu options after all of your Auto Attendants have been imported.
'@

Write-Host $menuInstructions -ForegroundColor Yellow -BackgroundColor Black

# Asks user if they would like their menus moved over
while ($mapNoMenus -ne "yes" -and $mapNoMenus -ne "no")
    {
    $mapNoMenus = Read-Host "Would you like us to skip building your menus? [Yes|No]"
    switch($mapNoMenus)
        {
            yes{Write-Host -ForegroundColor Green -BackgroundColor Black "`nWe will skip building your menus."}
            no{Write-Host -ForegroundColor Green -BackgroundColor Black "`nWe will build your menus"}
            default {Write-Host -ForegroundColor red -BackgroundColor Black "`nInvalid entry"}
        }
    }

# Checks to see if the user needs to be connected to SFB Online Powershell
if (!$connectedOnline)
{
    while ($online -ne "yes" -and $online -ne "no")
        {
            $online = Read-Host "Are you Migrating your Exchange UM Atuo Attendants to Cloud Auto Attendants? Select Yes if you need to connect to Skype For Business Online Powershell to begin importing your Auto Attendants. Select No if you are already connected to SFB Online Powershelll and do not need to be logged into SFB Online PowerShell. [Yes|No]"
            switch($importPrompt)
                {
                    yes{Write-Host -ForegroundColor Green -BackgroundColor Black "`nWe will connect to Skype For Business Online Powershell"}
                    no{Write-Host -ForegroundColor Green -BackgroundColor Black "`nWe will not connect you to Skype For Business Online Powershell"}
                    default {Write-Host -ForegroundColor red -BackgroundColor Black "`nInvalid entry"}
                }
        }
}

# Logs user into SFBO powershell if they selected Yes in the previous question.
If ($online -eq "yes")
    {
        Import-Module SkypeOnlineConnector
        $userCredential = Get-Credential
        $sfbSession = New-CsOnlineSession -Credential $userCredential
        Import-PSSession $sfbSession 
    }

$connectedOnline = Get-PSSession | ?{$_.ComputerName -like "*online.lync.com" -and $_.State -eq "Opened"}

if (!$connectedOnline)
     {
        Write-Host "You are not connected to SFBO Powershell, this script will now exit. Please choose to connect to SFBO powerhsell next time you run this script." -ForegroundColor Red -BackgroundColor Black 
        Start-Sleep -s 15
        Exit
    }



# Main loop for Auto Attendant Menu settings
foreach ($UMAA in $UMAAs)
        {
            # Sets Each AAs main variables
            Write-Host "Creating" $UMAA.Name -ForegroundColor Green -BackgroundColor Black
            $aaname = $UMAA.Name
            $defaultmenuname = $UMAA.Name + " Menu"
            $afterHoursMenuName = $UMAA.Name + " After hours Menu"
            $defaultcallflowname = $UMAA.Name + " CallFlow"
            $afterhourscallflowname = $UMAA.Name + " AfterHours CallFlow"
            $operator = $UMAA.OperatorSipAddress

           

        # If user chose to have their menu's mapped this will get key mappings that have Extensions assigned and map them in Cloud Auto Attendant // Need to fix flow for menus
        if ($mapNoMenus -eq "no")

            {
            $mapTheseKeys = @()

            foreach ($mapping in $UMAA.convertedBusinessHoursKeyMapping)
                {
                if (($mapping.Extension -ne "") -and ($mapping.ExtensionSipAddress -ne ""))
                    {
                    $extObjectId = (Get-CsOnlineUser $mapping.ExtensionSipAddress).ObjectId
                    $extEntity = New-CsAutoAttendantCallableEntity -Identity $extObjectId -Type User
                    [hashtable]$menuOptionParameters =
                    @{
                        Action = "TransferCallToTarget"
                        DtmfResponse = "Tone"+$mapping.Key
                        CAllTarget = $extEntity
                        VoiceResponses = $mapping.Description
                     }
   
                # Creates all key mappings
                $mapTheseKeys += New-CsAutoAttendantMenuOption @menuOptionParameters
          
                  }
                  }
                if (($umaa.BusinessHoursTransferToOperatorEnabled) -and ($umaa.OperatorSipAddress -ne ""))
                    {
                        $transferToOperator = new-CsAutoAttendantMenuOption -Action "TransferCallToOperator" -DtmfResponse "Tone0"
                        $mapTheseKeys += $transferToOperator
                    }
                

            Write-Host "Creating" $defaultmenuname -ForegroundColor Green -BackgroundColor Black

       
            [hashtable]$menuParameters =
            @{
                Name = $defaultmenuname 

            }
                 # Creates prompts // Need to fix this
            if ($UMAA.BusinessHoursMainMenuCustomPromptEnabled)
                {
                    Write-Host "Adding your Business hours custom prompt audio file." -BackgroundColor Black -ForegroundColor Green
                    $name = $UMAA.Name
                    $promptname = "converted" + $UMAA.BusinessHoursMainMenuCustomPromptFilename
                    $content = Get-Content ".\AAPrompts\$name\BusinessHoursMainMenuCustomPromptFile\$promptname" -Encoding byte -ReadCount 0
                    $audioFile = Import-CsOnlineAudioFile -ApplicationId "OrgAutoAttendant" -FileName $promptname -Content $content
                    $BusinessHoursMainMenuCustomPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
                }
            Else
                {
                    $BusinessHoursMainMenuCustomPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Welcome to Your AutoAttendant"
                }

            if ($mapTheseKeys.Count -gt 0)
                {
                    $menuParameters.MenuOptions = $mapTheseKeys
                    $menuParameters.Prompts = $BusinessHoursMainMenuCustomPrompt
                }

            else 
                {
                    $menuoption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
                    $menuParameters.MenuOptions = $menuoption
                }


                # Need to fix this
            if (($UMAA.CallSomeoneEnabled) -and ($mapTheseKeys.Count -gt 0))
             {
                $menuParameters.EnableDialByName = $true
             }

            $menu = New-CsAutoAttendantMenu @menuParameters

            
            }
        Else
            {
                Write-Host "Creating" $defaultmenuname -ForegroundColor Green -BackgroundColor Black
                $menuoption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
                $menu = New-CsAutoAttendantMenu -Name $defaultmenuname -MenuOptions $menuoption
            }

        Write-Host "Creating" $defaultcallflowname -ForegroundColor Green -BackgroundColor Black

        [hashtable]$defaultCallFlowParameters =
                @{
        Name = $defaultcallflowname
        Menu = $menu
                 }
        # Uploads prompts if they are enabled in UM and user chose to upload prompts
        if (($UMAA.BusinessHoursWelcomeGreetingEnabled) -and ($importPrompt -eq "yes"))
            {
                Write-Host "Adding your Business hours audio file." -BackgroundColor Black -ForegroundColor Green
                $name = $UMAA.Name
                $promptname = "converted" + $UMAA.BusinessHoursWelcomeGreetingFilename
                $content = Get-Content ".\AAPrompts\$name\BusinessHoursWelcomeGreetingFile\$promptname" -Encoding byte -ReadCount 0
                $audioFile = Import-CsOnlineAudioFile -ApplicationId "OrgAutoAttendant" -FileName $promptname -Content $content
                $businesshoursHoursGreetingaudioFilePrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                $defaultCallFlowParameters.Greetings = $businesshoursHoursGreetingaudioFilePrompt

            }

        # Creates call Flow
        $callflow = New-CsAutoAttendantCallFlow @defaultCallFlowParameters

        [hashtable]$parameters =
        @{
            Name            = $aaName
            LanguageId      = $UMAA.Language
            DefaultCallFlow = $callFlow
            TimeZoneId      = $UMAA.TimeZone
        }

        $afterhourscallflows = @()
        $afterhoursCallHandlingAssociations = @()

        [string]$BusinessScheduleString = $UMAA.BusinessHoursSchedule

        # If Auto attendant isn't running a 24/7 scheulde this creates the after hours Menus, prompts, schedules, holidays. Need to make sure Holidays get transferred if they have a 24/7 schedule
        if ($BusinessScheduleString -ne "Sun.12:00 AM-Sat.11:45 PM Sat.11:45 PM-Sun.12:00 AM" )
            {
            if ($mapNoMenus -eq "no")
                {
             $mapAfterHourTheseKeys = @()

             foreach ($afterHourssMapping in $UMAA.convertedAfterHoursKeyMapping)
                {
                if (($afterHourssMapping.Extension -ne "") -and ($afterHourssMapping.ExtensionSipAddress -ne ""))
                    {
                        $extObjectId = (Get-CsOnlineUser $afterHourssMapping.ExtensionSipAddress).ObjectId
                        $extEntity = New-CsAutoAttendantCallableEntity -Identity $extObjectId -Type User
                        [hashtable]$AftermenuOptionParameters =
                        @{
                            Action = "TransferCallToTarget"
                            DtmfResponse = "Tone"+$afterHourssMapping.Key
                            CAllTarget = $extEntity
                            VoiceResponses = $afterHourssMapping.Description
                         }

                    }
                        # Creates All After Hours key mappings
                        $mapAfterHourTheseKeys += New-CsAutoAttendantMenuOption @AftermenuOptionParameters

                    }
                    if ($umaa.AfterHoursTransferToOperatorEnabled)
                        {
                            $afterHoursTransferToOperator = new-CsAutoAttendantMenuOption -action "TransferCallToOperator" -DtmfResponse "Tone0"
                            $mapAfterHourTheseKeys += $afterHoursTransferToOperator
                        }
                
            Write-Host "Creating" $afterHoursMenuName -ForegroundColor Green -BackgroundColor Black

            [hashtable]$afterHoursMenuParameters =
            @{
                Name = $afterHoursMenuName 

            }

            if ($UMAA.AfterHoursMainMenuCustomPromptEnabled)
                {
                    Write-Host "Adding your After hours custom prompt audio file." -BackgroundColor Black -ForegroundColor Green
                    $name = $UMAA.Name
                    $promptname = "converted" + $UMAA.AfterHoursMainMenuCustomPromptFilename
                    $content = Get-Content ".\AAPrompts\$name\AfterHoursWelcomeGreetingFile\$promptname" -Encoding byte -ReadCount 0
                    $audioFile = Import-CsOnlineAudioFile -ApplicationId "OrgAutoAttendant" -FileName $promptname -Content $content
                    $AfterHoursMainMenuCustomPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
                }
            Else
                {
                    $AfterHoursMainMenuCustomPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Welcome to Your AutoAttendant"
                }

             if ($mapAfterHourTheseKeys.Count -gt 0)
                {
                    $afterHoursMenuParameters.MenuOptions = $mapAfterHourTheseKeys
                    $afterHoursMenuParameters.Prompts = $AfterHoursMainMenuCustomPrompt
                }
            else 
                {
                    $menuoption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
                    $afterHoursMenuParameters.MenuOptions = $menuoption
                }


            if (($UMAA.CallSomeoneEnabled) -and ($mapAfterHourTheseKeys.Count -gt 0))
                 {
                    $afterHoursMenuParameters.EnableDialByName = $true
                 }

            $menu = New-CsAutoAttendantMenu @afterHoursMenuParameters

            }
        Else
            {
                Write-Host "Creating" $afterHoursMenuName -ForegroundColor Green -BackgroundColor Black
                $menuoption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
                $menu = New-CsAutoAttendantMenu -Name $afterHoursMenuName -MenuOptions $menuoption
            }
    


            Write-Host "Creating" $afterhourscallflowname -ForegroundColor Green -BackgroundColor Black

            [hashtable]$afterhoursCallFlowParameters =
            @{
                Name = $afterhourscallflowname
                Menu = $menu
            }

            if (($UMAA.AfterHoursWelcomeGreetingEnabled) -and ($importPrompt -eq "yes"))
                {
                    Write-Host "Adding your After hours audio file." -BackgroundColor Black -ForegroundColor Green
                    $name = $UMAA.Name
                    $promptname = "converted" + $UMAA.AfterHoursWelcomeGreetingFilename
                    $content = Get-Content ".\AAPrompts\$name\AfterHoursWelcomeGreetingFile\$promptname" -Encoding byte -ReadCount 0
                    $audioFile = Import-CsOnlineAudioFile -ApplicationId "OrgAutoAttendant" -FileName $promptname -Content $content
                    $afterhoursHoursGreetingaudioFilePrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                    $afterhoursCallFlowParameters.Greetings = $afterhoursHoursGreetingaudioFilePrompt

                }


            $afterhourscallflow = New-CsAutoAttendantCallFlow @afterhoursCallFlowParameters

            $afterhoursschedules = get-schedule($UMAA.BusinessHoursSchedule)
            $afterhoursCallHandlingAssociation = new-csAutoAttendantCallHandlingAssociation -Type Afterhours -scheduleID $afterhoursschedules.Id -CallFlowId $afterhourscallflow.Id

            $afterhourscallflows += $afterhourscallflow
            $afterhoursCallHandlingAssociations += $afterhoursCallHandlingAssociation

            $parameters.CallFlows               = $afterhourscallflows
            $parameters.CallhandlingAssociation = $afterhoursCallHandlingAssociations

     }
            

        if ($UMAA.HolidaySchedule -ne "")
            {
            foreach ($h in $UMAA.HolidaySchedule)
                {
                $holidayName = get-HolidayName($h)
                Write-Host "Creating" $holidayName "Schedule" -ForegroundColor Green -BackgroundColor Black
                
                $holidayDateRange = Get-HolidayDateRange($h)
                $holidaySchedule = New-CsOnlineSchedule -Name $holidayName -FixedSchedule -DateTimeRanges $holidayDateRange
                $name = $UMAA.Name
                Write-Host "Uploading" $holidayName "Prompt" -ForegroundColor Green -BackgroundColor Black
               
                $holidayPromptName = "converted" + (Get-HolidayPrompt($h))
                $holidaycontent = Get-Content ".\AAPrompts\$name\Holidays\$holidayName\$holidayPromptName" -Encoding byte -ReadCount 0
                
                $holidayAudioFile = Import-CsOnlineAudioFile -ApplicationId "OrgAutoAttendant" -FileName $holidayPromptName -Content $holidaycontent
                $holidayPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $holidayAudioFile
                
               $menuoption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
                $menu = New-CsAutoAttendantMenu -Name $holidayname -MenuOptions $menuoption
                $holidayMenuCallFlow = New-CsAutoAttendantCallFlow -Name $holidayName -Menu $menu -Greetings $holidayPrompt
                $holidayCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $holidaySchedule.Id -CallFlowId $holidayMenuCallFlow.Id

                $afterhourscallflows += $holidayMenuCallFlow
                $afterhoursCallHandlingAssociations += $holidayCallHandlingAssociation


                }
            $parameters.CallFlows               = $afterhourscallflows
            $parameters.CallhandlingAssociation = $afterhoursCallHandlingAssociations

            }

            $TTSSupported = @()
            $TTSLanguages  = Get-CsAutoAttendantSupportedLanguage|? {$_.VoiceResponseSupported -eq $true}
            foreach($TTSLanguage in $TTSLanguages)
            {
                if ($UMAA.Language -eq $TTSLanguage.Id)
                    {
                    $TTSSupported += $TTSLanguage
                    }
            }


        if ($UMAA.SpeechEnabled -and $TTSSupported.count)
            {
                Write-Host "Enabling Voice Response" -ForegroundColor Green -BackgroundColor Black
                $parameters.EnableVoiceResponse = $true
            }

         if ($UMAA.OperatorSipAddress -ne "")
            {
                Write-Host "Adding operator" $UMAA.OperatorSipAddress -ForegroundColor Green -BackgroundColor Black
                $operatorObjectId = (Get-CsOnlineUser $UMAA.OperatorSipAddress).ObjectId
                $operatorEntity = New-CsAutoAttendantCallableEntity -Identity $operatorObjectId -Type User
                $parameters.Operator = $operatorEntity 
            }



        New-CsAutoAttendant @parameters

        $aaname + " has been created"
        }



$endInstructions = 
@'
Your Auto Attendants have been created. Please go to the Post Setup Instructions in the Read Me document to finish configuring your Auto Attendants and begin testing them.
'@

Write-Host $endInstructions -ForegroundColor Yellow -BackgroundColor Black