# Script
<# .SYNOPSIS
     Exports Exchange UM Auto Attendants 
.DESCRIPTION
     Export Exchange UM auto Attendants and formats AAs menu options. Creates a full JSON file with all UM settings. Also downloads all UM AA prompts and will format them so they can be uploaded to SFBO to recreate Cloud Auto Attendants.
.NOTES
     Author     : Nathan Bennett - nabennet@microsoft.com
#>

#Global Variables
$Online = ""
$ffmpeg = ".\ffmpeg.exe"
$ffmpegInPlace = Test-Path $ffmpeg -PathType Leaf
$convert = ""
$path = $home + "\UMAAs.json"
$print = $false
$PowershellVersion = $PSVersionTable.PSVersion.Major -ge 6

$welcomeBanner =
@'
****************************************************************
****************************************************************
**  Welcome to the Exchange UM Auto Attendants Export Script  **
**        Please make sure you have read the Read Me          **
****************************************************************
****************************************************************
'@


write-Host $welcomeBanner -ForegroundColor Green -BackgroundColor Black

# Asks user if they would like to have their audio prompt files converted to be able to upload them to Skype For Business Online
while ($convert -ne "yes" -and $convert -ne "no")
    {
        $convert = Read-Host "We will download all of your audio prompts. In order for these prompts to be uploaded to your new Cloud Auto Attendants they need to be converted. If you already have run this script and downloaded/converted prompts this script will overwrite the current files. Would you like us to convert your audio prompts? [Yes|No]"
        switch($convert)
            {
                yes{write-host "`nWe will convert your audio files`n" -ForegroundColor Green -BackgroundColor Black}
                no{write-host "`nWe will not convert your audio files`n" -ForegroundColor Green -BackgroundColor Black}
                default {write-host "`nInvalid entry`n" -ForegroundColor red -BackgroundColor Black}
            }
    }


# If the user chooses to convert their audio files but do not have "ffmpeg" in the current directory this will fail, display an error message and exit after 15 seconds
if (!$ffmpegInPlace -and $convert -eq "yes")
    {
        Write-Host "ffmpeg.exe was not found in this directory, please move it here and rerun this script" -BackgroundColor Black -ForegroundColor Red
        Start-Sleep -s 15
    Exit
    }


# Checks to see if the user needs to be connected to Exchange Online Powershell
while ($Online -ne "yes" -and $Online -ne "no")
    {
        $Online = Read-Host "Do you need to connect to Exchange Online Powershell? Select Yes if you need to connect to Exchange Online Powershell to begin Exporting your Auto Attendants. Select No if you are already connected to Exchange Online Powershell or using Exchange management shell to export Exchange onPrem Auto Attendants and do not need to be logged into Exchange Online PowerShell. [Yes|No]"
        switch($Online)
            {
                yes{ write-host "`nWe will connect to O365 Exchange Online Powershell" -ForegroundColor Green -BackgroundColor Black}
                no{ write-host "`nWe will not connect you to O365 Exchange Online Powershell" -ForegroundColor Green -BackgroundColor Black}
                default {write-host "`nInvalid entry" -ForegroundColor Red -BackgroundColor Black}
            }
    }

# Connects to Exchange Online Powershell if the user answered yes to the previous question.
If ($Online -eq "YES")
    {
        $UserCredential = get-credential
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        Import-PSSession $Session -DisableNameChecking
    }


# Formatted UMAA class with additional fields for Cloud Voicemail.
Class ExportUMAA
{
   [String]$Name
   [String]$Language
   [String]$InfoAnnouncementEnabled
   [bool]$SpeechEnabled
   [bool]$AllowDialPlanSubscribers
   [bool]$AllowExtensions
   [String]$AllowedInCountryOrRegionGroups
   [String]$AllowedInternationalGroups
   [bool]$CallSomeoneEnabled
   [String]$ContactScope
   [String]$ContactAddressList
   [bool]$SendVoiceMsgEnabled
   [String[]]$BusinessHoursSchedule
   [String]$PilotIdentifierList
   [String]$UMDialPlan
   [String]$DTMFFallbackAutoAttendant
   [String[]]$HolidaySchedule
   [String]$TimeZone
   [String]$TimeZoneName
   [String]$BusinessLocation
   [String]$MatchedNameSelectionMethod
   [String]$WeekStartDay
   [String]$Status
   [String]$OperatorExtension
   [String]$InfoAnnouncementFilename
   [bool]$StarOutToDialPlanEnabled
   [bool]$NameLookupEnabled
   [bool]$ForwardCallsToDefaultMailbox
   [String]$DefaultMailbox
   [String]$BusinessName
   [String]$BusinessHoursWelcomeGreetingFilename
   [bool]$BusinessHoursWelcomeGreetingEnabled
   [String]$BusinessHoursMainMenuCustomPromptFilename
   [bool]$BusinessHoursMainMenuCustomPromptEnabled
   [bool]$BusinessHoursTransferToOperatorEnabled
   [String[]]$BusinessHoursKeyMapping
   [bool]$BusinessHoursKeyMappingEnabled
   [String]$AfterHoursWelcomeGreetingFilename
   [bool]$AfterHoursWelcomeGreetingEnabled
   [String]$AfterHoursMainMenuCustomPromptFilename
   [bool]$AfterHoursMainMenuCustomPromptEnabled
   [bool]$AfterHoursTransferToOperatorEnabled
   [String[]]$AfterHoursKeyMapping
   [bool]$AfterHoursKeyMappingEnabled
   [string]$OperatorSipAddress
   [array]$convertedBusinessHoursKeyMapping
   [array]$convertedAfterHoursKeyMapping

}

# Class created for formatted key mappings in each menu
Class KeyMapping
{
   [String]$Key
   [String]$Description
   [String]$Extension
   [string]$AutoAttendantName
   [string]$PromptFileName
   [string]$AsrPhrases
   [String]$LeaveVoicemailFor
   [String]$TransferToMailbox
   [string]$AnnounceBusinessLocation
   [String]$AnnounceBusinessHours
   [String]$ExtensionSipAddress
}


# Grabs all Auto Attendants
$UMAAs = get-umautoattendant


# Creates initial directory for AA Prompts
$AApromptpath = Test-Path .\AAPrompts
if (!$AApromptpath)
{
    New-Item -Name AAPrompts -ItemType "directory" -path .\ | out-null
}

# Array for each exported Auto Attendant
$ExportUMAAs = @()




# Main loop for Auto Attendant Menu settings
foreach ($UMAA in $UMAAs)
    {

    # Gets key mappings from Business hours Menu and creates formatted list
    if ($UMAA.BusinessHoursKeyMappingEnabled)
        {
        Write-Host "Extracting your Business Hours Menu from" $UMAA.Name "Auto Attendant" -BackgroundColor Black -ForegroundColor Green
        $bHKeyMappings = $UMAA.BusinessHoursKeyMapping
        $BusinessHoursKeyMapping = @()
            foreach ($keyMap in $bHKeyMappings)
            {
                $mappings = [regex]::Match("$keyMap", "(.*),(.*),(.*),(.*),(.*),(.*),(.*),(.*),(.*),(.*)")
                $key = $mappings.Groups[1].value
                $Description = $mappings.Groups[2].value
                $Extension = $mappings.Groups[3].value
                $AutoAttendantName = $mappings.Groups[4].value
                $PromptFileName = $mappings.Groups[5].value
                $AsrPhrases = $mappings.Groups[6].value
                $LeaveVoicemailFor = $mappings.Groups[7].value
                $TransferToMailbox = $mappings.Groups[8].value
                $AnnounceBusinessLocation = $mappings.Groups[9].value
                $AnnounceBusinessHours = $mappings.Groups[10].value
                $extSIP = ""

                # Get's SIP address of Menu options assigned to an Extension
                if ($Extension -ne "")
                    {
                    $ext = Get-UMMailbox | ? {$_.extensions -like $Extension}
                    $extSIP = $ext.SIPResourceIdentifier
                    }

                # Creates key mapping objects
                $keyMapping = new-Object KeyMapping
                $KeyMapping.key = $key
                $KeyMapping.Description = $Description
                $KeyMapping.Extension = $Extension
                $KeyMapping.AutoAttendantName = $AutoAttendantName
                $KeyMapping.PromptFileName = $PromptFileName
                $KeyMapping.AsrPhrases = $AsrPhrases
                $KeyMapping.LeaveVoicemailFor = $LeaveVoicemailFor
                $KeyMapping.TransferToMailbox = $TransferToMailbox
                $KeyMapping.AnnounceBusinessLocation = $AnnounceBusinessLocation
                $KeyMapping.AnnounceBusinessHours = $AnnounceBusinessHours
                $KeyMapping.ExtensionSipAddress = $extSIP

                $BusinessHoursKeyMapping += $keyMapping

            }

        }

    # Gets key mappings from After hours Menu and creates formatted list
    if ($UMAA.AfterHoursKeyMappingEnabled)
        {
        Write-Host "Extracting your After Hours Menu from" $UMAA.Name "Auto Attendant" -BackgroundColor Black -ForegroundColor Green
        #$AfterHoursKeyMappingJSON = $home + "\AfterHoursKeyMapping.json"
        $aHKeyMappings = $UMAA.AfterHoursKeyMapping
        $AfterHoursKeyMapping = @()
        foreach ($keyMap in $aHKeyMappings)
            {
                $mappings = [regex]::Match("$keyMap", "(.*),(.*),(.*),(.*),(.*),(.*),(.*),(.*),(.*),(.*)")
                $key = $mappings.Groups[1].value
                $Description = $mappings.Groups[2].value
                $Extension = $mappings.Groups[3].value
                $AutoAttendantName = $mappings.Groups[4].value
                $PromptFileName = $mappings.Groups[5].value
                $AsrPhrases = $mappings.Groups[6].value
                $LeaveVoicemailFor = $mappings.Groups[7].value
                $TransferToMailbox = $mappings.Groups[8].value
                $AnnounceBusinessLocation = $mappings.Groups[9].value
                $AnnounceBusinessHours = $mappings.Groups[10].value
                $extSIP = ""

                if ($Extension -ne "")
                    {
                    $ext = Get-UMMailbox | ? {$_.extensions -like $Extension}
                    $extSIP = $ext.SIPResourceIdentifier
                    }

                $keyMapping = new-Object KeyMapping
                $KeyMapping.key = $key
                $KeyMapping.Description = $Description
                $KeyMapping.Extension = $Extension
                $KeyMapping.AutoAttendantName = $AutoAttendantName
                $KeyMapping.PromptFileName = $PromptFileName
                $KeyMapping.AsrPhrases = $AsrPhrases
                $KeyMapping.LeaveVoicemailFor = $LeaveVoicemailFor
                $KeyMapping.TransferToMailbox = $TransferToMailbox
                $KeyMapping.AnnounceBusinessLocation = $AnnounceBusinessLocation
                $KeyMapping.AnnounceBusinessHours = $AnnounceBusinessHours
                $KeyMapping.ExtensionSipAddress = $extSIP

                $AfterHoursKeyMapping += $keyMapping

            }

        }

    Write-Host "Exporting" $UMAA.Name "Auto Attendant" -BackgroundColor Black -ForegroundColor Green

    # Checks Auto attendant to find an assigned operator and grabs operators SIP address
    $operatorSIP = ""
    if ($UMAA.OperatorExtension -ne "")
        {
        $operator = Get-UMMailbox | ? {$_.extensions -like $UMAA.OperatorExtension}
        $operatorSIP = $operator.SIPResourceIdentifier
        }

    # Creates Auto Attendant Objects into a master list that includes key mappings
    $ExportUMAA = New-Object ExportUMAA
    $ExportUMAA.Name = $UMAA.Name
    $ExportUMAA.Language  = $UMAA.Language 
    $ExportUMAA.InfoAnnouncementEnabled  = $UMAA.InfoAnnouncementEnabled 
    $ExportUMAA.SpeechEnabled = $UMAA.SpeechEnabled
    $ExportUMAA.AllowDialPlanSubscribers  = $UMAA.AllowDialPlanSubscribers 
    $ExportUMAA.AllowExtensions  = $UMAA.AllowExtensions 
    $ExportUMAA.AllowedInCountryOrRegionGroups = $UMAA.AllowedInCountryOrRegionGroups
    $ExportUMAA.AllowedInternationalGroups  = $UMAA.AllowedInternationalGroups 
    $ExportUMAA.CallSomeoneEnabled  = $UMAA.CallSomeoneEnabled 
    $ExportUMAA.ContactScope = $UMAA.ContactScope
    $ExportUMAA.ContactAddressList  = $UMAA.ContactAddressList 
    $ExportUMAA.SendVoiceMsgEnabled  = $UMAA.SendVoiceMsgEnabled 
    $ExportUMAA.BusinessHoursSchedule = $UMAA.BusinessHoursSchedule
    $ExportUMAA.PilotIdentifierList  = $UMAA.PilotIdentifierList 
    $ExportUMAA.UMDialPlan  = $UMAA.UMDialPlan 
    $ExportUMAA.DTMFFallbackAutoAttendant = $UMAA.DTMFFallbackAutoAttendant
    $ExportUMAA.HolidaySchedule  = $UMAA.HolidaySchedule 
    $ExportUMAA.TimeZone  = $UMAA.TimeZone 
    $ExportUMAA.TimeZoneName = $UMAA.TimeZoneName
    $ExportUMAA.BusinessLocation  = $UMAA.BusinessLocation 
    $ExportUMAA.MatchedNameSelectionMethod  = $UMAA.MatchedNameSelectionMethod 
    $ExportUMAA.WeekStartDay = $UMAA.WeekStartDay
    $ExportUMAA.Status  = $UMAA.Status 
    $ExportUMAA.OperatorExtension  = $UMAA.OperatorExtension 
    $ExportUMAA.InfoAnnouncementFilename = $UMAA.InfoAnnouncementFilename
    $ExportUMAA.StarOutToDialPlanEnabled  = $UMAA.StarOutToDialPlanEnabled 
    $ExportUMAA.NameLookupEnabled  = $UMAA.NameLookupEnabled 
    $ExportUMAA.ForwardCallsToDefaultMailbox = $UMAA.ForwardCallsToDefaultMailbox
    $ExportUMAA.DefaultMailbox  = $UMAA.DefaultMailbox 
    $ExportUMAA.BusinessName  = $UMAA.BusinessName 
    $ExportUMAA.BusinessHoursWelcomeGreetingFilename  = $UMAA.BusinessHoursWelcomeGreetingFilename 
    $ExportUMAA.BusinessHoursWelcomeGreetingEnabled = $UMAA.BusinessHoursWelcomeGreetingEnabled
    $ExportUMAA.BusinessHoursMainMenuCustomPromptFilename  = $UMAA.BusinessHoursMainMenuCustomPromptFilename 
    $ExportUMAA.BusinessHoursMainMenuCustomPromptEnabled  = $UMAA.BusinessHoursMainMenuCustomPromptEnabled 
    $ExportUMAA.BusinessHoursTransferToOperatorEnabled = $UMAA.BusinessHoursTransferToOperatorEnabled
    $ExportUMAA.BusinessHoursKeyMapping  = $UMAA.BusinessHoursKeyMapping 
    $ExportUMAA.BusinessHoursKeyMappingEnabled  = $UMAA.BusinessHoursKeyMappingEnabled 
    $ExportUMAA.AfterHoursWelcomeGreetingFilename = $UMAA.AfterHoursWelcomeGreetingFilename
    $ExportUMAA.AfterHoursWelcomeGreetingEnabled  = $UMAA.AfterHoursWelcomeGreetingEnabled 
    $ExportUMAA.AfterHoursMainMenuCustomPromptFilename  = $UMAA.AfterHoursMainMenuCustomPromptFilename 
    $ExportUMAA.AfterHoursMainMenuCustomPromptEnabled  = $UMAA.AfterHoursMainMenuCustomPromptEnabled 
    $ExportUMAA.AfterHoursTransferToOperatorEnabled = $UMAA.AfterHoursTransferToOperatorEnabled
    $ExportUMAA.AfterHoursKeyMapping  = $UMAA.AfterHoursKeyMapping 
    $ExportUMAA.AfterHoursKeyMappingEnabled  = $UMAA.AfterHoursKeyMappingEnabled 
    $ExportUMAA.OperatorSIPAddress = $operatorSIP
    if ($UMAA.BusinessHoursKeyMappingEnabled)
        {
        $ExportUMAA.convertedBusinessHoursKeyMapping = $BusinessHoursKeyMapping
        $ExportUMAA.convertedAfterHoursKeyMapping = $AfterHoursKeyMapping
        }

    $ExportUMAAs += $ExportUMAA

    #Creates directories for all Auto Attendant Prompts and converts them if that option was selected in the beginning.
    $name = $UMAA.Name
    $AAnamepath = Test-Path .\AAPrompts\$name
    if (!$AAnamepath)
        {
            New-Item -Name $name -ItemType "directory" -Path .\AAPrompts | out-null
        }

    Write-Host "Exporting" $name "Prompts" -ForegroundColor Green -BackgroundColor Black
    if ($UMAA.BusinessHoursWelcomeGreetingEnabled)
        {
            $AABusinessHoursWelcomeGreetingPath = Test-Path .\AAPrompts\$name\BusinessHoursWelcomeGreetingFile
            if (!$AABusinessHoursWelcomeGreetingPath)
                {
                    New-Item -Name BusinessHoursWelcomeGreetingFile -ItemType "directory" -Path .\AAPrompts\$name | out-null  
                }              
            $promptname = $UMAA.BusinessHoursWelcomeGreetingFilename
            if (!$PowershellVersion)
            {
            $prompt = Export-UMPrompt -PromptFileName $promptname -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\BusinessHoursWelcomeGreetingFile\$promptname -Value $prompt.AudioData -Encoding byte -Force
            }
            else
            {
                $prompt = Export-UMPrompt -PromptFileName $promptname -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\BusinessHoursWelcomeGreetingFile\$promptname -Value $prompt.AudioData -AsByteStream -Force
                    
            }
            if ($convert -eq "yes") 
                {
                    .\ffmpeg.exe -i .\AAPrompts\$name\BusinessHoursWelcomeGreetingFile\$promptname -b:a 256k -ar 16000  .\AAPrompts\$name\BusinessHoursWelcomeGreetingFile\converted$promptname -loglevel quiet -y
                }
                
        }

    if ($UMAA.AfterHoursWelcomeGreetingEnabled)
        {
        $AAAfterHoursWelcomeGreetingPath = Test-Path .\AAPrompts\$name\AfterHoursWelcomeGreetingFile
            if (!$AAAfterHoursWelcomeGreetingPath)
                {
                    New-Item -Name AfterHoursWelcomeGreetingFile -ItemType "directory" -Path .\AAPrompts\$name | out-null    
                }          
            $promptname = $UMAA.AfterHoursWelcomeGreetingFilename
                if (!$PowershellVersion)
                {
                    $prompt = Export-UMPrompt -PromptFileName $promptname -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\AfterHoursWelcomeGreetingFile\$promptname -Value $prompt.AudioData -Encoding byte -Force
                    }
                    else
                    {
                    $prompt = Export-UMPrompt -PromptFileName $promptname -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\AfterHoursWelcomeGreetingFile\$promptname -Value $prompt.AudioData -AsByteStream -Force
                    }
            if ($convert -eq "yes") 
                {
                    .\ffmpeg.exe -i .\AAPrompts\$name\AfterHoursWelcomeGreetingFile\$promptname -b:a 256k -ar 16000  .\AAPrompts\$name\AfterHoursWelcomeGreetingFile\converted$promptname -loglevel quiet -y
                }
                
        }

    if ($UMAA.InfoAnnouncementEnabled -ne "False")
        {
        $AAInfoAnnouncemenPath = Test-Path .\AAPrompts\$name\InfoAnnouncementFile
            if (!$AAInfoAnnouncemenPath)
                {
                    New-Item -Name InfoAnnouncementFile -ItemType "directory" -Path .\AAPrompts\$name | out-null   
                }            
            $promptname = $UMAA.InfoAnnouncementFilename
            if (!$PowershellVersion)
                {
                $prompt = Export-UMPrompt -PromptFileName $promptname -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\InfoAnnouncementFile\$promptname -Value $prompt.AudioData -Encoding byte -Force
                }
                else
                {
                    $prompt = Export-UMPrompt -PromptFileName $promptname -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\InfoAnnouncementFile\$promptname -Value $prompt.AudioData -AsByteStream -Force
                }
                
                
        }

    if ($UMAA.BusinessHoursMainMenuCustomPromptEnabled)
        {
        $AABusinessHoursMainMenuCustomPrompt = Test-Path .\AAPrompts\$name\BusinessHoursMainMenuCustomPromptFile
            if (!$AABusinessHoursMainMenuCustomPrompt)
                {
                    New-Item -Name BusinessHoursMainMenuCustomPromptFile -ItemType "directory" -Path .\AAPrompts\$name | out-null     
                }           
            $promptname = $UMAA.BusinessHoursMainMenuCustomPromptFilename
            if (!$PowershellVersion)
                {
                $prompt = Export-UMPrompt -PromptFileName $promptname -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\BusinessHoursMainMenuCustomPromptFile\$promptname -Value $prompt.AudioData -Encoding byte -Force
                }
                else
                {
                $prompt = Export-UMPrompt -PromptFileName $promptname -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\BusinessHoursMainMenuCustomPromptFile\$promptname -Value $prompt.AudioData -AsByteStream -Force
                }
            if ($convert -eq "yes") 
                {
                    .\ffmpeg.exe -i .\AAPrompts\$name\BusinessHoursMainMenuCustomPromptFile\$promptname -b:a 256k -ar 16000  .\AAPrompts\$name\BusinessHoursMainMenuCustomPromptFile\converted$promptname -loglevel quiet -y
                }
                
        }

    if ($UMAA.AfterHoursMainMenuCustomPromptEnabled)
        {
        $AAAfterHoursMainMenuCustomPrompt = Test-Path .\AAPrompts\$name\AfterHoursMainMenuCustomPromptFile
            if (!$AAAfterHoursMainMenuCustomPrompt)
                {
                    New-Item -Name AfterHoursMainMenuCustomPromptFile -ItemType "directory" -Path .\AAPrompts\$name | out-null        
                }                           
            $promptname = $UMAA.AfterHoursMainMenuCustomPromptFilename
                if (!$PowershellVersion)
                {
                $prompt = Export-UMPrompt -PromptFileName $promptname -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\AfterHoursMainMenuCustomPromptFile\$promptname -Value $prompt.AudioData -Encoding byte -Force
                }
                else
                {
                $prompt = Export-UMPrompt -PromptFileName $promptname -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\AfterHoursMainMenuCustomPromptFile\$promptname -Value $prompt.AudioData -AsByteStream -Force
                }
            if ($convert -eq "yes") 
                {
                    .\ffmpeg.exe -i .\AAPrompts\$name\AfterHoursMainMenuCustomPromptFile\$promptname -b:a 256k -ar 16000  .\AAPrompts\$name\AfterHoursMainMenuCustomPromptFile\converted$promptname -loglevel quiet -y
                }
                
        }
    if ($UMAA.HolidaySchedule -ne "")
        {
        $AAHolidaysPath = Test-Path .\AAPrompts\$name\Holidays
            if (!$AAHolidaysPath)
                {
                    New-Item -Name Holidays -ItemType "directory" -Path .\AAPrompts\$name | Out-Null
                }
        foreach ($h in $UMAA.HolidaySchedule)
            {
                $holidayGroups = [regex]::Match("$h", "(.*),(.*),(.*),(.*)")
                $holidayName = $holidayGroups.Groups[1].value
                $holidayAudioFile = $holidayGroups.Groups[2].value
                $AAeachHoliday = Test-Path .\AAPrompts\$name\Holidays\$holidayName
                if (!$AAeachHoliday)
                    {
                        New-Item -Name $holidayName -ItemType "directory" -Path .\AAPrompts\$name\Holidays | out-null
                    }
                if (!$PowershellVersion)
                {
                $prompt = Export-UMPrompt -PromptFileName $holidayAudioFile -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\Holidays\$holidayName\$holidayAudioFile -Value $prompt.AudioData -Encoding byte -Force 
                }
                else
                {
                $prompt = Export-UMPrompt -PromptFileName $holidayAudioFile -UMAutoAttendant $UMAA.Name; Set-Content -Path .\AAPrompts\$name\Holidays\$holidayName\$holidayAudioFile -Value $prompt.AudioData -AsByteStream -Force 
                }
                if ($convert -eq "yes") 
                {
                    .\ffmpeg.exe -i .\AAPrompts\$name\Holidays\$holidayName\$holidayAudioFile -b:a 256k -ar 16000  .\AAPrompts\$name\Holidays\$holidayName\converted$holidayAudioFile -loglevel quiet -y
                }
                    
            }

        }

    }

# Creates master JSON file
$ExportUMAAs | ConvertTo-Json -Depth 3 | Out-File $path 

if ($print){
get-content $path | Write-Host
}

Write-Host "Done! Result written to:" $path -BackgroundColor Black -ForegroundColor Green

