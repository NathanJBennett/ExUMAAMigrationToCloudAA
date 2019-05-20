<h1>Exchange UM Auto Attendant to Cloud Auto Attendant Migration Scripts</h1>

The Exchange UM AA to Cloud AA migration scripts have been created to assist Microsoft O365 tenant admins in migrating their Exchange UM Auto Attendant to the new Skype For Business Online/Teams Auto Attendant platform.


<h2>Getting started</h2>
•	Review the documentation to understand the concepts behind planning your Cloud Auto Attendants and requirements. (https://docs.microsoft.com/en-us/skypeforbusiness/hybrid/plan-cloud-auto-attendant)

•	Create a basic Cloud Auto Attendant through SFBO PowerShell examples.
(https://docs.microsoft.com/en-us/powershell/module/skype/new-csautoattendant?view=skype-ps)


<h2>Important Details</h2>

**Export Process**

[GitHub](http://github.com)
ExportUMAAs.ps1 will connect you to Exchange Online and get all of your configured Exchange Online UM Auto Attendants and export them into a JSON format in your $Home directory with a filename of: UMAAs.json. It also takes any Menu option that has a configured extension and gets the user’s SIP URI. 
It will also download all of your audio prompts and create directories for these in this format:
.\AAPrompts\(AA name)\(File Prompt)\
The script will ask the user if they want to have their audio files converted in order to be uploaded to Cloud Auto Attendant. This is necessary as Exchange UM imported prompts are in a MPEG 2 format and the requirements to upload an audio prompt to SFBO is outlined here: (https://docs.microsoft.com/en-us/powershell/module/skype/import-csonlineaudiofile?view=skype-ps). The script uses ffmpeg.exe to convert your audio files. This executable must be in the directory you are running the script from. You may download FFMpeg from: https://ffmpeg.org/download.html
Note: If you would like to connect to your onprem Exchange Server and download all of your onPrem Auto Attendants, use your onPrem Exchange management shell to run this script and select “No” when the prompt asks you if you would like to connect to Exchange online.

**Import Process**

ImportCloudAAs.PS1 will import your JSON file and allow you to select the Auto Attendants you would like to have imported to Cloud Auto Attendant. 
It will then ask if you converted your audio files in the Export process and would like those files to be uploaded to your new Cloud Auto Attendants.
You will then be prompted if you would like the script to build out your menus. The only menu options (as of now) that will be configured will be extensions you had configured in your Exchange UM Auto Attendants.
You’ll then be prompted to connect to SFBO Powershell if you are not already connected.
The script will then begin to configure each of your new Cloud Auto Attendants you selected and upload your audio prompts. If speech was enabled on your Exchange UM Auto Attendant but the language is not supported in Cloud Auto Attendant this script will disable voice response. While connected to SFBO Powershell you can find Cloud AA Languages that support Voice response by running:
 Get-CsAutoAttendantSupportedLanguage| ?{$_.VoiceResponseSupported -eq $true} 
Once the script finishes, follow the Post setup instructions to create your onprem disabled object and associate that with your new Cloud Auto Attendant so you can begin testing.
After this is done you can log into the Teams admin portal (https://admin.teams.microsoft.com/auto-attendants) and review your new Auto Attendants and their settings.
Note: the ImportCloudAAs.ps1 script needs to load the Schedules.psm1 module in order to run. Make sure this is in the directory you are running the script from.


<h2>Instructions</h2>

1.	Download all files into a working directory. 
2.	Unzip ffmeg.zip in the same directory or download ffmpeg from: https://ffmpeg.org/download.html Make sure that ffmpeg.exe is in your working directory.
3.	Open Powershell and navigate to your working directory
4.	Run .\ExportEXUMAA.ps1
5.	Follow onscreen instructions until the script completes.
6.	Run .\ImportCloudAAs.ps1
7.	Follow onscreen instructions until the script completes.
8.	Follow the Post setup Instructions below.


<h2>Post setup Instructions</h2>

Now that the basic framework of your Auto Attendants has been created its time to review the settings and begin testing your Cloud Auto Attendants. Go to the Teams admin portal to review your auto attendants and their settings.

In order to begin testing your Auto Attendants we will need to create a resource account and associate the Auto Attendant with that endpoint. For instructions for using a Direct Routing number or assigning a Microsoft service number please go here: (https://docs.microsoft.com/en-us/microsoftteams/create-a-phone-system-auto-attendant#step-1---get-started)

Many tenants migrating from Exchange UM AA to Cloud AA will want to use the same onprem number that is assigned to their Exchange UM AA. When using Exchange UM AA with SFB you may recall you setup disabled user object (New-csexumcontact), this is very similar to what is needed to Cloud Auto Attendant. You will create a HybridApplicationEndpoint (disabled User object) in SFB onPremise and then associate that with your Cloud Auto Attendant. First you will create the mapping without assigning the phone number so you can test by simply calling the SIP Address of the HybridApplicationEndpoint.
Run the following in SFB onpremise, you will need to provide the correct OU, you can also change the displayname and sip address if needed:
    
New-CsHybridApplicationEndpoint -SipAddress "(AutoAttendantName)@(your SIP domain)" -DisplayName "Auto Attendant Display Name" -ApplicationId ce933385-9390-45d1-9512-c8d228074e07 -OU (Desired OU)"

All Cloud Auto Attendants have the same Application ID of : ce933385-9390-45d1-9512-c8d228074e07
You will need to get the Hybrid Application Endpoints Ids, you can retrieve those by running the following: 

Get-CsHybridApplicationEndpoint | select DisplayName, Name

The Hybrid Application Endpoint ID will be the value of Name. Make sure this account gets dir sync'd online. From the server you have Azure Ad Connect you can run the following to force an immediate sync:

Start-ADSyncSyncCycle -PolicyType Delta

You will need to get the Auto Attendant Configuration ID, you can get this by running the following while connected to SFBO PowerShell: 

   	 Get-CsAutoAttendant | Select Name, ID

Lastly you need to create an ApplicationInstanceAssociation that links the Auto Attendant to the onprem object (HybridApplicationEndpoint). Run the following:

   	 New-CsOnlineApplicationInstanceAssociation -ConfigurationType AutoAttendant -ConfigurationId <AA configuration ID> -Identities <Identity of Hybrid Application Endpoint>
	
You can check this by running: 

    	Get-CsOnlineApplicationInstanceAssociation

You can then start testing your newly created Auto Attendants by calling the Hybrid Application Endpoint's sip address When you are ready to point the Exchange Online Auto Attendant number to your Cloud attendants simply update the Line URI on the HybridApplicationEndpoint and wait for this to be sync'd online.    Ex: Set-CsHybridApplicationEndpoint -LineUri +**********"
Note: The HybridApplicationEndpoint (resource account) will need to be licensed in order to assign a phone number.

You may need to remove the Line URI for the Exchange UM auto attendant. 
    Ex: Set-CsExUmContact -DisplayNumber ""

Remember if you have Nested Auto Attendants and you don't need to use one of your onPrem numbers you can create a OnlineApplicationInstance instead of an onPremise Disabled user object (HybridApplicationEndpoint). You will still need to create the OnlineApplicationInstanceAssociation for this. However, it is not necessary to license this account if no phone number is assigned.


<h2>Additional Notes</h2>

The following settings will not be moved over:
InfoAnnouncementEnabled,
InfoAnnouncementFilename,
-Info announcement is not an option in Cloud Auto Attendant, we will still pull down this prompt for you.
AllowDialPlanSubscribers,
AllowExtensions,
AllowedInCountryOrRegionGroups,
AllowedInternationalGroups,
-Cloud Auto Attendant does not have extensions and will only call out to phone numbers you set in the Menu Options
ContactScope,
ContactAddressList,
Cloud Auto Attendant has a similar option to this call Dial Scope. You can include or Exclude users in the organization’s directory when DialByName is configured in Cloud Auto Attendant. You can find more information about this here: https://docs.microsoft.com/en-us/powershell/module/skype/new-csautoattendantdialscope?view=skype-ps
You can pull all the user’s sip address in a dialplan by connecting to Exchange Online and running:
Get-UMMailbox | ?{$_.umdialplan -eq "(DialPlanName)"} | select SIPResourceIdentifier 

SendVoiceMsgEnabled – Not allowed in Cloud Auto Attendant,
DTMFFallbackAutoAttendant – Only options that can be used are DTMF or Enable voice inputs,
MatchedNameSelectionMethod – Not configurable, Cloud Auto Attendant allows you to say the First Name or Last Name of user,
BusinessLocation – Not configurable,
ForwardCallsToDefaultMailbox – Not configurable,
DefaultMailbox – Not configurable,
WeekStartDay – Not configurable,
StarOutToDialPlanEnabled – Not configurable,
Business Name - Not Configurable
