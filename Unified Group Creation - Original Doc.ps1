<#
This script is designed to take the csv export of the informer report "O365 Unified groups for Course-Sections" (created by Warren).  It should then:
 - generate Unified Groups based on these cross-linked section IDs
 - Create Teams directly linked to these unified groups
 - populate the teams with students as appropriate
 - remove students as appropriate

Some of the things this script can NOT currently accommodate is the ability to verify that the owner of a group is who is listed on each line

It is astronomically faster to blind remove or add students to groups, ignoring whether this action is required, as checking group membership is significantly slower.

I'd like to be able to create these teams as class teams instead of generic collaboration teams, but am unsure how to do that utilizing the groupID as provided by the new-unifiedgroup
#>

#Helps avoid Script Running issues
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser;   #If this scripeed is resaved on this PC, will run without issue.
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted #Avoids Asking permission to download files from PSGallery

# Check if the Microsoft Graph module is already installed
if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
    # Module not installed, so install it
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
}


######### THIS MIGHT BE TEMPORARY.  NOT SURE IF THESE NEEEDS TO BE REMOVED TO USE GRAPH
if (-not(Get-Module -Name MicrosoftTeams -ListAvailable)) {
    Install-Module -Name MicrosoftTeams -Scope CurrentUser -Force
} 

<#########THIS MIGHT BE TEMPORARY NOT SURE IF THESE NEED TO BE REMOVED OT USE GRAPH
if (-not(Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Install-Module -Name ExchangeOnlineManagement -Scope AllUsers -Force

} 
#>


Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Files
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Teams
Import-Module Microsoft.Graph.Groups

$scopes = @("User.Read.All", "Files.read.All", "Group.Read.All", "Group.ReadWrite.All", "Team.ReadBasic.All")  #Microsoft 365 Permissions 

Connect-MgGraph -Scopes $scopes -NoWelcome      #Connect to Microsoft Graph with the provided permissions

#To get the site ID, you type in the search bar, "siteURL + /_api/
#So for the IT ops site, it would be https://nlc3.sharepoint.com/sites/it-ops-infra/_api/
#press enter to search, and a file should download.  Within the file, search for <d:Id m:type=”Edm.Guid”> 
#Then, you will find the GUID, which is a long Id number....
#
$siteId = "098fefe0-1133-449e-82df-24013c732708"
$driveiD = (Get-MgSiteDrive -SiteId $siteId).Id    #Id of the sharepoint site drive. 
$rootFolder = Get-MgDriveRoot -DriveId $driveId    #Gets the Root folder.  Contains the targetFolderPath (next line) 
$targetFolderPath = "Automation/PowerAutomate/Informer_TeamsGroupScript_Outputs"   #this path contains our file.
$folderNames = $targetFolderPath -split '/'        #useful function to create array of subfolders based on provided string.

$currentfolderId = $rootfolder.Id

#Recursively searches for the desired file Folder using the supplied $targetFolderPath
foreach($folderName in $folderNames){

    $children = Get-MgDriveItemChild -DriveId $driveId -DriveItemId $currentFolderId
    $folder = $children | Where-Object { $_.Name -eq $folderName -and $_.Folder }

    if($folder) {

        $currentFolderId = $folder.Id
    }
    else {
        Write-Host "Subfolder '$folderName' not found."
        break
    }
}

#Gets the child files within the final subfolder of the targetFolderPath
$children = Get-MgDriveItemChild -DriveId $driveId -DriveItemId $currentFolderId

#Gets the newest file
$file = $children | Where-Object { $_.LastModifiedDateTime } | Sort-Object LastModifiedDateTime -Descending | Select-Object -First 1

#Creates a temporary file
$tempFilePath = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.csv'

#Outputs the newest file to the temporary file path
Get-MgDriveItemContent -DriveId $driveId -DriveItemId $file.Id -OutFile $tempFilePath | Out-Null

#Gets file content to use and manipulate as desired.  It's better to do it this way, as working directly from sharepoint is cumbersome (supposedly)
#This method allows us also to keep the original code in the Unified Group Script similar to the original.
$fileContent = Get-Content -Path $tempFilePath









# --- Generate the credentials to connect to O365 for easier script execution ---
if ($env:username -eq "dbuczek") {
    $username = 'dbuczek@nlc.bc.ca';
    $SecPass = '01000000d08c9ddf0115d1118c7a00c04fc297eb0100000034d8c811f6a15245a42e601d7c5910490000000002000000000003660000c000000010000000edbace923455e776bafbfcd202e83cfd0000000004800000a000000010000000e5354d39b965cea71cbf38dc315a991318000000b183ffbf0c590a2bbcb9cbc77010ebe6d110871901e0aad114000000f5710fd8eca912a96f76b993bfb23aaa28371950';
    $Password = ConvertTo-SecureString $SecPass;
    $Cred = New-Object -typename System.Management.Automation.PSCredential -argumentlist $Username,$Password;
}
else {
    $Cred = get-credential;
}












# --- Connect to Exchange Online Powershell and Microsoft Teams ---

#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection;
#Import-PSSession $Session -ErrorAction SilentlyContinue -AllowClobber;

#Import-Module ExchangeOnlineManagement;
Import-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue;
#Connect-MicrosoftTeams -AccountId $username -ErrorAction SilentlyContinue;
Connect-MicrosoftTeams -Credential $Cred;
#Connect-ExchangeOnline -Credential $Cred;

# --- DECLARATIONS ---;


$CSVFilePath = $tempFilePath;
#Headers copied from CSV "REG_CODE","STUDENT","STATUS","FACULTY_EMAIL","DESCRIPTION"
$CSVHeaders = 'Section','SEmail','Status','Instructor','Description';

#DC-

#This Skipped Sections document will be used to check if groups already exist.  I think originally this was for if teachers dont want groups created 
#Ed says we don't maintain a list like that anymore, so i think that makes this file pointless.  Thinking to just get rid of it.
#(will silence errors for groups that we manually skip or that might already exist)
$SkippedCourses = Get-Content "\\dc-casht\C$\dbuczek_working_documents\O365 Groups from D2L\Skipped Sections.txt";
$List = Import-CSV $CSVFilePath -Header $CSVHeaders -Delimiter "," | select -skip 1;

$CSVFilePath


$GroupObject = "";
$Count = 0;
[array]$StudentGroups = @();
[array]$CreatedGroups = @();
[array]$CreatedTeams = @();
$skippedInstructorsCount = 0;
$skippedCoursesCount = 0;
$lastcheckeddate=0;

write-host "CSV File:" $CSVFilePath;
## Prompt boxes
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Refresh Groups and Teams";
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Do not Refresh Groups and Teams";
$cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Cancel this script run entirely";
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $cancel);

## Check whether internal Groups and Teams and Teams need to be refreshed
$title = "Refresh internal Groups/Teams list?" ;
$message = "Would you like to refresh the internal list of Groups and Teams?`nSelecting yes will require an additional 10-20 minutes at the beginning for the check.`nOnly select no if you have recently refreshed of Groups and Teams in this session.`nIf you do not understand this choice select Cancel.";
$result = $host.ui.PromptForChoice($title, $message, $options, 1);
switch ($result) {
    0{write-host "Collecting Groups, this will take some time... (5-10 minutes)";
        
        # Retrieve all Microsoft 365 groups and select their display names (uses Microsoft Graph
        $groups = Get-MgGroup -All
        [array]$existingGroups = $groups | Select-Object -ExpandProperty DisplayName


        write-host -ForegroundColor Green "Groups Collected!  Total groups:" $ExistingGroups.count;
        write-host "Collecting Teams, this will take some time... (5-10 minutes)";
        
        $ProgressPreference = "SilentlyContinue";


        # Retrieve all Microsoft Teams and select their display name
        $teams = Get-MgTeam
        [array]$ExistingTeams = $teams | Select-Object -ExpandProperty DisplayName

        write-host -ForegroundColor Green "Teams Collected!  Total teams:" $ExistingTeams.count;
        write-host "Beginning to process the CSV file.";
        $progressPreference = "Continue";}
    1{Write-Host -foregroundcolor Red "In Test Mode, not refreshing Groups or Teams!";}
    2{Write-Host "Cancel"; break;}
}



# --- Let's process the actual CSV file now ---
foreach ($Line in $List) {
    write-progress -Activity "Processing..." -CurrentOperation "Line $($list.IndexOf($line)) of $($List.Length)" -PercentComplete (($list.IndexOf($line)/$List.Length)*100);
    #Some internal declarations based on the specific line we're working on
    $GroupName = $Line.Section;
    $Student = $Line.SEmail;
    $Instructor = $Line.Instructor;
    $Description = $Line.Description;
    $Status = $Line.Status;
    
    #Get unique Id for the instructor who will be the owner of the group 
    $user = Get-MgUser -Filter "mail eq '$instructor'"
    $userId = $user.Id

    #####-----Email Unique Name Generator--------#####

    #Generate a mail nickname by removing spaces and special characters
    #This is required to create a similar default email naming convention as was used in "Exchange Online" (I.e. Not graph)
    $mailNickname = $GroupName -replace '[^a-zA-Z0-9]', ''
    $uniqueMailNickname = $mailNickname
    $count = 1
    
    #Increments digit if name already taken
    while (Get-MgGroup -Filter "mailNickname eq '$uniqueMailNickname'" -ErrorAction SilentlyContinue) {
        $uniqueMailNickname = "$mailNickname$count"
        $count++
    }

    #####-----End of Email name Generator ------######





    #Overhead for keeping track of what line we're on - debug
    $Count++;
    #write-host "Processing line #"$Count":"$Line;
    #if ($Count % 100 -eq 0) {
        #write-host "Currently processing line" $Count;
    #}

    #Skip lines without Instructors listed
    if ($Instructor -eq "TBA") {
        $skippedInstructorsCount++;
        continue;
    }

    #catch for sections that explicitly do not want teams made
    if ($SkippedCourses -match $GroupName) {
        $skippedCoursesCount++;
        continue;
    }
   
    #Remove the student from the group first, in case the group doesn't exist already it won't be created
    if ($Status -eq "NotInClass") {
        
        
        $GroupObject = Get-MgGroup -Filter "DisplayName eq '$GroupName'" -ErrorAction SilentlyContinue
        $GroupId = $groupObject.Id

        $user = Get-MgUser -Filter "userPrincipalName eq '$Student'"

        # Extract the Object ID
        $objectId = $user.Id

        #Will need to add proper student as user
        Remove-MgGroupMemberDirectoryObjectByRef -GroupId $groupId -DirectoryObjectId $objectId -ErrorAction SilentlyContinue;

        
        #THIS USES EXCHANGE TOO
        #Remove-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $Student -Confirm:$false -ErrorAction SilentlyContinue;
        continue;
    }

    #Check if the unified group exists already, create it if not
    if (($CreatedGroups -match $GroupName) -OR ($ExistingGroups -match $GroupName)) { 
        #write-host $GroupName "Group already exists."; 
        $CreatedGroups += ,$GroupName;
    }
    else {
        write-host "Creating Group" $GroupName "and waiting for 10 seconds";
        
        #Settings for the group to be created.
        $NewGroupSettings = @{ 
            "displayName" = $GroupName #works
            "mailEnabled" = "true" #works
            "mailNickname"= $uniqueMailNickname #works
            "securityEnabled" = "false" #works
            "groupTypes" = @("Unified") #works
            "description" = $Description #works
            "owners@odata.bind" = @("https://graph.microsoft.com/v1.0/users/$userId")  #works
            "resourceBehaviorOptions" = @("WelcomeEmailDisabled") #works
            "visibility" = "Private"
        }  

        #Creates the new group using the $NewGroupSettings
        $group = New-MgGroup -BodyParameter $NewGroupSettings 

        #Gets GroupId from new group just created.
        $groupId = $group.Id

        $CreatedGroups += ,$GroupName;

        #WE CAN TEST TO SEE IF THIS IS NECESSARY
        Start-Sleep -Seconds 10;
    }

    #Check if the associated team exists already, create it if not
    if (($CreatedTeams -match $GroupName) -OR ($ExistingTeams -match $GroupName)) { 
        #write-host $GroupName "Team already exists.";
        $CreatedTeams += ,$GroupName;
    }
    else {
        write-host "Creating Team" $GroupName "and waiting for 10 seconds";

        $GroupObject = Get-MgGroup -Filter "DisplayName eq '$GroupName'" -ErrorAction SilentlyContinue
        $GroupId = $groupObject.Id
        
        $params = @{
	        "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"    #Creates Standard Group Template
            "group@odata.bind" = "https://graph.microsoft.com/v1.0/groups('$GroupId')"               #Uses Existing Unified Group To create linked team.  Includes group ID, Visibility, DisplayName, Description
        }

        $newTeam = New-MgTeam  -BodyParameter $params   

        #$GroupObject = Get-UnifiedGroup $GroupName -ErrorAction SilentlyContinue
        #New-Team -GroupID $GroupObject.ExternalDirectoryObjectID -ErrorAction SilentlyContinue | out-null;
        $CreatedTeams += ,$GroupName;
        Start-Sleep -Seconds 10;
        $GroupObject = "";
    }

    #Add the student from the group as appropriate
    #It is significantly faster to re-perform this operation than it is to compare whether it needs to be done
    if ($Status -eq "InClass") {
        #write-host "Adding" $Student "to Group" $GroupName;
        
        $GroupObject = Get-MgGroup -Filter "DisplayName eq '$GroupName'" -ErrorAction SilentlyContinue
        $GroupId = $groupObject.Id

        $user = Get-MgUser -Filter "userPrincipalName eq '$Student'"

        # Extract the Object ID
        $objectId = $user.Id

        #Will need to add proper student as user
        New-MgGroupMember -GroupId $groupId -DirectoryObjectId $objectId -ErrorAction SilentlyContinue;


        #THIS LIKELY USES EXCHANGE TOO
        #Add-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $Student -ErrorAction SilentlyContinue;
    }    
}



write-host "The script has completed.";
write-host "The total number of skipped lines was $($skippedInstructorsCount + $skippedCoursesCount).";
write-host "The total number of skipped TBA lines was $skippedInstructorsCount.";
write-host "The total number of lines skipped due to Course title was $skippedCoursesCount.";
write-host "If any errors appeared during the script run please either screenshot the entire error message including the lines before and after and provide the screenshot to Dan."  ;

#We created this temporary file, but we don't need it anymore
Remove-Item -Path $tempFilePath
