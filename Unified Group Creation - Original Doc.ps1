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


#Generate a mail nickname by removing spaces and special characters
#This is required to create a similar default email naming convention as was used in "Exchange Online" (I.e. Not graph)
function generateUniqueName {
    
    param (
        [string]$GroupName
    )

    $mailNickname = $GroupName -replace '[^a-zA-Z0-9]', ''
    $uniqueMailNickname = $mailNickname
    $nameCount = 1

    #Increments digit if name already taken
    while (Get-MgGroup -Filter "mailNickname eq '$uniqueMailNickname'" -ErrorAction SilentlyContinue) {
        $uniqueMailNickname = "$mailNickname$nameCount"
        $nameCount++
    }

    return $uniqueMailNickname;
}

function getUniqueInstructorId {

    param (
        [string]$nameOfInstructor
    )
        $user = Get-MgUser -Filter "mail eq '$nameOfInstructor'" -errorAction silentlycontinue
        return $user.Id
    
}

#This function just fetches all the groups or teams, depending on if "group" or "team" is entered as a parameter
#Looks complicated, but it just writes some output, gets the lists, and sends whichever list required back as an array
function getAllGroupsOrTeams {

    param(
        [string]$groupType
    )

    write-host "Collecting"$groupType"s, this will take some time... (5-10 minutes)"

    $teamsOrGroups = $null
    [array]$existingList = @()

    if($groupType -eq "group") {
        $teamsOrGroups = Get-MgGroup -all
    }
    elseif($groupType -eq "team"){
        $teamsOrGroups = Get-MgTeam -all
    }

    [array]$existingList = $teamsOrGroups | Select-Object -ExpandProperty DisplayName
    
    write-host -ForegroundColor Green "All"$groupType"s collected!  Total "$groupType"s: " $existingList.count;

    return $existingList

}

function getGroupId {
    param (
        [string]$groupName
    )

    $GroupObject = Get-MgGroup -Filter "DisplayName eq '$GroupName'" -ErrorAction SilentlyContinue
    return $groupObject.Id
}

function outputLineDataToConsole {
    param (
        [string]$instructor,
        [string]$groupName,
        [string]$GroupId,
        [string]$Student,
        [string]$severity
    ) 
    write-host "    Instructor: "$instructor -foregroundcolor $severity  
    write-host "    Group ID: "$groupId -foregroundColor $severity
    Write-host "    GroupName: "$GroupName -ForegroundColor $severity
    Write-host "    StudentName: "$Student -ForegroundColor $severity
}

#Helps avoid Script Running issues
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser;   #If this scripeed is resaved on this PC, will run without issue.
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted #Avoids Asking permission to download files from PSGallery

# Check if the Microsoft Graph module is already installed
if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
    # Module not installed, so install it
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Files
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Teams
Import-Module Microsoft.Graph.Groups

$scopes = @("User.Read.All", "Files.read.All", "Group.Read.All", "Group.ReadWrite.All", "Team.ReadBasic.All, Sites.Read.All")  #Microsoft 365 Permissions 

Connect-MgGraph -Scopes $scopes -NoWelcome      #Connect to Microsoft Graph with the provided permissions

#To get the site ID, you type in the search bar, "siteURL + /_api/web/id
#Then find the id inside the brackets
#So for the IT ops site, it would be https://nlc3.sharepoint.com/sites/it-ops-infra/_api/web/id
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
[array]$StudentGroups = @();
[array]$CreatedGroups = @();
[array]$CreatedTeams = @();
$skippedInstructorsCount = 0;
$skippedCoursesCount = 0;
$lastcheckeddate=0;
$errorList = @();                #This list provides a list of all items that generated a severe error


write-host "CSV File:" $CSVFilePath;
#We created this temporary file, but we don't need it anymore
Remove-Item -Path $tempFilePath   

# Retrieve all Microsoft 365 groups and select their display names (uses Microsoft Graph

[array]$existingGroups = getAllGroupsOrTeams -groupType "group"
[array]$existingTeams = getAllGroupsOrTeams -groupType "team"

#Stores error outputs to help keep code from cluttering functions.  These can be accessed before running access methods
#to modify azure resources
 $errorOutputs = @{}

$errorOutputs["1"] = @{

"Full" = "A teacher has not been assigned as instructor field Labelled as `'TBA`'.  Group could not be created."
"Short" = "Teacher name = 'TBA'."
}
$errorOutputs["2"] = @{
    "Full" = "Instructor email exists, but is disabled.  Group could not be created."
    "Short" = "Instructor account is Disabled."
}

$errorOutputs["3"] = @{
    "Full" = "Invalid Instructor email address.  Group could not be created."
    "Short" = "Instructor Email not Found in Azure."
}
$errorOutputs["4"] = @{
    "Full" = "Unknown Error During Account Creation"
    "Short" = "Unknown Error During Account Creation"
}
$errorOutputs["5"] = @{
    "Full" = "Attempted to add a group member to group, but duplicate groups exist with same name"
    "Short" = "Duplicate Group(s) Exist"
}
$errorOutputs["6"] = @{
    "Full" = "Attempted to remove a group member from group, but duplicate groups exist with same name"
    "Short" = "Duplicate Group(s) Exist"
}
$errorOutputs["7"] = @{
    "Full" = "Uknown Error happened during removal of student"
    "Short" = "Duplicate Group(s) Exist"
}


# --- Let's process the actual CSV file now ---
foreach ($Line in $List) {

    #Some internal declarations based on the specific line we're working on
    $GroupName = $Line.Section.trim();
    $Student = $Line.SEmail.trim();
    $Instructor = $Line.Instructor.trim();
    $Description = $Line.Description.trim();
    $Status = $Line.Status.trim();
        

    #Calls Declared functoin to get unique Id for the instructor who will be the owner of the group 
    $userId = getUniqueInstructorId -nameOfInstructor $instructor

    #Calls Declared Function to generate a unique name
    $uniqueMailNickname = generateUniqueName -GroupName $GroupName

    #####STORE ERROR CONDITIONS####
    
    [array]$errorConditionsCaught = $null

    if($userId -eq $null -and $instructor -eq "TBA") {
        $errorConditionsCaught += 1;           

    }
    elseif($userId -eq $null -and $instructor -ne "TBA"){
        $azAccount = Get-MgUser -Filter "userPrincipalName eq '$instructor'" -Property "displayName,accountEnabled" -ErrorAction SilentlyContinue

        #Catches if the account exists but is disabled...
        if($azAccount.AccountEnabled -eq $false){
            $errorConditionsCaught += 2;          
        }
        else{
            $errorConditionsCaught += 3;   
        }
    }

    #Skip lines without Instructors listed
    if ($Instructor -eq "TBA") {
        $skippedInstructorsCount++;
    }

    #catch for sections that explicitly do not want teams made
    if ($SkippedCourses -match $GroupName) {
        $skippedCoursesCount++;
    }
   
    #Check if the unified group exists already, create it if not
    if (($CreatedGroups -match $GroupName) -OR ($ExistingGroups -match $GroupName)) { 
        $CreatedGroups += ,$GroupName;
    }
 
    elseif($errorConditionsCaught.Count -gt 0){#Leave empty
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
        try{

            #Creates the new group using the $NewGroupSettings
            $group = New-MgGroup -BodyParameter $NewGroupSettings -ErrorAction Stop
            
            #Gets GroupId from new group just created.
            $groupId = $group.Id
            write-host "New Group created with name: "$groupName" and group id "$groupID 
            $CreatedGroups += ,$GroupName;
            
        }
        catch{
                $errorConditionsCaught += 4;  
        }

        #WE CAN TEST TO SEE IF THIS IS NECESSARY
        Start-Sleep -Seconds 10;
    }


    $GroupId = getGroupId -groupName $groupName

    if($groupId -ne $null) {

        $studentObject = Get-MgUser -Filter "userPrincipalName eq '$Student'"
        $ObjectId = $studentObject.Id


        #Add the student from the group as appropriate
        #It is significantly faster to re-perform this operation than it is to compare whether it needs to be done
        if ($Status -eq "InClass") {
            try{
                #Will need to add proper student as user
                New-MgGroupMember -GroupId $groupId -DirectoryObjectId $objectId -ErrorAction SilentlyContinue;
            }
            catch [System.Management.Automation.ParameterBindingException] {
                $errorConditionsCaught += 5;    
            }
        }

        #Remove the student from the group first, in case the group doesn't exist already it won't be created
        if ($Status -eq "NotInClass") {
            try {
                Remove-MgGroupMemberDirectoryObjectByRef -GroupId $groupId -DirectoryObjectId $objectId -ErrorAction Stop
            }
            catch [System.Management.Automation.ParameterBindingException] {
                $errorConditionsCaught += 6;  
            }
            catch {
                if($_.Exception.Message -like "*does not exist or one of its queried reference-property objects are not present*"){
                    if($groupId -eq $null){
                        Write-host "A student was failed to remove from a group, but the group doesn't exist yet." -Foreground yellow

                    }else{
                        write-host "Cannot remove student from group because student not found in group." -ForegroundColor yellow
                    }
                
                    outputLineDataToConsole -groupName $GroupName -GroupId $groupId -Student $student -instructor $instructor -severity "yellow"
                }
                else {
                    $errorConditionsCaught += 7;  
                }  

            
            }
        }
    }

    
    elseif($errorConditionsCaught.Count -gt 0){
        for($x = 0; $x -lt $errorConditionsCaught.Count; $x++){
            try {
                $key = $errorConditionsCaught[$x].toString()
                Write-Host  $errorOutputs[$key]["Full"] -ForegroundColor Red
                outputLineDataToConsole -groupName $GroupName -GroupId $groupId  -Student $Student -instructor $instructor -severity "red"
                Add-Member -InputObject $Line -MemberType NoteProperty -Name "Error Type" -Value $errorOutputs[$key].Short   
                $errorList += $Line
                
            }
            #Catches bad hash table values 
            catch {}
        }
    }

}

#This sleep adds additional 20 seconds between running groups actions and teams actions.  This ensures that all group
#actions are completed, avoiding Teams still provisioning e
Start-Sleep -Seconds 20;

foreach($Line in $List) {
 
        $GroupName = $Line.Section.trim(); 
 
        #Check if the associated team exists already, create it if not
        if (($CreatedTeams -match $GroupName) -OR ($ExistingTeams -match $GroupName)) { 
            #write-host $GroupName "Team already exists.";
            $CreatedTeams += ,$GroupName;
        }
        else {
            write-host "Creating Team" $GroupName "and waiting for 10 seconds";


            $GroupId = getGroupId -groupName $GroupName



            $params = @{
	            "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"    #Creates Standard Group Template
                "group@odata.bind" = "https://graph.microsoft.com/v1.0/groups('$GroupId')"               #Uses Existing Unified Group To create linked team.  Includes group ID, Visibility, DisplayName, Description
            }

            Start-Sleep -Seconds 10;

            $newTeam = New-MgTeam  -BodyParameter $params   
            write-host "New Team created with name: "$GroupName "and TeamId of" $groupId
            $CreatedTeams += ,$GroupName;
        
            $GroupObject = "";
        }
    


}



write-host "The script has completed.";
write-host "The total number of skipped lines was $($skippedInstructorsCount + $skippedCoursesCount).";
write-host "The total number of skipped TBA lines was $skippedInstructorsCount.";
write-host "The total number of lines skipped due to Course title was $skippedCoursesCount.";
write-host "If any errors appeared during the script run please either screenshot the entire error message including the lines before and after and provide the screenshot to Dan."  ;




$errorList

#Gets Class or Section and makes a list
$errorBoolean =  $errorList | Select-Object -ExpandProperty Section

$sectionChecked = @() #used to check if section was checked already

foreach($Line in $List) {



    #Checks if group shows up in error List.  If it does, we can't change the name
    if($errorBoolean.trim() -match $Line.Section.trim()){
        $sectionChecked += $Line.Section.Trim();
    }
    else {
                #Checks if group has already been checked
        if($sectionChecked -match $Line.Section.Trim()) {

        } 
        else {
            $misMatchFound = $false

            #Checks each line against the instructor.  If a mismatch is found, we won't process.
            foreach($innerLine in $List) {
                
                $misMatchFound = $false

                if($Line.Section.Trim() -eq $innerLine.Section.Trim()){ 

                    if($Line.Instructor.Trim() -eq $innerLine.Instructor.trim()){
                         write-host "MISMATCH IS NOT TRUE"
                    }
                    else {
                        write-host "MISMATCH FOUND"
                        write-host "InnerLine: " $innerLine.Section $innerLine.Instructor
                        write-Host "OuterLine: " $Line.Section $line.Instructor
                         $misMatchFound = $true
                        break;
                    }

                   
                }

            }

            if($misMatchFound -eq $false){
                Write-host "CHANGE OWNER OF GROUP"
                Write-Host $Line.Section
                #Check if group exists...
                #Check against value in azure

                if (getGroupId -groupName $Line.Section) {
                    Write-Host "Group exists. Line 495"


                    $groupId = getGroupId -groupName $Line.Section

                    $user = Get-MgUser -Filter "userPrincipalName eq '$($Line.Instructor)'"
                    $userId = $User.Id

                    #Get owner object from group
                    $owners = Get-MgGroupOwner -GroupId $groupId
                    $ownersArray = @()

                    #Puts all owners into an array
                    foreach($owner in $owners){
                        $ownerUPN = Get-MgUser -UserId $owner.Id
                        $ownersArray += $ownerUPN.UserPrincipalName



                    }

                    if($ownersArray -notmatch $Line.Instructor){
                        New-MgGroupOwnerByRef -GroupId $groupId -BodyParameter @{ "@odata.id"="https://graph.microsoft.com/v1.0/users/$userId" }
                
                        foreach($owner in $ownersArray) {
                            if($owner -notmatch $Line.Instructor){
                            
                                $userToRemove = Get-MgUser -Filter "userPrincipalName eq '$owner'"
                                $userIdToRemove = $userToRemove.Id
                                $userIdToRemove
                                Remove-MgGroupOwnerByRef -GroupId $groupId -DirectoryObjectId $userIdToRemove

                            }
                        }
                    }
                    

                    
                }
                


            }

            $sectionChecked += $Line.Section.Trim();
        }
    }
}
