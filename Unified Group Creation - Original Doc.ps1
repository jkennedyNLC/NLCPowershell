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
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser;
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection;
#Import-PSSession $Session -ErrorAction SilentlyContinue -AllowClobber;
Install-Module -Name ExchangeOnlineManagement -Scope AllUsers
Import-Module ExchangeOnlineManagement;
Import-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue;
#Connect-MicrosoftTeams -AccountId $username -ErrorAction SilentlyContinue;
Connect-MicrosoftTeams -Credential $Cred;
Connect-ExchangeOnline -Credential $Cred;

# --- DECLARATIONS ---;
#$CSVFilePath = "\\dc-casht\C$\dbuczek_working_documents\O365 Groups from D2L\O365 Unified groups for Course-Sections TESTING.csv";
#$CSVFilePath = "\\dc-casht\C$\dbuczek_working_documents\O365 Groups from D2L\O365 Unified groups for Course-Sections.csv";
$CSVFilePath = Get-ChildItem "\\dc-casht\C$\dbuczek_working_documents\O365 Groups from D2L\Informer Exports" | sort LastWriteTime | select -last 1;
#Headers copied from CSV "REG_CODE","STUDENT","STATUS","FACULTY_EMAIL","DESCRIPTION"
$CSVHeaders = 'Section','SEmail','Status','Instructor','Description';
$SkippedCourses = Get-Content "\\dc-casht\C$\dbuczek_working_documents\O365 Groups from D2L\Skipped Sections.txt";
$List = Import-CSV $CSVFilePath.FullName -Header $CSVHeaders -Delimiter "," | select -skip 1;
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
        [array]$ExistingGroups = Get-UnifiedGroup -ResultSize Unlimited | select-object DisplayName;
        write-host -ForegroundColor Green "Groups Collected!  Total groups:" $ExistingGroups.count;
        write-host "Collecting Teams, this will take some time... (5-10 minutes)";
        $ProgressPreference = "SilentlyContinue";
        [array]$ExistingTeams = Get-Team | select-object DisplayName;
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
        Remove-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $Student -Confirm:$false -ErrorAction SilentlyContinue;
        continue;
    }

    #Check if the unified group exists already, create it if not
    if (($CreatedGroups -match $GroupName) -OR ($ExistingGroups -match $GroupName)) { 
        #write-host $GroupName "Group already exists."; 
        $CreatedGroups += ,$GroupName;
    }
    else {
        write-host "Creating Group" $GroupName "and waiting for 10 seconds";
        New-UnifiedGroup -DisplayName $GroupName -Owner $Instructor -Notes $Description -AccessType Private | Out-Null;
        $CreatedGroups += ,$GroupName;
        Set-UnifiedGroup -Identity $GroupName -UnifiedGroupWelcomeMessageEnabled:$false -AcceptMessagesOnlyFromSendersOrMembers $Instructor -HiddenFromAddressListsEnabled:$true -AlwaysSubscribeMembersToCalendarEvents:$TRUE -SubscriptionEnabled:$TRUE -AutoSubscribeNewMembers:$TRUE;
        Start-Sleep -Seconds 10;
    }

    #Check if the associated team exists already, create it if not
    if (($CreatedTeams -match $GroupName) -OR ($ExistingTeams -match $GroupName)) { 
        #write-host $GroupName "Team already exists.";
        $CreatedTeams += ,$GroupName;
    }
    else {
        write-host "Creating Team" $GroupName "and waiting for 10 seconds";
        $GroupObject = Get-UnifiedGroup $GroupName -ErrorAction SilentlyContinue
        New-Team -GroupID $GroupObject.ExternalDirectoryObjectID -ErrorAction SilentlyContinue | out-null;
        $CreatedTeams += ,$GroupName;
        Start-Sleep -Seconds 10;
        $GroupObject = "";
    }

    #Add the student from the group as appropriate
    #It is significantly faster to re-perform this operation than it is to compare whether it needs to be done
    if ($Status -eq "InClass") {
        #write-host "Adding" $Student "to Group" $GroupName;
        Add-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $Student -ErrorAction SilentlyContinue;
    }    
}

write-host "The script has completed.";
write-host "The total number of skipped lines was $($skippedInstructorsCount + $skippedCoursesCount).";
write-host "The total number of skipped TBA lines was $skippedInstructorsCount.";
write-host "The total number of lines skipped due to Course title was $skippedCoursesCount.";
write-host "If any errors appeared during the script run please either screenshot the entire error message including the lines before and after and provide the screenshot to Dan."  ;