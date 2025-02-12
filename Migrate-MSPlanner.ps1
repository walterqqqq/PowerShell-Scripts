<#
.SYNOPSIS
Tenant to tenant planner migration - migrate-planner.ps1

.NOTES
Author:         Jakob Schaefer, thinformatics AG
Version:        1.0
Creation Date:  2020/07/02

.DESCRIPTION
This script is designed to migrate planner plans from a o365 group in a source tenant to a o365 group in a destination tenant. 
The script is based on the publication of Alexander Holmeset (https://alexholmeset.blog/2019/10/14/planner-tenant-to-tenant-migration/) and was adjusted to fullfill 
authentication and content requirements.

Please note, this script will only migrate the most recent 400 tasks!

What will be migrated:
Plans, Categories, Buckets, Tasks, 
TaskDetails including: Assignments, Progress, Priority, StartDate, Due Date, CheckList Items, Notes

What will not be migrated:
Task Attachments, Task Comments

Prerequisites:
- The O365 Group which should host the planner in the destination Tenant has to exist
- To allow the task assignment in the destination tenant, the related user have to exist already in the destination Tenant and should be member / owner of the affected O365 Group.
- To get all plans from a o365 group, that the user which is used to authenticate against the source tenant , is at least a group member
- We map the planner members by theire UPN address. We do this by replacing the e-mail domain in the username (john.doe@sourcetenant.com > john.doe@destinationtenant.com). Use the parameters $sourceDomain and $destinationDomain to get this done. You can consider to make this parameters static to minimize the effort to start the script
- The user which is used for authentication in this script has to be member and owner in the source and destination O365 Group

Prepare Script:

We make use of an app based authentication to export and import planner data. To get this done we have to create Apps in the source and Destination tenant. 
To do so proceed as described here:
1. Login to the Source Tenant as Admin with the Role "Application Administrator permissions" (e.g.) to https://aad.portal.azure.com and navigate to "Azure Active Directory" > App Registrations
2. Create a new registration with "New registration"
3. Determine a name for the new registration. E.g.: "PlannerMigration"
4. The specified account type should be "Accounts in this organizational directory only"
5. Determine a name for the Redirect URI (Web): https://plannermigration.sourcetenant.onmicrosoft.com
6. In the new created App registration navigate Overview and note the ApplicationID
7. Navigate to "Certificates & secrets".
8. Create a new client secret by click on "+ New client secret"
9. Chose an expiry date range (1 year should be enough in common migration projects) and give it a description (e.g.: Creation Time+Requested User Name: 052020 Jakob)
10. After click on "Add" you will get the client secret. Please note it immediately, you can´t make it visible again later. 
11. Navigate to API Permissions and click on "Add a permission"
12. Chose Microsoft Graph API > Delegated Permissions > and chose the following permissions: Group.Read.All, User.Read.All
13. Grant admin consent to the new application by click on "Grant admin consent for ..."
14. Insert the noted ClientID, ClientSecret and Redirect URI in this script at line 96 ff (TBD)
15. Repeat steps 1.-14. in the destination Tenant. In step 12 at the destination Tenant chose the following permissions: Group.ReadWrite.All, User.Read.All
16. Insert the noted ClientID, ClientSecret and Redirect URI in this script at line 101 ff (TBD)

ErrorHandling:
- If you wan´t to repeat a planner migration you will have the plan twice in the destination. The tasks will not be merged. Consider to remove the primarily created planner before running this script again for the same plan.
- Planner Tile Conflicts will not be a problem. A display Name of a plan can exist in unlimited instances


.PARAMETER SourceGroupID
Enter the ID of the Group which contains plans that should be migrated
.PARAMETER DestinationGroupID
Enter the ID of the Group to which the planner should be migrated
.PARAMETER sourceDomain
This parameter is used for the usermapping, Insert the Domain with the @-prefix of the recipients of the source tenant here, which have assignments in the plan
.PARAMETER destinationDomain
This parameter is used for the usermapping, Insert the Domain with the @-prefix of the recipients of the destination tenant here.
.PARAMETER singlePlanId
Use this parameter if you only want to migrate a single plan out of a group to the destination group.
If you have a link to the planner, you can find out the group by extracting it from the URL
e.G.: https://tasks.office.com/M365x763939.onmicrosoft.com/en-US/Home/Planner/#/plantaskboard?groupId=eabf0a16-53e3-4443-9557-130fc4f3773e&planId=fmmWu0CxoE2gOAlxhSnyJ5gAFqs3 = eabf0a16-53e3-4443-9557-130fc4f3773e
.PARAMETER userCheck
if $true a primt will come up if no related user was found in the destination tenant to assign the task. You can manually adjust the assignee with this inforation. Be aware that the mapping has to be done for every single task again.

.EXAMPLE
.\migrate-planner_doc.ps1 -SourceGroupID 927e47e7-09f4-4b38-8aac-80f68d121801 -DestinationGroupID dba6eaa0-e714-4f81-8cb8-8295611a3cfa -sourceDomain "oldworld.com" -destinationDomain "@newworld.de"
All planners from Group 927e47e7-09f4-4b38-8aac-80f68d121801 will be migrated to the destination Group dba6eaa0-e714-4f81-8cb8-8295611a3cfa. The assignment of the task will be john.doe@oldworld.com -> john.doe@neworld.de. Task assignments with users which can´t be found in the destination tenant, will be skipped.

.EXAMPLE
.\migrate-planner_doc.ps1 -SourceGroupID 927e47e7-09f4-4b38-8aac-80f68d121801 -DestinationGroupID dba6eaa0-e714-4f81-8cb8-8295611a3cfa -sourceDomain "oldworld.com" -destinationDomain "@newworld.de" -singlePlanID "7F6Hc9ReTkCh7do1q6BXSpcAD_i6"
The Plan with the ID 7F6Hc9ReTkCh7do1q6BXSpcAD_i6 from Group 927e47e7-09f4-4b38-8aac-80f68d121801 will be migrated to the destination Group dba6eaa0-e714-4f81-8cb8-8295611a3cfa. The assignment of the task will be john.doe@oldworld.com -> john.doe@neworld.de. Task assignments with users which can´t be found in the destination tenant, will be skipped.

.EXAMPLE
.\migrate-planner_doc.ps1 -SourceGroupID 927e47e7-09f4-4b38-8aac-80f68d121801 -DestinationGroupID dba6eaa0-e714-4f81-8cb8-8295611a3cfa -sourceDomain "oldworld.com" -destinationDomain "@newworld.de" -singlePlanID "7F6Hc9ReTkCh7do1q6BXSpcAD_i6" -usercheck $true
The Plan with the ID 7F6Hc9ReTkCh7do1q6BXSpcAD_i6 from Group 927e47e7-09f4-4b38-8aac-80f68d121801 will be migrated to the destination Group dba6eaa0-e714-4f81-8cb8-8295611a3cfa. The assignment of the task will be john.doe@oldworld.com -> john.doe@neworld.de. If the script can´t find a dependant user in the destination tenant, you are able to insert the dependant account here manually.

#>


[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)][string]$SourceGroupID,
    [Parameter(Mandatory=$true)][string]$DestinationGroupID, 
    [Parameter(Mandatory=$false)][boolean]$usercheck, 
    [Parameter(Mandatory=$false)][string]$singlePlanID,
    [Parameter(Mandatory=$false)][string]$sourceDomain,
    [Parameter(Mandatory=$false)][string]$destinationDomain 
    
)

#Source App Data: Fill in Client ID, Secret and RedirectUri here:
$ClientIDSource="Enter Client ID for Source"
$ClientSecretSource="Enter Secret for Source"
$redirectUriSource="https://plannermigration.auerhaan.onmicrosoft.com"

#Destination App Data: Fill in Client ID, Secret and RedirectUri here:
$clientidDestination = "Enter Client ID for Destination"
$clientSecretDestination = "Enter Secret for Destination"
$redirectUriDestination = "https://plannermigration.mc1s.onmicrosoft.com"


# The resource URI
$resource = "https://graph.microsoft.com"

# Function to popup Auth Dialog Windows Form
Function Get-AuthCode {
    Add-Type -AssemblyName System.Windows.Forms

    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url -f ($Scope -join "%20")) }

    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri        
        if ($Global:uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null

    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
    }

    $output
}

#####################################################################
#################SourceConnection###############################
#####################################################################

Write-Host -ForegroundColor Gray -Object "Action: Please Logon to the AzureAD of the SOURCE Tenant." # TBD Permissions Admin or any user?
#Connect-AzureAD
# UrlEncode the ClientID and ClientSecret and URL's for special characters 
Add-Type -AssemblyName System.Web
$clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($ClientIDSource)
$clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($ClientSecretSource)
$redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($redirectUriSource)
$resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resource)
$scopeEncoded = [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/user.readwrite.all")

# Get AuthCode
$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientIDSource&resource=$resourceEncoded&prompt=admin_consent&scope=$scopeEncoded"
Get-AuthCode
# Extract Access token from the returned URI
$regex = '(?<=code=)(.*)(?=&)'
$authCode  = ($uri | Select-string -pattern $regex).Matches[0].Value

#Write-output "Received an authCode, $authCode"

#get Access Token
$body = "grant_type=authorization_code&redirect_uri=$redirectUriSource&client_id=$clientIdSource&client_secret=$clientSecretEncoded&code=$authCode&resource=$resource"
$tokenResponse = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
    -Method Post -ContentType "application/x-www-form-urlencoded" `
    -Body $body `
    -ErrorAction STOP

    

#####################################################################
#################DestinationConnection###############################
#####################################################################

Write-Host -ForegroundColor Gray -Object "Action: Please Logon to the AzureAD of the DESTINATION Tenant." # TBD Permissions Admin or any user?
#Connect-AzureAD
# UrlEncode the ClientID and ClientSecret and URL's for special characters 
Add-Type -AssemblyName System.Web
$clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($clientidDestination)
$clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($clientSecretDestination)
$redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($redirectUriDestination)
$resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resource)
$scopeEncoded = [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/user.readwrite.all")

# Get AuthCode
$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientIDDestination&resource=$resourceEncoded&prompt=admin_consent&scope=$scopeEncoded"
Get-AuthCode
# Extract Access token from the returned URI
$regex = '(?<=code=)(.*)(?=&)'
$authCode  = ($uri | Select-string -pattern $regex).Matches[0].Value

#Write-output "Received an authCode, $authCode"

#get Access Token
$body = "grant_type=authorization_code&redirect_uri=$redirectUriDestination&client_id=$clientIdDestination&client_secret=$clientSecretEncoded&code=$authCode&resource=$resource"
$tokenResponseDESTINATION = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
    -Method Post -ContentType "application/x-www-form-urlencoded" `
    -Body $body `
    -ErrorAction STOP


##################################################################now do stuff

#define token vars
$tokenSource=$tokenResponse.access_token
$tokenDestination=$tokenResponseDESTINATION.access_token


#find source group
$groupIDSource = $SourceGroupID
$findgroupuri="https://graph.microsoft.com/v1.0/groups/"+$groupIDSource
$group=Invoke-RestMethod -Headers @{Authorization = "Bearer $tokenSource"} -Uri $findgroupuri -Method Get -ErrorAction Stop


#find destination group
$groupIDDestination = $DestinationGroupID
$findgroupuri="https://graph.microsoft.com/v1.0/groups/"+$groupIDDestination
$groupsdestination=Invoke-RestMethod -Headers @{Authorization = "Bearer $tokenDestination"} -Uri $findgroupuri -Method Get -ErrorAction Stop

#get planners in the o365 Group
if($singlePlanID){
    $sourceplans=@()
    $urisingleplan = 'https://graph.microsoft.com/v1.0/planner/plans/' + $singlePlanID
    $queryurisingleplan = Invoke-RestMethod -Method GET -Uri $urisingleplan -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokenSource" }
    $sourceplans=$queryurisingleplan
}else {
    $groupID = $group.id
    $groupDisplayName = $group.displayName
    $uri2 = 'https://graph.microsoft.com/v1.0/groups/' + $groupID + '/planner/plans'
    $query2 = Invoke-RestMethod -Method GET -Uri $uri2 -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokenSource" }
    $sourceplans = $query2.value
}

$i=0
Foreach ($plan in $sourceplans) {
    $i++
    Write-Host -ForegroundColor Gray -Object ("INFO: Processing Plan $i/$($sourceplans.count) - Plan Title: $($plan.title)")

    #Plan ID in source tenant.
    $planID = $plan.id

    #Finds all buckets in plan.
    $uri5 = 'https://graph.microsoft.com/v1.0/planner/plans/' + $planID + '/buckets/'
    $query5 = Invoke-RestMethod -Method GET -Uri $uri5 -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokenSource" }
    $buckets = $query5.value | Sort-Object orderhint -Descending

    #Finds all categories in plan
    $uriDestination545 = 'https://graph.microsoft.com/v1.0/planner/plans/' + $planID + '/details/'
    $query545 = Invoke-RestMethod -Method GET -Uri $uriDestination545 -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokenSource" }
    $categories = ((($query545.categoryDescriptions).psobject.members) | Where-Object { $_.MemberType -eq 'NoteProperty' }) | Select-Object name, value


    #Destination Plan Definition
    $GroupDestinationID = $groupsdestination.id
    $RequestBody2 = @"
    {
        "owner": "$groupdestinationID",
        "title": "$($plan.title)",
        }
"@
    #Creates plan in destination tenant.
    $uriDestination2 = 'https://graph.microsoft.com/v1.0/planner/plans/'
    $queryDestination2 = Invoke-RestMethod -Method POST -Uri $uriDestination2   -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokenDestination" } -Body $requestbody2
    $planDestination = $queryDestination2
    $planDestinationID = $queryDestination2.id
    
    #Create labels in destination tenant.
    $uriDestination223423 = 'https://graph.microsoft.com/v1.0/planner/plans/' + $planDestinationID + '/details'
    $queryDestination223423 = Invoke-RestMethod -Method GET -Uri $uriDestination223423   -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokenDestination" }
    $planDestinationEtag = $queryDestination223423.'@odata.etag'

    $headers = @{ }
    $headers.Add("if-match", $planDestinationEtag)
    $headers.Add("Authorization", "Bearer $tokendestination")

    
    if (($categories | Where-Object { $_.name -eq 'category1' }).value) { 
        $category1 = '"category1": "' + (($categories | Where-Object { $_.name -eq 'category1' }).value) + '",' 
    } Else { 
        $category1 = '"category1" : null,' 
    }
    

    if (($categories | Where-Object { $_.name -eq 'category2' }).value) { 
        $category2 = '"category2": "' + (($categories | Where-Object { $_.name -eq 'category2' }).value) + '",' 
    } Else {
        $category2 = '"category2" : null,'
    }

    if (($categories | Where-Object { $_.name -eq 'category3' }).value) { 
        $category3 = '"category3": "' + (($categories | Where-Object { $_.name -eq 'category3' }).value) + '",' 
    } Else { 
        $category3 = '"category3" : null,' 
    }

    if (($categories | Where-Object { $_.name -eq 'category4' }).value) { 
        $category4 = '"category4": "' + (($categories | Where-Object { $_.name -eq 'category4' }).value) + '",' 
    } Else { 
        $category4 = '"category4" : null,' 
    }

    if (($categories | Where-Object { $_.name -eq 'category5' }).value) { 
        $category5 = '"category5": "' + (($categories | Where-Object { $_.name -eq 'category5' }).value) + '",' 
    } Else { 
        $category5 = '"category5" : null,' 
    }

    if (($categories | Where-Object { $_.name -eq 'category6' }).value) { 
        $category6 = '"category6": "' + (($categories | Where-Object { $_.name -eq 'category6' }).value) + '",' 
    }Else { 
        $category6 = '"category6" : null,' 
    }
    

    $RequestBody223423 = @"
    {
        "categoryDescriptions": {
            $category1
            $category2
            $category3
            $category4
            $category5
            $category6
        }
        }
"@

    $queryDestination2342342 = Invoke-RestMethod -Uri $uriDestination223423 -Headers $headers -Method PATCH -Body $RequestBody223423 -ContentType "application/json; charset=utf-8"

    $BucketOverview =@()
    Foreach ($bucket in $buckets) {

        #Creates plan buckets in destiantion tenant.
        $bucketnameDestination = $bucket.name
        $planDestinationID = $planDestination.id
        $RequestBody3 = @"
                        {
                "name": "$bucketnamedestination",
                "planId": "$planDestinationID",
                "orderHint": " !"
            }
"@
        $uriDestination3 = 'https://graph.microsoft.com/v1.0/planner/buckets/'
        $queryDestination3 = Invoke-RestMethod -Method POST -Uri $uriDestination3   -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokenDestination" } -Body $requestbody3
        $BucketDestination = $queryDestination3

        $Object = [PSCustomObject]@{

            'BucketIDSource'      = $bucket.id
            'BucketIDDestination' = $BucketDestination.id
            'BucketName'          = $bucketnameDestination

        }
        $BucketOverview += $Object
    }

    
    #Finds all tasks of the plan in source tenant.
    $uri55 = 'https://graph.microsoft.com/v1.0/planner/plans/' + $planID + '/tasks/'
    $query55 = Invoke-RestMethod -Method GET -Uri $uri55 -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokenSource" }
    $tasks = $query55.value | Sort-Object orderHint  -Descending

    $i2=0
    foreach ($task in $tasks) {
        $i2++
        Write-Host -ForegroundColor Gray -Object ("INFO: Processing Task $i2/$($tasks.count) - Task title: $($task.title)")

        $taskBucketID = $task.bucketId
        $taskTitle = $task.title
        $taskID = $task.id
        $taskPercentComplete = $task.percentComplete

        $TaskBucketDestinationID = ($BucketOverview | Where-Object { $_.BucketIDSource -eq $taskBucketID }).BucketIDDestination
        $TaskstartDateTime = If($task.startDateTime){ $task.startDateTime | Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ'}
        $TaskdueDateTime = If($task.dueDateTime){$task.dueDateTime | Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ'}

        If (!$TaskstartDateTime) { 
            '' 
        }Else { 
            $TaskstartDateTime = [string]('"startDateTime" : "' + $TaskstartDateTime + '",') 
        }
        If (!$TaskdueDateTime) { 
            '' 
        } Else { 
            $TaskdueDateTime = [string]('"dueDateTime" : "' + $TaskdueDateTime + '",') 
        }

        $taskappliedCategories = ((($task.appliedCategories).psobject.members) | Where-Object { $_.membertype -eq 'NoteProperty' }).name
        $CustomAppliedCategory = @()
        $AppliedCategoriesBody = @()
            If ($taskappliedCategories) {
                
                foreach ($appliedcategory in $taskappliedCategories) {
                    $applied = '"' + $appliedcategory + '": true,
                    '
                    $CustomAppliedCategory += $applied
                }

                $AppliedCategoriesBody = @"
                "appliedCategories": {
                    $customappliedcategory
                },
"@

            }

if ($assignees) {
    $Addusers = @()
    foreach ($assignee in $assignees) {
        # Initialize variables
        $DestinationUserID = $null
        $DestinationUPN = $null
        $SourceUPN = $null

        # Attempt to retrieve the source user details
        try {
            $uri555 = 'https://graph.microsoft.com/v1.0/users/' + $assignee
            $query555 = Invoke-RestMethod -Method GET -Uri $uri555 -ContentType "application/json; charset=utf-8" `
                -Headers @{ Authorization = "Bearer $tokenSource" }
            $SourceUPN = $query555.userPrincipalName
        }
        catch {
            Write-Host -ForegroundColor Yellow "WARNING: Unable to retrieve source user for ID '$assignee'. Skipping assignment."
            continue
        }

        if ([string]::IsNullOrEmpty($SourceUPN)) {
            Write-Host -ForegroundColor Yellow "WARNING: SourceUPN is empty for user ID '$assignee'. Skipping assignment."
            continue
        }

        # Map the source UPN to the destination UPN (replace source domain with destination domain)
        $DestinationUPN = $SourceUPN.Replace($sourceDomain, $destinationDomain)

        # Lookup the destination user
        if ($SourceUPN -like "*#EXT#*") {
            # For guest users, try looking up by mail property
            try {
                $uri5555 = "https://graph.microsoft.com/v1.0/users/?`$filter=mail eq '$($query555.mail)'"
                $query5555 = Invoke-RestMethod -Method GET -Uri $uri5555 -ContentType "application/json; charset=utf-8" `
                    -Headers @{ Authorization = "Bearer $tokenDestination" }
                if ($query5555.value.Count -gt 0) {
                    $DestinationUserID = $query5555.value[0].id
                }
                else {
                    Write-Host -ForegroundColor Yellow "WARNING: Destination guest user with mail '$($query555.mail)' not found. Skipping assignment."
                    continue
                }
            }
            catch {
                Write-Host -ForegroundColor Yellow "WARNING: Error retrieving destination guest user for '$($query555.mail)'. Skipping assignment."
                continue
            }
        }
        else {
            try {
                $uri5555 = 'https://graph.microsoft.com/v1.0/users/' + $DestinationUPN
                $query5555 = Invoke-RestMethod -Method GET -Uri $uri5555 -ContentType "application/json; charset=utf-8" `
                    -Headers @{ Authorization = "Bearer $tokenDestination" }
                $DestinationUserID = $query5555.id
            }
            catch {
                Write-Host -ForegroundColor Yellow "WARNING: Destination user '$DestinationUPN' not found. Skipping assignment."
                continue
            }
        }

        # If a destination user was found, build the assignment snippet
        if ($DestinationUserID) {
            $AddUser = @"
"$DestinationUserID": {
    "@odata.type": "#microsoft.graph.plannerAssignment",
    "orderHint": " !"
},
"@
            $Addusers += $AddUser
        }
    }  # End foreach ($assignee in $assignees)

    # Build the assignments section only if we have any valid mapped assignments
    if ($Addusers.Count -gt 0) {
        $AssignmentsSection = @"
"assignments": {
    $($Addusers -join "`n")
},
"@
    }
    else {
        $AssignmentsSection = ""
    }

    # Build the request body including assignments
    $RequestBody4 = @"
{
    "planId": "$planDestinationID",
    "bucketId": "$TaskBucketDestinationID",
    "title": "$tasktitle",
    "percentComplete": "$taskpercentcomplete",
    $TaskstartDateTime
    $TaskdueDateTime
    $AppliedCategoriesBody
    $AssignmentsSection
}
"@
}
else {
    # No assignments – build the request without the assignments section
    $RequestBody4 = @"
{
    "planId": "$planDestinationID",
    "bucketId": "$TaskBucketDestinationID",
    "title": "$tasktitle",
    "percentComplete": "$taskpercentcomplete",
    $TaskstartDateTime
    $TaskdueDateTime
    $AppliedCategoriesBody
}
"@
}

        start-sleep 5

        #Creates task in destination tenant.
        $uriDestination4 = 'https://graph.microsoft.com/v1.0/planner/tasks/'
        $queryDestination4 = Invoke-RestMethod -Method POST -Uri $uriDestination4 -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokendestination" } -Body $RequestBody4
        $TaskIDDestination = $queryDestination4.ID

        $uriDestination6 = 'https://graph.microsoft.com/v1.0/planner/tasks/' + $TaskIDDestination + '/details'
        $queryDestination6 = Invoke-RestMethod -Method GET -Uri $uriDestination6 -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokendestination" }

        $uri4323426 = 'https://graph.microsoft.com/v1.0/planner/tasks/' + $taskID + '/details'
        $query4323426 = Invoke-RestMethod -Method GET -Uri $uri4323426 -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokenSource" }
        
        $TaskEtagDestination = $queryDestination6.'@odata.etag'
        $Description =  $query4323426.description
        
        
        
        If (!$Description) { 
            $Description = '"description" : null,' 
        }Else { 
            $Description = '"description" : "' + $Description + '"' 
        }

        $RequestBody23111 = @"
                {
                    $Description  
                }
"@

        $headers = @{ }
        $headers.Add("if-match", $TaskEtagDestination)
        $headers.Add("Authorization", "Bearer $tokendestination")

        #Updates task in destination tenant.
        $queryDestination6154 = Invoke-RestMethod -Uri $uriDestination6 -Headers $headers -Method PATCH -Body $RequestBody23111 -ContentType "application/json; charset=utf-8"

        $uri123 = 'https://graph.microsoft.com/v1.0/planner/tasks/' + $taskID + '/details'
        $query123 = Invoke-RestMethod -Method GET -Uri $uri123 -ContentType "application/json; charset=utf-8" -Headers @{Authorization = "Bearer $tokenSource" }
        $TaskDetails = (($query123.checklist).psobject.Properties).value

        Foreach ($checklistitem in $TaskDetails) {

            #Destination
            $checklistItemGuid = (New-Guid).Guid
            $IsChecked = $checklistitem.isChecked
            $CheckListItemTitle = $checklistitem.title

            $RequestBody5 = @"
                {
                    "checklist": {
                        "$checklistItemGuid": {
                            "@odata.type": "#microsoft.graph.plannerChecklistItem",
                            "isChecked": "$IsChecked",
                            "title": "$checklistitemtitle"
                        }
                    }
                }
"@

            $headers = @{ }
            $headers.Add("if-match", $TaskEtagDestination)
            $headers.Add("Authorization", "Bearer $tokendestination")
            
            #Creates checklist for task in destiantion tenant.
            $uriDestination5 = "https://graph.microsoft.com/v1.0/planner/tasks/" + "$TaskIDDestination" + "/Details"
            $queryDestination5 = Invoke-RestMethod -Uri $uriDestination5 -Headers $headers -Method PATCH -Body $RequestBody5 -ContentType "application/json; charset=utf-8"
        }   
    }
}
