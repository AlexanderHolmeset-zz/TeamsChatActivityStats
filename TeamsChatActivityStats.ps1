
### Teams Chat Activity Stas                    ###

### Version 1.0                                 ###

### Author: Alexander Holmeset                  ###

### Twitter: twitter.com/alexholmeset           ###

### Blog: alexholmeset.blog                     ###

#Description:
#You specify a Group/team object ID, then the script archives all conversations in every channel for this team in a HTML file.
#Have in mind information protection policies like GDPR when working on information like this.


#Parameters
#
#This is a generic PowerShell Azure Client ID
#$clientId = "1950a258-227b-4e31-a9cf-717495945fc2"
$redirectUri = "urn:ietf:wg:oauth:2.0:oob"
$resourceURI = "https://graph.microsoft.com"
$authority = "https://login.microsoftonline.com/common"
#Its recomended to register your own Azure application, to controll what access rights you need.
#Here is a example of a client ID and redirect URI i have created.
$clientId = "b379ceb0-108c-467f-9749-c8c13f9131b1"
#$redirectUri = "https://login.microsoftonline.com/M365x792147.onmicrosoft.com/oauth2"


#Remove commenting on username and password if you want to run this without a prompt.
#$Office365Username='user@domain'
#$Office365Password='VeryStrongPassword' 


#pre requisites
try {
$AadModule = Import-Module -Name AzureAD -ErrorAction Stop -PassThru
}
catch {
throw 'Prerequisites not installed (AzureAD PowerShell module not installed)'
}
$adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
$adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
[System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
[System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
 
##option without user interaction
if (([string]::IsNullOrEmpty($Office365Username) -eq $false) -and ([string]::IsNullOrEmpty($Office365Password) -eq $false))
{
$SecurePassword = ConvertTo-SecureString -AsPlainText $Office365Password -Force
#Build Azure AD credentials object
$AADCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential" -ArgumentList $Office365Username,$SecurePassword
# Get token without login prompts.
$authResult = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($authContext, $resourceURI, $clientid, $AADCredential);
}
else
{
# Get token by prompting login window.
$platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Always"
$authResult = $authContext.AcquireTokenAsync($resourceURI, $ClientID, $RedirectUri, $platformParameters)
}
$accessToken = $authResult.result.AccessToken



$Total = [int]"0"
$TotalRoot = [int]"0"
$TotalReplies = [int]"0"
$ToCSV = @()
$Date = get-date


#Gets all channels in a Team
$apiUrl = 'https://graph.microsoft.com/beta/groups?$select=id,resourceProvisioningOptions,DisplayName'
$myProfile = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -Method Get
$Teams = $myprofile.value | Select-Object ID,DisplayName,resourceProvisioningOptions



#Where to store the HTML file:
$Storage = 'c:\temp\test.html'
"" | Out-File $Storage

foreach($Team in $Teams){
If($team.resourceProvisioningOptions -eq "Team"){
#Group/team object ID. 
$TeamID = $Team.ID
$TeamDisplayName = $Team.DisplayName
$TotalTeam = [int]"0"
$RootCountTeam = [int]"0"
$RepliesCountTeam = [int]"0"





#Gets all channels in a Team
$apiUrl = "https://graph.microsoft.com/beta/teams/$TeamID/channels"
$myProfile = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -Method Get
$TeamChannels = $myprofile.value | Select-Object ID,DisplayName
$teamdisplayname


"<br>
<br>
------------------------------<br>
$TeamDisplayName<br>"| Out-File -Append $Storage
 
foreach($Channel in $TeamChannels) {
$TotalChannel = [int]"0"
#Gets all root messages/conversations in a channel.
$apiUrl = "https://graph.microsoft.com/beta/teams/$TeamID/channels/"+$channel.id+"/messages"
$myProfile = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -Method Get
$ChannelMessages = $myprofile.value | Select-Object Body,From,ID,attachments,createdDateTime | Sort-Object
$Channeldisplayname = $channel.displayName
$ChannelMessagesCount = ($ChannelMessages).count
$Total = $Total + [int]$ChannelMessagesCount
$TotalTeam = $TotalTeam + [int]$ChannelMessagesCount
$TotalRoot = $TotalRoot + [int]$ChannelMessagesCount
$RootCountTeam = $RootCountTeam + [int]$ChannelMessagesCount
$TotalChannel = $TotalChannel + [int]$ChannelMessagesCount


$RepliesCountChannel = [int]"0"
foreach($channelmessage in $ChannelMessages){

#Gets all replies in a channel.
$apiUrl = "https://graph.microsoft.com/beta/teams/$TeamID/channels/"+$channel.id+"/messages/"+$ChannelMessage.id+"/replies"
$myProfile = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -Method Get
$RepliesCount = ($myprofile.value).count
$RepliesCountChannel = $RepliesCountChannel + [int]$RepliesCount
$Total = $Total + [int]$RepliesCount
$TotalTeam = $TotalTeam + [int]$RepliesCount
$TotalChannel =  $TotalChannel+[int]$RepliesCount
$TotalReplies = $TotalReplies + [int]$RepliesCount
$RepliesCountTeam = $RepliesCountTeam + [int]$RepliesCount


    }

"<br>
$ChannelDisplayname<br>
Number of root messages in channel: $ChannelMessagesCount<br>
Number of replies channel: $RepliescountChannel<br>
Total messages in channel: $TotalChannel<br>
" | Out-File -Append $Storage









}

"<br>
<br>
Number of root messages in team: $RootCountTeam<br>
Number of Replies in Team: $RepliesCountTeam<br>
Total Messages in Team: $TotalTeam<br>
<br>
<br>" | out-file  -Append $Storage




$Object=[PSCustomObject]@{
    Team = $TeamDisplayName
    TotalMessages = $TotalTeam
    Date = $Date
     }#EndPSCustomObject
    $ToCSV+=$object
    }
}


"Total number of root messages in Teams: $TotalRoot<br>
Total number of replies in Teams: $TotalReplies<br>
Total number of messages in Teams: $Total" | out-file -Append $Storage
$ToCSV | export-csv c:\temp\TeamsStats.csv
