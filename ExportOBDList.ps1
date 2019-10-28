#ExportODBLists
#Exports  ODB experience settings, stored in several SharePoint Lists

param([string]$siteUrl, [bool]$exportAllFields=$false, [bool]$useStoredCreds=$true, [string]$exportFolder)
Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\SharePointPnPPowerShellOnline\3.12.1908.1\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\SharePointPnPPowerShellOnline\3.12.1908.1\Microsoft.SharePoint.Client.Runtime.dll"

if (!$siteUrl)
{
    Write-Host "Please specify a OneDrive site using -siteUrl."
    return
}

if ($useStoredCreds)
{
    Write-Host "Retrieving stored Windows credentials for $siteUrl."
    $cred = Get-StoredCredential -Target $siteUrl
    if (!$cred)
    {
        Write-Host "Didn't find stored credential for $siteUrl. Please provide credentials to connect."
        $cred = Get-Credential
    }
}
else
{
   $cred = Get-Credential
}

$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.UserName,$cred.Password)
$webURL = $siteUrl
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
$ctx.Credentials = $credentials

#root folders of lists to export
$SWMRoot = "Reference " #starts with this string
$notificationsRoot = "notificationSubscriptionHiddenList6D1E55DA25644A22"
$activityFeedRoot = "userActivityFeedHiddenListF4387007BE61432F8BDB85E6"
$accessRequestsRoot = "Access Requests"
$microfeedRoot = "PublishedFeed"
$SPHomeCacheRoot = "SharePointHomeCacheList"
$sharingLinksRoot = "Sharing Links"
$socialRoot = "Social"

#list fields to eexport
$SWMFields = @("ID","Created","Modified","Title","RemoteItemPath","OwnerDisplayName","OwnerSipAddress","RemoteItemFileSystemObjectType",
                "RemoteItemCreatorDisplayName","RemoteItemCreatorSipAddress","RemoteItemCreatedDateTime",
                "RemoteItemModifierDisplayName","RemoteItemModifierSipAddress","RemoteItemModifiedDateTime",
                "SWMSharedByDisplayName","SWMSharedBySipAddress","SWMSharedByDateTime",
                "RemoteItemLastAccessedDateTime","RemoteItemServerRedirectedUrl","RemoteItemServerRedirectedEmbedUrl")
$accessRequestsFields = @("ID","Created","Modified","Title","RequestId","RequestedObjectTitle","RequestedObjectUrl","PermissionType","PermissionLevelRequested","RequestDate",
                            "RequestedByDisplayName","RequestedBy","ReqByUser",
                            "RequestedForDisplayName","RequestedFor","ReqForUser",
                            "ApprovedBy","AcceptedBy","Status","Expires","WelcomeEmailSubject","WelcomeEmailBody","ExtendedWelcomeEmailBody","Conversation")
$microfeedFields = @("ID","Created","Modified","Title","MicroBlogType","PostAuthor","RootPostOwnerID","RootPostUniqueID","ReplyCoungett","Order","ContentData")
$notificationsFields = @("ID","Created","Modified","Title","SubscriptionId","PoolName","SecondaryPoolName","AppType","NotificationHandle",
                            "SecondsToExpiry","DestinationType","SubmissionDateTime","ExpirationDateTime","Locale","DeviceId","HostName","NotificationCounter",
                            "SingleSignOutKey","NotificationScenarios","ComplianceAssetId","AppAuthor","AppEditor")
$SPHomeCacheFields = @("ID","Created","Modified","Author","Editor","Title","Value")
$sharingLinksFields = @("ID","Created","Modified","Title","SharingDocId","ComplianceAssetId","CurrentLink","AvailableLinks")
$socialFields = @("ID","Created","Modified","Author","Editor","Title","Url","Hidden","HasFeed","SocialProperties")
$activityFeedFields = @("ID","Created","Modified","Title","ActivityId","ItemId","PushNotificationsSent","EmailNotificationSent","IsActorActivity","IsRead","Order",
                        "ItemChildCount","FolderChildCount","ActivityEventType","ActivityEvent")


#get lists in the web
try{
    $lists = $ctx.web.Lists
    $ctx.load($lists)
    $ctx.executeQuery()
}
catch{
	write-host "$($_.Exception.Message)" -foregroundcolor red
}


#identify the lists to export
$listsToExport = @()
foreach($list in $lists)
{
    $ctx.load($list)
    $ctx.load($list.RootFolder)
    $ctx.executeQuery()
    $listTitle = [string]$list.Title
    $listRoot = $list.RootFolder.Name
    
    Write-host ("Processing List: " + $list.Title + " with " + $list.ItemCount + " items").ToUpper() -ForegroundColor Yellow
    Write-host (">> List Root Folder: " + $listRoot) -ForegroundColor Yellow

    if ($listRoot.StartsWith($SWMRoot,"CurrentCultureIgnoreCase") -and $list.ItemCount -ge 1)
    {
        Write-Host ">> Found: Shared With Me List" -ForegroundColor Green
        $listDetails = @{listType = "Shared With Me List"; listTitle = $listTitle; listRoot = $listRoot; listFields = $SWMFields}
        $listsToExport += $listDetails
    }
    elseif ($listRoot -eq $notificationsRoot)
    {
        Write-Host ">> Found: Notifications List" -ForegroundColor Green
        $listDetails = @{listType = "Notifications List"; listTitle = $listTitle; listRoot = $listRoot; listFields = $notificationsFields}
        $listsToExport += $listDetails
    }
    elseif ($listRoot -eq $activityFeedRoot)
    {
        Write-Host ">> Found: User Activity Feed List" -ForegroundColor Green
        $listDetails = @{listType = "User Activity Feed List"; listTitle = $listTitle; listRoot = $listRoot; listFields = $activityFeedFields}
        $listsToExport += $listDetails
    }
    elseif ($listRoot -eq $accessRequestsRoot)
    {
        Write-Host ">> Found: Access Requests List" -ForegroundColor Green
        $listDetails = @{listType = "Access Requests List"; listTitle = $listTitle; listRoot = $listRoot; listFields = $accessRequestsFields}
        $listsToExport += $listDetails
    }
    elseif ($listRoot -eq $microfeedRoot)
    {
        Write-Host ">> Found: MicroFeed List" -ForegroundColor Green
        $listDetails = @{listType = "Microfeed List"; listTitle = $listTitle; listRoot = $listRoot; listFields = $microfeedFields}
        $listsToExport += $listDetails
    }
    elseif ($listRoot -eq $SPHomeCacheRoot)
    {
        Write-Host ">> Found: SharePoint Home Cache List" -ForegroundColor Green
        $listDetails = @{listType = "SharePoint Home Cache List"; listTitle = $listTitle; listRoot = $listRoot; listFields = $SPHomeCacheFields}
        $listsToExport += $listDetails
    }
    elseif ($listRoot -eq $sharingLinksRoot)
    {
        Write-Host ">> Found: Sharing Links List" -ForegroundColor Green
        $listDetails = @{listType = "Sharing Links List"; listTitle = $listTitle; listRoot = $listRoot; listFields = $sharingLinksFields}
        $listsToExport += $listDetails
    }
    elseif ($listRoot -eq $socialRoot)
    {
        Write-Host ">> Found: Social List" -ForegroundColor Green
        $listDetails = @{listType = "Social List"; listTitle = $listTitle; listRoot = $listRoot; listFields = $socialFields}
        $listsToExport += $listDetails
    }

}

#export list function
function exportList
{
    Param ([string] $listTitle, [string[]]$listFields, [string]$exportFile)

    Write-Host ("Exporting List: " + $listTitle).ToUpper() -ForegroundColor Green
    Write-Host (">> File location: $exportFile") -ForegroundColor Green

    #Get the list items
	$list = $lists.GetByTitle($listTitle)
	$listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
    $fieldColl = $list.Fields
	$ctx.load($listItems)
    $ctx.load($fieldColl)
	$ctx.executeQuery()
    
    if ($listFields) #if you're passing a specific set of fields, in a specific order, process those
    {
         #Array to Hold List Items 
         $listItemCollection = @() 
         #Fetch each list item value to export to excel
         foreach($item in $listItems)
         {
            $exportItem = New-Object PSObject 

            Foreach ($field in $listFields)
            {
                    if($NULL -ne $item[$field])
                    {
                        #Expand the value of Person or Lookup fields
                        $fieldType = $item[$field].GetType().name
                        if (($fieldType -eq "FieldLookupValue") -or ($fieldType -eq "FieldUserValue"))
                        {
                            $fieldValue = $item[$field].LookupValue
                        }
                        elseif ($fieldType -eq "FieldUrlValue")
                        {
                            $fieldValue = $item[$field].Url
                        }
                        else
                        {
                            $fieldValue = $item[$field]   
                        }
                    }
                    $exportItem | Add-Member -MemberType NoteProperty -name $field -value $fieldValue
            }
            #Add the object with above properties to the Array
            $listItemCollection += $exportItem
         }
        #Export the result Array to CSV file
        $listItemCollection | Export-CSV $exportFile -NoTypeInformation
    }
    else #export all fields for the list
    {
         #Array to Hold List Items 
         $listItemCollection = @() 
         #Fetch each list item value to export to excel
         foreach($item in $listItems)
         {
            $exportItem = New-Object PSObject 
            Foreach($field in $fieldColl)
            {
                    if($NULL -ne $item[$field.InternalName])
                    {
                        #Expand the value of Person or Lookup fields
                        $fieldType = $item[$field.InternalName].GetType().name
                        if (($fieldType -eq "FieldLookupValue") -or ($fieldType -eq "FieldUserValue"))
                        {
                            $fieldValue = $item[$field.InternalName].LookupValue
                        }
                        elseif ($fieldType -eq "FieldUrlValue")
                        {
                            $fieldValue = $item[$field].Url
                        }
                        else
                        {
                            $fieldValue = $item[$field.InternalName]   
                        }
                    }
                    $exportItem | Add-Member -MemberType NoteProperty -name $field.InternalName -value $fieldValue
            }
            #Add the object with above properties to the Array
            $listItemCollection += $exportItem
         }
        #Export the result Array to CSV file
        $listItemCollection | Export-CSV $exportFile -NoTypeInformation
    }
    
}

#export the lists
foreach ($list in $listsToExport)
{
     #if we have a valid folder for export, use it, otherwise export to the current directory
     if ($exportFolder -and (Test-Path $exportFolder -PathType Container))
     {
         $filepath = Join-Path -Path $exportFolder -ChildPath ($list["listTitle"] + ".csv")
     }
     else
     {
         $filepath = ($list["listTitle"] + ".csv")
     }

     #export the lists
     if ($exportAllFields)
     {
         exportList -listTitle $list["listTitle"] -exportFile $filepath
     }
     else
     {
         exportList -listTitle $list["listTitle"] -listFields $list["listFields"] -exportFile $filepath
     }
} 