# Set the global error action preference to stop on errors
$ErrorActionPreference = "Stop"

<######################################################################################################################>
<#                                            IMPORT MODULES AND SETTINGS                                             #>                                           
<######################################################################################################################>
Import-Module .\Modules\AccessToken.psm1
$importSettings = Get-Content "setting.json" -Raw

# Assign value to related variables
$setting = $importSettings| ConvertFrom-Json
$clientId = $setting.ClientID
$clientSecret = $setting.ClientSecret
$tenantID = $setting.TenantID
$oneDriveUPN = $setting.OneDriveUPN
$siteIDs = $setting.SiteIDs
$modifiedDate = $setting.ModifiedDate

<######################################################################################################################>
<#                                                 GET ACCESS CODE                                                    #>                                           
<######################################################################################################################>

# Get Access token
$accessToken = Get-AccessToken -ClientID $clientId -ClientSecret $clientSecret -TenantID $tenantID

# Create header to send within https headers
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization",  "Bearer $accessToken")
$headers.Add('Prefer', '')

<######################################################################################################################>
<#                                             REFRESH TOKEN BEFORE EXPIRE                                            #>                                           
<######################################################################################################################>
# Set Timer to Refresh access token
# Create a timer that triggers before current access token expired
$timer = [System.Timers.Timer]::new(3400000) # interval is set to 3400000 milliseconds (or 56.67 minutes)

# Define the event handler
$refreshToken = Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action {
    $accessToken = Get-AccessToken -ClientID $clientId -ClientSecret $clientSecret -TenantID $tenantID
    $headers.Authorization = "Bearer $accessToken"
    Write-Output "******* New Token Issued *********"
}
# Register Event
$refreshToken

# Start the timer
$timer.Start()

<######################################################################################################################>
<#                                               COLLECT RESULTS AND STORE                                            #>                                           
<######################################################################################################################>
# OneDrive User ID and Drive ID
$driveID = (Invoke-RestMethod "https://graph.microsoft.com/v1.0/users/$oneDriveUPN/drive" -Method Get -Headers $headers).id
$userID = (Invoke-RestMethod "https://graph.microsoft.com/v1.0/users/$oneDriveUPN" -Method Get -Headers $headers).id

# Save all the collected, modified URL to a Folder Hierarchy Path and filtered data into this Array variable
$finalResult = [System.Collections.ArrayList]@()

# It Stores Arrays of FolderNames after splitting hierarchy of Folder Path
# Example:  ABC/Company/Document convert to {Folder0 = ABC, Folder1 = Company, Folder2 = Document}
$FolderCollections = @{}

# It Stores Arrays of Folder Paths from newPath variable in Set-Path Function
$FolderArrays = [System.Collections.ArrayList]@()

# It Stores key and value for Folder ID that retrieved from OneDrive
$FolderIDList = @{}


<######################################################################################################################>
<#                                         SET SHAREPOINT AND ONEDRIVE PATH                                           #>                                           
<######################################################################################################################>

# Get the Documents URL and Modify to create Hirarchy of Site/Library/Document
function Set-Path {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$Items,

        [Parameter(Mandatory=$true)]
        [string]$SiteName,

        [Parameter(Mandatory=$false)]
        [string]$SubSiteName = "Empty",

        [Parameter(Mandatory=$true)]
        [string]$ListName,

        [Parameter(Mandatory=$true)]
        [string]$SiteID,

        [Parameter(Mandatory=$false)]
        [string]$SubSiteID = "Empty",

        [Parameter(Mandatory=$true)]
        [string]$ListID
    )
    
    # Modify the DisplayNames to create folder Name for OneDrive Parent Folders
    $siteDisplayName = "Site-" + $SiteName
    $libraryDisplayName = "Library-" + $ListName
    $subSiteDisplayName = "SubSite-" + $SubSiteName
    
    # Iterate through each items and Modify the URL Path Structure
    foreach($item in $Items){
        $urldecode = [System.Web.HttpUtility]::UrlDecode($item.webUrl)
        $splitUrl = $urldecode.Split('/')
        $itemID = $item.id
        $lastFolderIndex = $splitUrl.Count - 2

        # Create a new Path for Site or Subsite items
        if($SubSiteID -eq "Empty"){
            # Select the file path except the FileName to get last parent folder of the file
            if($isRootSite){
                $newPath = [System.Collections.ArrayList]@($siteDisplayName, $libraryDisplayName)
                if(($lastFolderIndex-4) -ge 0) {
                    $newPath.AddRange($splitUrl[4..$lastFolderIndex])
                }
            } else {
                $newPath = [System.Collections.ArrayList]@($siteDisplayName, $libraryDisplayName)
                if(($lastFolderIndex-6) -ge 0){
                    $newPath.AddRange($splitUrl[6..$lastFolderIndex])
                } 
            }

            # Create Sharepoint Documents Endpoint, this is require for later when we move from Sharepoint to OndDrive
            $sharePointItemEndPoint = "https://graph.microsoft.com/v1.0/sites/$SiteID/lists/$ListID/items/$itemID/driveitem"

        } else {
            if($isRootSite){
                $newPath = [System.Collections.ArrayList]@($siteDisplayName, $subSiteDisplayName, $libraryDisplayName )
                if(($lastFolderIndex-5) -ge 0) {
                    $newPath.AddRange($splitUrl[5..$lastFolderIndex])
                }
            } else {
                $newPath = [System.Collections.ArrayList]@($siteDisplayName, $subSiteDisplayName, $libraryDisplayName)
                if(($lastFolderIndex-7) -ge 0) {
                    $newPath.AddRange($splitUrl[5..$lastFolderIndex])
                }
            }

            $sharePointItemEndPoint = "https://graph.microsoft.com/v1.0/sites/$SiteID/sites/$SubSiteID/lists/$ListID/items/$itemID/driveitem"
        }
        
        # Add each modified Folder path to FolderLists variable 
        $FolderArrays.Add($newPath)

        # This will require later that links to OneDrive Folder Path to transfer file from Sharepoint to Onedrive location
        $finalResult.Add(@{
            SharePointItemEndpoint = $sharePointItemEndPoint
            FolderIDKey = $newPath -join ''
        })
    }
}


<######################################################################################################################>
<#                                               RETRIEVE items                                                       #>                                           
<######################################################################################################################>

# This function will retrieve all the items within the Site or SubSite library
function Get-Items {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteID,

        [Parameter(Mandatory=$false)]
        [string]$SubSiteID = "Empty",

        [Parameter(Mandatory=$true)]
        [string]$ListID,

        [Parameter(Mandatory=$true)]
        [string]$SiteName,

        [Parameter(Mandatory=$true)]
        [string]$ListName,

        [Parameter(Mandatory=$false)]
        [string]$SubSiteName = "Empty"
    )

    $headers.Prefer = 'HonorNonIndexedQueriesWarningMayFailRandomly'

    # Graph API parameters that will filter contentType by Documents and Modified Date
    $graphApiParameters = '?$expand=fields&$filter=fields/ContentType eq ' + "'Document' and fields/Modified gt '$modifiedDate'" 

    # Assign graph endpoint of sharepoint items for Site or SubSite collection
    if($SubSiteID -eq 'Empty'){
        $itemsEndpoint = "https://graph.microsoft.com/v1.0/sites/$SiteID/lists/$ListID/items" + $graphApiParameters
    } else {
        $itemsEndpoint = "https://graph.microsoft.com/v1.0/sites/$SiteID/sites/$SubSiteID/lists/$ListID/items" + $graphApiParameters
    }
    
    # Get the items
    $items = Invoke-RestMethod $itemsEndpoint -Method "GET" -Headers $headers
    
    # If the library is empty or no documents ignore 
    if($null -eq $items.value[0]){return 0}
    
    if($SubSiteID -eq 'Empty'){
        Set-Path -Items $items.value -SiteName $SiteName -ListName $ListName -ListID $ListID -SiteID $SiteID
    } else {
        Set-Path -Items $items.value -SiteName $SiteName -SubSiteName $SubSiteName -ListName $ListName -SiteID $SiteID -SubSiteID $SubSiteID -ListID $ListID
    }
    
    # Get Next link to page throught the items
    $nextLink = $items.'@odata.nextLink'

    # Retrieve items in each page until the nextLink value is null or empty
    while ($null -ne $nextLink) {
        $nextItems = Invoke-RestMethod $nextLink -Method "GET" -Headers $headers
        if($SubSiteID -eq 'Empty'){
            Set-Path -Items $nextItems.value -SiteName $SiteName -ListName $ListName -ListID $ListID -SiteID $SiteID
        } else {
            Set-Path -Items $nextItems.value -SiteName $SiteName -SubSiteName $SubSiteName -ListName $ListName -SiteID $SiteID -SubSiteID $SubSiteID -ListID $ListID
        }
        
        $nextLink = $nextItems.'@odata.nextLink'
    }
}




<######################################################################################################################>
<#                                                     GET LISTS                                                      #>                                           
<######################################################################################################################>

# Retrieve all the lists in a site or Subsite and select only documentLibrary, After that iterate through each libraries to Get Documents
function Get-SiteLibraries {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteID,

        [Parameter(Mandatory=$false)]
        [string]$SubSiteID = "Empty",

        [Parameter(Mandatory=$true)]
        [string]$SiteName,

        [Parameter(Mandatory=$false)]
        [string]$SubSiteName = "Empty"
    )
    
    # If true get the Site collections lists or get the Subsite Collection lists
    if($SubSiteID -eq 'Empty'){
        $listsEndpoint = "https://graph.microsoft.com/v1.0/sites/$SiteID/lists"
    } else {
        $listsEndpoint = "https://graph.microsoft.com/v1.0/sites/$SiteID/sites/$SubSiteID/lists"
    }

    $response = Invoke-RestMethod $listsEndpoint -Method "GET" -Headers $headers
    $libraries = $response.value | where-object {$_.list.template -eq 'documentLibrary'}

    foreach($library in $libraries){
        $listName = $library.displayName
        if($SubSiteID -eq 'Empty'){
            Get-Items -SiteID $SiteID -ListID $library.id -SiteName $SiteName -ListName $listName
        } else {
            Get-Items -SiteID $SiteID -SubSiteID $SubSiteID -ListID $library.id -SiteName $SiteName -SubSiteName $SubSiteName -ListName $listName
        }
    }
}




<######################################################################################################################>
<#                                                   START Execution                                                  #>                                           
<######################################################################################################################>

# Iterate through each Site 
foreach($siteID in $siteIDs){
    [bool]$isRootSite = $false
    $getRootSite = Invoke-RestMethod "https://graph.microsoft.com/v1.0/sites/root" -Method "GET" -Headers $headers
    $site = Invoke-RestMethod "https://graph.microsoft.com/v1.0/sites/$siteID" -Method "GET" -Headers $headers
    $siteName = $site.displayName

    #Check if the site is the root site or not
    if($site.webUrl -eq $getRootSite.webUrl ){
        $isRootSite = $true      
    } 

    Get-SiteLibraries -SiteID $siteID -SiteName $siteName


    $subSites = Invoke-RestMethod "https://graph.microsoft.com/v1.0/sites/$siteID/sites" -Method "GET" -Headers $headers

    # Check if Subsite exist in Root site
    if($true -eq $subSites.value){
        foreach($subSiteID in ($subSites.value.id)){
            $subSite = Invoke-RestMethod "https://graph.microsoft.com/v1.0/sites/$siteID/sites/$subSiteID" -Method "GET" -Headers $headers
            $subSiteName = $subSite.displayName
            
            Get-SiteLibraries -SiteID $siteID -SubSiteID $subSiteID -SubSiteName $subSiteName -SiteName $siteName
        }
    }
}


<######################################################################################################################>
<#                                               COLLECT FOLDER, FILTER AND CREATE                                    #>                                           
<######################################################################################################################>

function Set-OneDriveFolder {
    $removeDuplicatePath = $FolderArrays | Select-Object -Unique

    # Determine the maximum depth of folder hierarchy (get Highiest length of an array)
    $maxDepth = ($removeDuplicatePath | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    
    # Seperate Hierarchy of the Path (eg, Cheese/Cake/Sweet, add 'Cheese' in Folder0 array, 'Cake' in Folder 1 and so on)
    for ($folderDepth = 0; $folderDepth -lt $maxDepth; $folderDepth++) {
        $FolderCollections.Add("Folder$folderDepth", [System.Collections.ArrayList]@())
        foreach ($folder in $removeDuplicatePath) {
            if ($folder.Count -gt $folderDepth) { # for example current folder is a/b/c 3 in array and the folderDepth is 5 it ignore the array
                $folderName = $folder[$folderDepth]
                $fullPath = $folder[0..$folderDepth]
                $FolderCollections."Folder$folderDepth".Add(@{ Name = $folderName; FullPath = $fullPath })
            }
        }   
    }

    # Keep only unique value in each Folder Array
    for ($folderDepth = 0; $folderDepth -lt $maxDepth; $folderDepth++) {
        $FolderCollections."Folder$folderDepth" = $FolderCollections."Folder$folderDepth" | Sort-Object -Property Name -Unique
    }

    # URL of the root endpoint of OneDrive
    $getFolderIDUrl = "https://graph.microsoft.com/v1.0/users/$userID/drive/root:"

    # Create Folder and Subfolder in OneDrive and get the Drive ID of each folder to transfer the file
    for ($folderDepth = 0; $folderDepth -lt $maxDepth; $folderDepth++) {
        foreach($folderLists in $FolderCollections."Folder$folderDepth"){
            # First Array of Folders in FolderCollection object is the root folder (eg, Site-C2conline)
            if($folderDepth -eq 0){
                $createFolderIDUrl = "https://graph.microsoft.com/v1.0/users/$userID/drive/root/children" 
            } else {
                $folderId = $FolderIDList.($folderLists.FullPath[0..($folderlists.FullPath.length - 2)] -join '')
                $createFolderIDUrl = "https://graph.microsoft.com/v1.0/users/$userID/drive/items/$folderID/children"
            }  

            # Creating a Property name for FolderIDList Hashtable or Object
            $folderIDName = $folderLists.FullPath -join ''

            # Get the OneDrive Folder ID, if already exist otherwise create folder and get the ID
            try {
                $folderPath = $folderLists.FullPath -join '/'
                $response = Invoke-RestMethod "$getFolderIDUrl/$folderPath" -Headers $headers -Method "GET"
            }
            catch {
                $errorString = $_.ToString()
                $hashtable = $errorString | ConvertFrom-Json
                if($hashtable.error.code -eq 'itemNotFound') {
                    $body = [PSCustomObject]@{
                        name = $folderLists.Name
                        folder = @{}
                    } | ConvertTo-Json
                $response = Invoke-RestMethod $createFolderIDUrl -Headers $headers -Method "POST" -Body $body -ContentType "application/json"
                }
            }
            # Add Property and Value
            $FolderIDList.Add($folderIDName, $response.id)

            # Display the progress bar for checking if folder exist or creating folders
            $percentComplete = ($folderDepth / $maxDepth) * 100
            Write-Progress -Activity "Maping Folders" -Status "$percentComplete%" -PercentComplete $percentComplete
        }
    }
}

# This Function starts transfer of the files
function Move-SharePointFiles {
    $headers.Prefer = "respond-async"
    $count = 1
    foreach($file in $finalResult){
        $folderIDKey = $file.FolderIDKey
        $itemID = $file.SharePointItemEndpoint        
        $body = [PSCustomObject]@{
            parentReference = [PSCustomObject]@{
                driveId = $driveID
                id = $FolderIDList."$folderIDKey"
            }  
        } | ConvertTo-Json

    Invoke-RestMethod $itemID -Method "PATCH" -Headers $headers -Body $body -ContentType "application/json"
    # Calculate Percentage
    $totalItems = $finalResult.Count
    $percentComplete = ($count / $totalItems) * 100
    
    # Display the progress bar
    Write-Progress -Activity "Processing items" -Status "$count of $totalItems" -PercentComplete $percentComplete
    $count += 1
    }
    Write-Progress -Activity "Processing items" -Status "Complted" -Completed
}

# Calling this function will get the Folder ID if already exit other wise it will create folder and get ID
Set-OneDriveFolder

# Finally this function will execute to start transfer the files from Sharepoint item location to OneDrive folder location
Move-SharePointFiles

# Clean up session
$timer.Stop()
Unregister-Event -SubscriptionId $refreshToken.Id
$timer.Dispose()
