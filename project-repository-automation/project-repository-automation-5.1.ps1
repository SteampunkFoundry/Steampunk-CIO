param
(
    [Parameter(Mandatory=$false)]
    [object] $WebhookData
)

$JsonRequestBody = $WebhookData.RequestBody
$RequestBody = ConvertFrom-Json -InputObject $JsonRequestBody
Write-Verbose -Message $RequestBody.Name
Write-Verbose -Message $RequestBody.ChargeCode
Write-Verbose -Message $RequestBody.Portfolio
Write-Verbose -Message $RequestBody.Owners

# Parameters for targeting the Projects site
$SourceSiteURL = "https://sesolutionsinc.sharepoint.com/sites/Projects"
$SourceSiteName = "Projects"
$ListURL = "Lists/Projects"
$TemplateListName = "Template"
$MembersGroupDirectoryName = "Customer Deliverables"

# Parameters for targeting the mike-test-site
# $SourceSiteURL = "https://sesolutionsinc.sharepoint.com/sites/Michael-Kim-Test-Site"
# $SourceSiteName = "Michael-Kim-Test-Site"
# $ListURL = "Projects"
# $TemplateListName = "Shared Documents"
# $MembersGroupDirectoryName = "testProject2" 

# Login
Connect-PnPOnline -ClientId <ClientID> -Url $SourceSiteURL -Tenant "sesolutionsinc.onmicrosoft.com" -Thumbprint <Thumbprint>

#Create the document library
$Name = $RequestBody.Name
$ChargeCode = $RequestBody.ChargeCode
$Portfolio = $RequestBody.Portfolio
$Title = $Name+" ("+$ChargeCode+")"
#Duplicate check based on charge code
$SearchResult = Get-PnPListItem -List $ListURL -Query "<View><Query><Where><Eq><FieldRef Name='Charge_x0020_Number'/><Value Type='Text'>$ChargeCode</Value></Eq></Where></Query></View>"
if ($SearchResult) {
    Write-Verbose -Message "A project with duplicate charge code exists! Please verify that the repository for the project you are trying to create does not already exist. For any questions, please contact the SharePoint site administrator."
    Write-Verbose -Message "Exiting the script." 
    Exit
} else {
    Write-Verbose -Message "Creating the $($Title) document library"
    $CreateNewLibrary = New-PnPList -Title $Title -Url "$ChargeCode" -Template DocumentLibrary 
}


#Copy the folder structure to the new document library
$Items = Get-PnPListItem -List $TemplateListName
$TargetUrl = "/sites/$SourceSiteName/$ChargeCode"
$MatchURL = "/sites/$SourceSiteName/$TemplateListName/"
$NotMatchURL = "/sites/$SourceSiteName/$TemplateListName/.*/"
foreach($Item in $Items){
    if (($Item.fieldValues.FileRef -match $MatchURL) -and ($Item.fieldValues.FileRef -notmatch $NotMatchURL)){
    Write-Verbose -Message "Copying folder: $($Item.FieldValues.FileLeafRef)"
    $CopyTemplate = Copy-PnPFile -SourceUrl "$($Item.FieldValues.FileRef)" -TargetUrl $TargetUrl -Force
    }
}

#Add List
Write-Verbose -Message "Creating List..."
$AddProjectList = Add-PnPListItem -List "Projects" -Values @{
    "Title" = $Title;
    "Charge_x0020_Number" = $ChargeCode;
    "Repository_x0020_Location" = $TargetUrl+"/Forms/AllItems.aspx, "+$Title;
    "Sector" = $Portfolio
}

# Permission groups
$OwnersGroupName = $Title+' - Owners'
$MembersGroupName = $Title+' - Members'
Write-Verbose -Message "Creating groups..."
$CreateOwnersGroup = New-PnPGroup -Title $OwnersGroupName
$CreateMembersGroup = New-PnPGroup -Title $MembersGroupName
Write-Verbose -Message "Adding owner(s) to the owners group..."
$SplitOwners = $RequestBody.Owners.Split(", ")| where {$_}
foreach($Owner in $SplitOwners){
	$AddMember = Add-PnPGroupMember -LoginName $Owner -Group $OwnersGroupName
}
#Break Permission Inheritance of the List
$BreakInheritance = Set-PnPList -Identity $Title -BreakRoleInheritance
#Grant permission on list to Group
$SetPermission = Set-PnPListPermission -Identity $Title -AddRole "Full Control" -Group $OwnersGroupName
$SetPermission = Set-PnPListPermission -Identity $Title -AddRole "Full Control" -Group 'steampunk Projects Owners'
$SetPermission = Set-PnPListPermission -Identity $Title -AddRole "Edit" -Group 'steampunk Projects Members'
#Grant permission to Members Group for the appropriate directory 
$MembersGroupDirectory = Get-PnPListItem -List $Title -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='File'>$MembersGroupDirectoryName</Value></Eq></Where></Query></View>"
$SetMembersPermission = Set-PnPListItemPermission -List $Title -Identity $MembersGroupDirectory -AddRole "Edit" -Group $MembersGroupName

Write-Verbose -Message "Done."
