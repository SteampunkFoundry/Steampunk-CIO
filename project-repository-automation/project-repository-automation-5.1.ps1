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

# Parameters
$SourceSiteURL = "https://sesolutionsinc.sharepoint.com/sites/Michael-Kim-Test-Site"

# Login
Connect-PnPOnline -ClientId <ClientID> -Url $SourceSiteURL -Tenant "sesolutionsinc.onmicrosoft.com" -Thumbprint <Thumbprint>

#Create the document library
$Name = $RequestBody.Name
$ChargeCode = $RequestBody.ChargeCode
$Portfolio = $RequestBody.Portfolio
$Title = $Name+" ("+$ChargeCode+")"
#Duplicate check based on charge code
$SearchResult = Get-PnPListItem -List "Projects" -Query "<View><Query><Where><Eq><FieldRef Name='Charge_x0020_Number'/><Value Type='Text'>$ChargeCode</Value></Eq></Where></Query></View>"
if ($SearchResult) {
    Write-Verbose -Message "A project with duplicate charge code exists! Please verify that the repository for the project you are trying to create does not already exist. For any questions, please contact the SharePoint site administrator."
    Write-Verbose -Message "Exiting the script." 
    Exit
} else {
    Write-Verbose -Message "Creating the $($Title) document library"
    $CreateNewLibrary = New-PnPList -Title $Title -Template DocumentLibrary
}


#Copy the folder structure to the new document library
$Items = Get-PnPListItem -List "Shared Documents"
$TargetUrl = "/sites/Michael-Kim-Test-Site/$($Name) $($ChargeCode)"
foreach($Item in $Items){
    if (($Item.fieldValues.FileRef -match '/sites/Michael-Kim-Test-Site/Shared Documents/testProject/') -and ($Item.fieldValues.FileRef -notmatch '/sites/Michael-Kim-Test-Site/Shared Documents/testProject/.*/')){
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
$GroupName = $Title+' - Owners'
Write-Verbose -Message "Creating owners group..."
$CreateGroup = New-PnPGroup -Title $GroupName
Write-Verbose -Message "Adding owner(s) to the group..."
$SplitOwners = $RequestBody.Owners.Split(", ")| where {$_}
foreach($Owner in $SplitOwners){
	$AddMember = Add-PnPGroupMember -LoginName $Owner -Group $GroupName
}
#Break Permission Inheritance of the List
$BreakInheritance = Set-PnPList -Identity $Title -BreakRoleInheritance -CopyRoleAssignments
#Grant permission on list to Group
$SetPermission = Set-PnPListPermission -Identity $Title -AddRole "Full Control" -Group $GroupName

Write-Verbose -Message "Done."
