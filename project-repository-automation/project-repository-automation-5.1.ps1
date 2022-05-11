param
(
    [Parameter(Mandatory=$false)]
    [object] $WebhookData
)

$JsonRequestBody = $WebhookData.RequestBody
$RequestBody = ConvertFrom-Json -InputObject $JsonRequestBody
write-output $RequestBody.Name
write-output $RequestBody.ChargeCode
write-output $RequestBody.Portfolio
write-output $RequestBody.Owners

# Parameters
$SourceSiteURL = "https://sesolutionsinc.sharepoint.com/sites/Michael-Kim-Test-Site"

# Login
Connect-PnPOnline -ClientId 2b3063f1-bfaa-4198-8cbd-815e4db7e8fb -Url $SourceSiteURL -Tenant "sesolutionsinc.onmicrosoft.com" -Thumbprint 4C3BB513DD1E75198C8127EC80024C0A463D2250

#Create the document library
$Name = $RequestBody.Name
$ChargeCode = $RequestBody.ChargeCode
$Portfolio = $RequestBody.Portfolio
$Title = $Name+" ("+$ChargeCode+")"
#Duplicate check based on charge code
$SearchResults = Submit-PnPSearchQuery -Query "Title:($($ChargeCode))" -MaxResults 1
if ($SearchResults.ResultRows[0]) {
    write-output "A project with duplicate charge code exists! Please verify that the repository for the project you are trying to create does not already exist. For any questions, please contact the SharePoint site administrator."
    write-output "Exiting the script."
    Exit
} else {
    write-output "Creating the $($Title) document library"
    $CreateNewLibrary = New-PnPList -Title $Title -Template DocumentLibrary
}

#Copy the folder structure to the new document library
$Items = Get-PnPListItem -List "Shared Documents"
$TargetUrl = "/sites/Michael-Kim-Test-Site/$($Name) $($ChargeCode)"
foreach($Item in $Items){
    if (($Item.fieldValues.FileRef -match '/sites/Michael-Kim-Test-Site/Shared Documents/testProject/') -and ($Item.fieldValues.FileRef -notmatch '/sites/Michael-Kim-Test-Site/Shared Documents/testProject/.*/')){
    write-output "Copying folder: $($Item.FieldValues.FileLeafRef)"
    $CopyTemplate = Copy-PnPFile -SourceUrl "$($Item.FieldValues.FileRef)" -TargetUrl $TargetUrl -Force
    }
}

#Add List
write-output "Creating List..."
$AddProjectList = Add-PnPListItem -List "Projects" -Values @{
    "Title" = $Title;
    "Charge_x0020_Number" = $ChargeCode;
    "Repository_x0020_Location" = $TargetUrl+"/Forms/AllItems.aspx, "+$Title;
    "Sector" = $Portfolio
}

# Permission groups
$GroupName = $Title+' - Owners'
write-output "Creating owners group..."
$CreateGroup = New-PnPGroup -Title $GroupName
write-output "Adding owner(s) to the group..."
$SplitOwners = $RequestBody.Owners.Split(", ")| where {$_}
foreach($Owner in $SplitOwners){
	$AddMember = Add-PnPGroupMember -LoginName $Owner -Group $GroupName
}
#Break Permission Inheritance of the List
$BreakInheritance = Set-PnPList -Identity $Title -BreakRoleInheritance -CopyRoleAssignments
#Grant permission on list to Group
$SetPermission = Set-PnPListPermission -Identity $Title -AddRole "Full Control" -Group $GroupName

write-output "Done."