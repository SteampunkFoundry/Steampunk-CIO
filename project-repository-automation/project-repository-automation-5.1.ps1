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

# Parameters
$SourceSiteURL = "https://sesolutionsinc.sharepoint.com/sites/Michael-Kim-Test-Site"

Connect-PnPOnline -ClientId 2b3063f1-bfaa-4198-8cbd-815e4db7e8fb -Url $SourceSiteURL -Tenant "sesolutionsinc.onmicrosoft.com" -Thumbprint 4C3BB513DD1E75198C8127EC80024C0A463D2250

#Create the document library
$Name = $RequestBody.Name
$ChargeCode = $RequestBody.ChargeCode
$Portfolio = $RequestBody.Portfolio
$Title = $Name+" ("+$ChargeCode+")"
#Duplicate check based on charge code
$SearchResults = Submit-PnPSearchQuery -Query "($($ChargeCode))"
if ($SearchResults.ResultRows[0]) {
    Write-Host -ForegroundColor RED "A project with duplicate charge code exists! Please verify that the repository for the project you are trying to create does not already exist. For any questions, please contact the SharePoint site administrator."
    Write-Host -ForegroundColor RED "Exiting the script."
    Exit
} else {
    Write-Host "Creating the $($Title) document library"
    New-PnPList -Title $Title -Template DocumentLibrary
}

#Copy the folder structure to the new document library
$Items = Get-PnPListItem -List "Shared Documents"
$TargetUrl = "/sites/Michael-Kim-Test-Site/$($Name) $($ChargeCode)"
foreach($Item in $Items){
    if (($Item.fieldValues.FileRef -match '/sites/Michael-Kim-Test-Site/Shared Documents/testProject/') -and ($Item.fieldValues.FileRef -notmatch '/sites/Michael-Kim-Test-Site/Shared Documents/testProject/.*/')){
    Write-Host "Copying folder: $($Item.FieldValues.FileLeafRef)"
    Copy-PnPFile -SourceUrl "$($Item.FieldValues.FileRef)" -TargetUrl $TargetUrl -Force
    }
}

#Add List
Write-Host "Creating List..."
Add-PnPListItem -List "Projects" -Values @{
    "Title" = $Title; 
    "Charge_x0020_Number" = $ChargeCode;
    "Repository_x0020_Location" = $TargetUrl+"/Forms/AllItems.aspx, "+$Title;
    "Sector" = $Portfolio
}
