param
(
    [Parameter(Mandatory=$false)]
    [object] $WebhookData
)

write-output "start"
write-output ("object type: {0}" -f $WebhookData.gettype())
write-output $WebhookData

#Manual string parsing, because powershell 7.1 brings invalid JSON from a webhook 
#See https://docs.microsoft.com/en-us/azure/automation/automation-runbook-types#known-issues---71-preview
$SplitData = $WebhookData.Split(" ")
write-output $SplitData
for ( $i = 0; $i -lt $SplitData.count; $i++ )
{
	if ($SplitData[$i] -eq '"Name":') {
		$NameRaw = $SplitData[$i + 1]
		$Name = $NameRaw.Split('"')[1]
		write-output $Name
	}
	if ($SplitData[$i] -eq '"ChargeCode":') {
		$ChargeCodeRaw = $SplitData[$i + 1]
		$ChargeCode = $ChargeCodeRaw.Split('"')[1]
		write-output $ChargeCode
	}
	if ($SplitData[$i] -eq '"Portfolio":') {
		$PortfolioRaw = $SplitData[$i + 1]
		$Portfolio = $PortfolioRaw.Split('"')[1]
		write-output $Portfolio
	}
}

#Parameters
$SourceSiteURL = "https://sesolutionsinc.sharepoint.com/sites/Michael-Kim-Test-Site"

#Connect to the source Site
Connect-PnPOnline -URL $SourceSiteURL -Interactive

#Create the document library
$Name = Read-Host "Please enter the name of the project"
$ChargeCode = Read-Host "Please enter the charge code of the project"
$Portfolio = Read-Host "Please specify the portfolio this project is under (CS&D or HC&J)"
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
