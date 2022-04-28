#Parameters
$SourceSiteURL = "https://sesolutionsinc.sharepoint.com/sites/Michael-Kim-Test-Site"

#Connect to the source Site
Connect-PnPOnline -URL $SourceSiteURL -Interactive

#Create the document library
$Name = Read-Host "Please enter the name of the project"
$ChargeCode = Read-Host "Please enter the charge code of the project"
$Portfolio = Read-Host "Please specify the portfolio this project is under (CS&D or HC&J)"
$Title = $Name+" ("+$ChargeCode+")"
Write-Host "Creating the $($Title) document library"
New-PnPList -Title $Title -Template DocumentLibrary

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
