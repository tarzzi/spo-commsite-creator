<#
    Sharepoint Online Communication Site Generator
    v.1.5

    Uses Douglas Finke's ImportExcel powershell module
    https://github.com/dfinke/ImportExcel
#>

<# Import functions #>
Get-ChildItem -Path "$PSScriptRoot\Functions\*.ps1" -Recurse | ForEach-Object {
    . $_.FullName 
}
if ($null -eq (Get-InstalledModule -Name "ImportExcel" -ErrorAction SilentlyContinue)) {
    Write-Host "ImportExcel not installed, installing..."
    Install-Module -Name ImportExcel
}

#TODO Get templates from maby root folder? and provision them
#https://mattipaukkonen.com/2020/07/02/distribute-sharepoint-page-templates-with-pnp-powershell-and-provisioning-templates/

Write-Host "Opening Excel file"
$filePath = $PSScriptRoot + '/parameters.xlsx'
$excel = Open-ExcelPackage -Path $filePath
$WorkSheet = $excel.Workbook.Worksheets["parameters"]

Write-Host "Getting parameters from file"
$tenantName = $WorkSheet.Cells['D3'].Text
$adminCenterURL = -join ("https://", $tenantName, "-admin.sharepoint.com")
$siteURL = $WorkSheet.Cells['A3'].Text
$siteOwner = $WorkSheet.Cells['D5'].Text
$sitePrefix = $WorkSheet.Cells['D7'].Text
#$siteTemplate = $WorkSheet.Cells['C3'].Text
#$lcid = $WorkSheet.Cells['D3'].Text // Finnish lcid not supported?

Write-Host "Connecting to tenant" -foregroundcolor Green
try {
    #Connect-PnPOnline -URL $AdminCenterURL -Credential $Cred  
    Connect-PnPOnline -URL $AdminCenterURL -Interactive      
}
catch { 
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}  

# Site creation
$rowNo = 3
$endOfTable = $false


<#
while(a tai b on tavaraa)
päivitä cells
jos a luo sivu
jos ei a ja b on, luo alisivu kunnes b tyhjä
jos a ja b tyhjä, quit
#>

# MAIN
# Do to each level 1 site
Do{
    # fix me plz
    # Init parameters
    $cellA = 'A' + $rowNo 
    $cellB = "B" + $rowNo
    $cellC = 'C' + $rowNo
    $siteTitle = $WorkSheet.Cells[$cellA].Text
    $subsiteTitle = $WorkSheet.Cells[$cellB].Text
    $siteURL = $WorkSheet.Cells[$cellC].Text
    $siteFullName = $sitePrefix + $siteTitle

    # No data found
    if ("" -eq $siteTitle -and "" -eq $subsiteTitle) {
        Write-Host "All sites created, exiting..."
        $endOfTable = $true
        break
    }
    if ("" -eq $siteURL) {
        $siteURL = $siteTitle.ToLower()
        #$siteURL -replace '[/.,&:]',''
    }

    $siteFullURL = -join ("https://", $tenantName, ".sharepoint.com/sites/", $siteURL)

    # Create site
    Write-Host "Checking, if site $siteFullName already exists..."
    $site = Get-PnPTenantSite -Url $siteFullURL -ErrorAction SilentlyContinue  
    
    if ($null -ne $site) {
        Write-Host "Site already exists..."
    }
    else {
        NewSite -Title $siteFullName -SiteUrl $siteFullURL -SiteOwner $siteOwner
    }
    $rowNo += 1
    # Check next entry below
    $cellA = 'A' + $rowNo 
    $cellB = "B" + $rowNo
    $siteTitle = $WorkSheet.Cells[$cellA].Text
    $subsiteTitle = $WorkSheet.Cells[$cellB].Text

    # No level 1 site below, check subsites
    if ("" -eq $siteTitle) {
        $cellB = "B" + $rowNo
        $subsiteTitle = $WorkSheet.Cells[$cellB].Text
        
        if ("" -ne $subsiteTitle) {
            Connect-PnPOnline -Url $siteFullURL -Interactive
        }
        while ("" -ne $subsiteTitle) {
            $cellC = 'C' + $rowNo
            $subSiteURL = $WorkSheet.Cells[$cellC].Text
            Write-Host $subsiteURL
            if ("" -eq $siteURL) {
                $siteURL = $siteTitle -replace ' ','-'
                #$siteURL -replace '[/.,&:]',''
            }
            
            NewSubsite -SiteCollectionURL $siteFullURL -Title $subsiteTitle -URL $siteURL
            # Traverse cell down
            $rowNo += 1
            $cellB = "B" + $rowNo
            $subsiteTitle = $WorkSheet.Cells[$cellB].Text
        }
        
    }
    else {
    }

    $cellA = "A$($rowNo)"
    $cellB = "B$($rowNo)"
    $aclear = $WorkSheet.Cells[$cellA].Text
    $bclear = $WorkSheet.Cells[$cellB].Text
    if(("" -eq $aclear) -and ("" -eq $bclear)){
        # must be the end since nothing below, or at subsites
        Write-Host "All sites created..."
        $endOfTable = $true
        break
    }

    
}while (!$endOfTable) 

#  create site
#   go down
#   is empty?
#       go right
#           has info
#           create sub

# TODO url field empty ? change url to lowercase : get url from url field
# TODO [-HubSiteId <Guid>] ? 

Close-ExcelPackage -ExcelPackage $excel -NoSave
