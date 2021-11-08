#Define Config Variables
#$AdminCenterURL = "https://tarmodev001-admin.sharepoint.com"
#$SiteURL = "https://tarmodev001.sharepoint.com/sites/testiintra"
#$SiteTitle = "Test PS Created Site"
#$SiteOwner = "tarmodeviadmin@tarmodev001.onmicrosoft.com"
#$Template = "SITEPAGEPUBLISHING#0" #Communication Site template
#$Timezone = 59 #Set correct timezone (Helsinki)
 
#Get Credentials to connect
#$Cred = Get-Credential

Write-Host "Opening Excel file"
$filePath = $PSScriptRoot+'/auto.parameters.xlsx'
$excel = Open-ExcelPackage -Path $filePath
$WorkSheet = $excel.Workbook.Worksheets["parameters"]

Write-Host "Getting parameters from file"
$totalNoOfRecords = $Worksheet.Dimension.Rows
#$totalNoOfRecords -= 5 

#Define Config Variables
$adminCenterURL = $WorkSheet.Cells['A3'].Text
$siteURL = $WorkSheet.Cells['B3'].Text
$siteTemplate = $WorkSheet.Cells['C3'].Text
$timezone = $WorkSheet.Cells['D3'].Text
$siteOwner = $WorkSheet.Cells['B6'].Text
$sitePrefix = $WorkSheet.Cells['C6'].Text
    
if ($totalNoOfRecords -gt 0) {  
    #Connect to Tenant Admin    
    Write-Host "Connecting to tenant"  
   #### Connect-PnPOnline -URL $AdminCenterURL -Credential $Cred
        # Declare the starting positions first row and column names  
        $rowNo = 6 
        #Get values from excel file 
        for ($i =0; $i -le $totalNoOfRecords; $i++) {
            $cell = 'A' + $rowNo 
            $siteName = $WorkSheet.Cells[$cell].Text
            $siteFullName = $sitePrefix + $siteName 
            Write-Host "Solu ($cell) Data: ($siteFullName)"
            if($siteName.Length -gt 1){
                # For each row since row 6, get variables, and create a new site
                Try
                {
                    #Check if the site exists already
                ### $site = Get-PnPTenantSite | Where-Object {$_.Url -eq $siteURL}
                    If ($null -eq $site )
                    {
                        #Create new communication site
                        #New-PnPTenantSite -Url $siteURL -Owner $siteOwner -Title $siteName -Template $siteTemplate -TimeZone $timezone -RemoveDeletedSite
                        write-host "Site $($siteFullName) Created Successfully!" -foregroundcolor Green
                        write-host $siteName                        
                    }
                    else
                    {
                        write-host "Site $($siteFullName) exists already!" -foregroundcolor Yellow
                    }
                }
                catch {
                    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
                }
                $rowNo += 1    
            }             
            else{
                Break
            }
        } 
    } 
    Write-Host "All sites created, exiting"
    Close-ExcelPackage -ExcelPackage $excel -NoSave