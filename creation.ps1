#Get Credentials to connect
$Cred = Get-Credential

Write-Host "Opening Excel file"
$filePath = $PSScriptRoot+'/parameters.xlsx'
$excel = Open-ExcelPackage -Path $filePath
$WorkSheet = $excel.Workbook.Worksheets["parameters"]
Write-Host "Getting parameters from file"
$totalNoOfRecords = $Worksheet.Dimension.Rows

#Define Site Variables
$adminCenterURL = $WorkSheet.Cells['A3'].Text
$siteURL = $WorkSheet.Cells['B3'].Text
$siteOwner = $WorkSheet.Cells['B6'].Text
$sitePrefix = $WorkSheet.Cells['C6'].Text
#$siteTemplate = $WorkSheet.Cells['C3'].Text
#$lcid = $WorkSheet.Cells['D3'].Text // Finnish lcid not supported???
    
if ($totalNoOfRecords -gt 0) {  
    #Connect to Tenant Admin    
    Write-Host "Connecting to tenant" -foregroundcolor Green
    Connect-PnPOnline -URL $AdminCenterURL -Credential $Cred
    # Declare the starting positions first row and column names  
    $rowNo = 6 
    #Get values from excel file 
    for ($i =0; $i -le $totalNoOfRecords; $i++) {
        $cell = 'A' + $rowNo 
        $siteName = $WorkSheet.Cells[$cell].Text
        $siteFullName = $sitePrefix + $siteName
        #Debug
        Write-Host "Solu ($cell) Data: ($siteFullName)"
        if($siteName.Length -gt 1){
            Try{
            #Check if the site exists already
            $site = Get-PnPTenantSite | Where-Object {$_.Url -eq $siteURL}
                If ($null -eq $site ){
                    $siteURLfull = $siteURL + $siteFullName
                    New-PnPSite -Type CommunicationSite -Url $siteURLfull -Owner $siteOwner -Title $siteName
                    write-host "Site $($siteFullName) Created Successfully!" -foregroundcolor Green
                    $siteURLfull = ""            
                }
                else {
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
    Write-Host "All sites created, exiting..."
    Close-ExcelPackage -ExcelPackage $excel -NoSave