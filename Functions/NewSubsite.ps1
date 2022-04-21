Function NewSubsite {
    param
    (
        $SiteCollectionURL = $(throw "Please Enter the Site Collection URL!"),
        $Title = $(throw "Please Enter the Subsite Title!"),
        $URL = $(throw "Please Enter the Subsite URL!")
        #$Template = $(throw "Please Provide the Site Template!")
    )
    Try {
            #Create subsite
            #New-PnPWeb -Title $Title -Url $URL -Template "STS#3"
            Add-PnPPage -Name $Title #-LayoutType Home  
            Write-Host "Created Subsite: " $Title -ForegroundColor Green
    }
    catch {
        write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
    }
}