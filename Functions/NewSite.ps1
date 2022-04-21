Function NewSite {
    param
    (
        $Title = $(throw "Please Enter the Subsite Title!"),
        $SiteURL = $(throw "Please Enter the Subsite URL!"),
        $SiteOwner = $(throw "Please Enter the Site Owner!")
    )
    Try {
        #Check if the site exists already
        $site = Get-PnPTenantSite | Where-Object { $_.Url -eq $SiteURL }
        If ($null -eq $site ) {
            New-PnPSite -Type CommunicationSite -Url $SiteURL -Owner $SiteOwner -Title $Title -Wait
            write-host "Site $($Title) Created Successfully!" -foregroundcolor Green       
        }
        else {
            write-host "Site called '$($Title)' exists already!" -foregroundcolor Yellow
        }
    }
    catch {
        write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
    }
}