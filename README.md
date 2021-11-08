# Sharepoint site creation automator

For creating multiple communication sites automated with powershell.

### About
Automate SharePoint Online communication site creation with this tool.
Gets data from parameters.xlsx excel-file, and pushes to SPO via creation.ps1. Uses [Douglas Finke's ImportExcel](https://github.com/dfinke/ImportExcel) powershell-module. Creates sites with minimal settings.

### How-to:
1. Fill parameters.xlsx Doc with correct info and replace TENANT with your tenant name.
2. Install module [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel/7.3.0) via ```PS> Install-Module -Name ImportExcel```
3. Run creation.ps1
4. Insert sharepoint-admin credentials
5. All done
