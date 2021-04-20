# Uncomment for execution and TLS1.2
# Set-ExecutionPolicy -ExecutionPolicy Unrestricted
# [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Required Variables
$dscMod = "Microsoft365DSC"
$Organization = Read-Host "Enter your organization in this format: domain.onmicrosoft.com"
$domain = $Organization.Split(".")[0]
$gloAd = (Get-Credential)
$mfaR = Read-Host "Does the global admin account require MFA (y or n)?"
$mfar.ToLower()
$workLoads = @("O365", "SC", "AAD", "Teams", "EXO", "Intune", "SPO", "OD")
$dtString = Get-Date -Format "MMddyyyyHHmm"
$month = Get-Date -Format "MMMM"
$spoAdmin = "https://$domain-admin.sharepoint.com"
$spoSite = "https://$domain.sharepoint.com"

# Check if DSC module is installed and current
if (!(Get-InstalledModule -Name $dscMod))
{
  Write-Output "The Microsoft365DSC Module is not installed on this machine."`n
  Write-Output "Let's install it now . . ."`n
  Install-Module -Name $dscMod
}
else
{
  Write-Output "The Microsoft365DSC Module is installed on this machine."`n
  Write-Output "Let's make sure the installed Microsoft365DSC Module is the latest version . . ."`n
  $psGV = Find-Module -Name $dscMod | Sort-Object Version -Descending | Select-Object Version -First 1
  $olVer = $psGV | Select-Object @{n = 'OnlineVersion'; e = { $_.Version -as [string] } }
  $olVString = $olVer | Select-Object OnlineVersion -ExpandProperty OnlineVersion
  $localVersion = (Get-InstalledModule -Name $dscMod).version

  if ($olVString -le $localversion)
  { Write-Output "The installed Microsoft365DSC Module is the latest version - $olVString"
  }
  else
  {
    Write-Output "The Current Release of the Microsoft365DSC Module needs to be installed"
    Write-Output "Your version is $localVersion and we are updating to the latest version - $olvstring"
    Install-Module -Name Microsoft365DSC -Force
  }
}

# Import Modules
Write-Output "Importing Modules"
Import-Module -Name $dscMod
Import-Module -Name Microsoft.Online.SharePoint.PowerShell -DisableNameChecking

if($mfaR -eq "n"){
    Write-Output "Authenticating with stored credentials"
    Write-Output "Connecting to AzureAD"
    Connect-AzureAD -Credential $gloAd
    Write-Output "Connecting to Exchange Online"
    Connect-ExchangeOnline -Credential $gloAd
    Write-Output "Connecting to Security & Compliance"
    Connect-IPPSSession -Credential $gloAd
    Write-Output "Connecting to Teams"
    Connect-MicrosoftTeams -Credential $gloAd
    Write-Output "Connecting to PNP Online"
    Connect-PnPOnline -Url $spoSite -Credentials $gloAd
    Write-Output "Connecting to SharePoint Online Service"
    Connect-SPOService  -Url $spoAdmin -Credential $gloAd
}else{
    Write-Output "Authenticating interactively"
    Write-Output "Connecting to AzureAD"
    Connect-AzureAD
    Write-Output "Connecting to Exchange Online"
    Connect-Exchangeonline
    Write-Output "Connecting to Security & Compliance"
    Connect-IPPSession
    Write-Output "Connecting to Teams"
    Connect-MicrosoftTeams
    Write-Output "Connecting to PNP Online"
    Connect-PnPOnline -Url $spoSite
    Write-Output "Connecting to SharePoint Online Service"
    Connect-SPOService  -Url $spoAdmin 
}
Write-Output "Finished making connections."


# Verify paths
Write-Output "Let's make sure the folders exist"
if (!(Test-Path -Path "$env:USERPROFILE\Documents\M365DSC" ))
{
  Write-Output "Creating the Base directory"
  New-Item "$env:USERPROFILE\Documents\M365DSC" -ItemType Directory
}

if (!(Test-Path -Path "$env:USERPROFILE\Documents\M365DSC\$($month)" ))
{
  Write-Output "Creating the $($month) directory"
  New-Item "$env:USERPROFILE\Documents\M365DSC\$($month)" -ItemType Directory
  Set-Location "$env:USERPROFILE\Documents\M365DSC\$($month)"
}
elseif (Test-Path -Path "$env:USERPROFILE\Documents\M365DSC\$($month)")
{
  Write-Output "Renaming the existing folder so it doesn't get overwritten"
  Rename-Item "$env:USERPROFILE\Documents\M365DSC\$($month)" -NewName "$env:USERPROFILE\Documents\M365DSC\$($month)2"

}

# Set root working folder
$baseDir = "$env:USERPROFILE\Documents\M365DSC\$month"

# Set Telemetry
Set-M365DSCTelemetryOption -Enabled $False

foreach ($wL in $workLoads)
{
  New-Item $baseDir\$wL -ItemType Directory
  Set-Location $baseDir\$wL
  Start-Transcript -Path "$baseDir\$wL\$($wL)-transcript-$($dtString).txt"
  Export-M365DSCConfiguration -GlobalAdminAccount $gloAd -GenerateInfo $true -Mode Full -Quiet -Workloads $wL -Path $("$baseDir\$wL") -FileName "$($wl)TenantConfig.ps1"
  New-M365DSCConfigurationToExcel -ConfigurationPath "$baseDir\$wL\$($wl)TenantConfig.ps1" -OutputPath "$baseDir\$($wL)Config.xlsx"
 
  # Close the open Excel workbook"
    Stop-Process -Name "excel"   
  
  # This converts the Excel xlsx output file to Csv
        $Sheet = "Report"
        $xlsx = "$baseDir\" + "$($wL)Config.xlsx"
        $csv = "$baseDir\" + "$($wL)Config.csv"

        $objExcel = New-Object -ComObject Excel.Application
        $objExcel.Visible = $False
        $objExcel.DisplayAlerts = $False
        $WorkBook = $objExcel.Workbooks.Open($xlsx)
        $WorkSheet = $WorkBook.sheets.item("$Sheet")

        $WorkBook.sheets | Select-Object -Property Name

        $xlCSV = 6
        $WorkBook.SaveAs($csv,$xlCSV)

        $objExcel.quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
  
  Set-Location $baseDir
  Stop-Transcript
}

## Combine Csv files into one
$CsvFiles = Get-ChildItem ("$baseDir\*") -Include *.Csv
$Excel = New-Object -ComObject Excel.Application 
$Excel.visible = $false
$Excel.sheetsInNewWorkbook = $CsvFiles.Count
$workbooks = $excel.Workbooks.Add()
$CsvSheet = 1
		
Foreach ($Csv in $Csvfiles)
		
{
$worksheets = $workbooks.worksheets
$CsvFullPath = $Csv.FullName
$SheetName = ($Csv.name -split "\.")[0]
$worksheet = $worksheets.Item($CsvSheet)
$worksheet.Name = $SheetName
$TxtConnector = ("TEXT;" + $CsvFullPath)
$CellRef = $worksheet.Range("A1")
$Connector = $worksheet.QueryTables.add($TxtConnector,$CellRef)
$worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
$worksheet.QueryTables.item($Connector.name).TextFileParseType  = 1
$worksheet.QueryTables.item($Connector.name).Refresh()
$worksheet.QueryTables.item($Connector.name).delete()
$worksheet.UsedRange.EntireColumn.AutoFit()
$CsvSheet++
		
}
		
$workbooks.SaveAs("$baseDir\$month", 51)
$workbooks.Saved = $true
$workbooks.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
		
#Remove the Csv files
Get-ChildItem ("$baseDir\*") -Include *.Csv |Remove-Item
