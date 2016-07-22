. (Join-Path $PSScriptRoot .\Get-DriveSpaceReport.ps1)
. (Join-Path $PSScriptRoot .\Send-DriveSpaceReport.ps1)


#Import-Module -Name SQLReporting
#Get-DriveSpaceReport -ComputerName notaninja-dc,localhost -Verbose  | Save-ReportData -LocalExpressDatabaseName DriveSpaceReport





#Get-DriveSpaceReport -ComputerName $cn -V2 -Verbose | Send-DriveSpaceReport -Recipient $to -Sender $fr -EmailServer $srv -AsAttachment -Verbose