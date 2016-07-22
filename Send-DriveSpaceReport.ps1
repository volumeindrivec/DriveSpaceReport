function Send-DriveSpaceReport {
<#
.SYNOPSIS
    Creates and sends the drive space report to the specified user in HTML formatting.
.DESCRIPTION
    Takes objects generated from the Get-DriveSpaceReport function, formats
    it to HTML, and then sends it to the specified recipient.
.NOTES
    Version                 :  0.2
    Author                  :  @sukotto_san
    Disclaimer              :  If you run it, you take all responsibility for it.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true)]$Objects,
    [Parameter(Mandatory=$True)][string[]]$Recipient,
    [Parameter(Mandatory=$True)][string]$Sender,
    [Parameter(Mandatory=$True)][string]$EmailServer,
    [string]$SqlConnectionString,
    [switch]$AsAttachment = $false,
    $WarningThreshold=30,
    $CriticalThreshold=15
)



Begin{
    #Import-Module -Name C:\Scripts\Modules\SQLReporting
    Write-Verbose "Begin Block"
    Write-Verbose "Initializing object arrays"
    $NormalObjects = @()
    $WarningObjects = @()
    $CriticalObjects = @()
    if (-not $Objects) { $Objects = Get-ReportData -TypeName Report.DriveSpaceInfo -ConnectionString $SqlConnectionString }
}

Process{
    foreach ($Object in $Objects)
    {
      Write-Verbose "Processing Object: $Object"
      if($Object.Date.Date -eq (Get-Date).Date) {
        if ( $Object.PctFree -lt ($CriticalThreshold / 100) ) {
            Write-Verbose "Adding object to Critical Objects"
            $FormattedObject = Select-Object -InputObject $Object -Property ComputerName,Drive,@{n="PctFree";e={"{0:P0}" -f [System.Convert]::ToDouble($object.PctFree)}},@{n="Free";e={"{0:N2}" -f ($_.Free/1GB)}},@{n="Size";e={"{0:N2}" -f ($_.Size/1GB)}},@{n="DateCaptured";e={ $_.Date.ToString("MM/dd/yyyy")}}
            Write-Verbose "Formatted object:  $FormattedObject"
            $CriticalObjects += $FormattedObject
        }
        elseif ( $Object.PctFree -lt ($WarningThreshold / 100) ) {
            Write-Verbose "Adding object to Warning Objects"
            $FormattedObject = Select-Object -InputObject $Object -Property ComputerName,Drive,@{n="PctFree";e={"{0:P0}" -f [System.Convert]::ToDouble($object.PctFree)}},@{n="Free";e={"{0:N2}" -f ($_.Free/1GB)}},@{n="Size";e={"{0:N2}" -f ($_.Size/1GB)}},@{n="DateCaptured";e={ $_.Date.ToString("MM/dd/yyyy")}}
            $WarningObjects += $FormattedObject
        }
        else {
            Write-Verbose "Adding object to Normal Object"
            $FormattedObject = Select-Object -InputObject $Object -Property ComputerName,Drive,@{n="PctFree";e={"{0:P0}" -f [System.Convert]::ToDouble($object.PctFree)}},@{n="Free";e={"{0:N2}" -f ($_.Free/1GB)}},@{n="Size";e={"{0:N2}" -f ($_.Size/1GB)}},@{n="DateCaptured";e={ $_.Date.ToString("MM/dd/yyyy")}}
            $NormalObjects += $FormattedObject
        }
      } # End foreach loop
    }
}

End{
    
    # CSS - Doesn't format well with Windows version of Outlook due to Word being used as rendering engine
    $css = '<style>
            table { width:98%; }
            td { text-align:center; padding:5px; }
            th { background-color:blue; color:white; }
            h3 { text-align:center }
            h6 { text-align:center }
            </style>'

    Write-Verbose "End Block"
    Write-Verbose "Building HTML report"

    $CriticalHTML = $CriticalObjects | ConvertTo-Html -Fragment -PreContent "<h3>CRITICAL - Less than $CriticalThreshold% free</h3>" | Out-String
    $WarningHTML = $WarningObjects | ConvertTo-Html -Fragment -PreContent "<h3>WARNING - Less than $WarningThreshold% free</h3>" | Out-String
    $NormalHTML = $NormalObjects | ConvertTo-Html -Fragment -PreContent "<h3>NORMAL - More than $WarningThreshold% free</h3>" | Out-String
    $FooterHtml = ConvertTo-Html -Fragment -PostContent "<h6>This report was run from:  $env:COMPUTERNAME on $(Get-Date)</h6>" | Out-String
    
    Write-Verbose "Sending Email:
          Recipient   : $Recipient
          Sender      : $Sender
          EmailServer : $EmailServer"

    if ($AsAttachment){
        $Report = ConvertTo-Html -Body "$CriticalHTML $WarningHTML $NormalHTML $FooterHtml $css" | Out-File $env:TMP\drivespace.html
        Write-Verbose "$Report"
        Send-MailMessage -to $Recipient -From $Sender -Subject "Drive Space Report" -Body "Please find the attached drive space report." -Attachments $env:TMP\drivespace.html -SmtpServer $EmailServer
    }
    else{
        $Report = ConvertTo-Html -Body "$CriticalHTML $WarningHTML $NormalHTML $FooterHtml $css" | Out-String
        Write-Verbose "$Report"
        Send-MailMessage -to $Recipient -From $Sender -Subject "Drive Space Report" -BodyAsHtml $Report -SmtpServer $EmailServer
    }

    
}

}
