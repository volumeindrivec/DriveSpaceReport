function Get-DriveSpaceReport {
<#
.SYNOPSIS
    Gets drive space on specified computers using PSRemoting.
.DESCRIPTION
    Gets drive space information using PSRemoting or straight
    WMI calls, depending on specified switches.
.PARAMETER ComputerName
    Computer name to run the function against.
.PARAMETER DriveType
    Indicates the drive type to be analyzed. Defaults to 3, which
    is Local Disk. All drive types from the Win32_LogicalDisk class
    are valid.
.PARAMETER V2
    Uses RPC instead of WS-MAN to be PowerShell V2 compatible.
    Under the hood, it's using Get-WmiObject versus Get-CimInstance.
.PARAMETER NoSql
    Tells the module not to dump the data to the SQL database.
.EXAMPLE
    Get-DriveSpaceReport -ComputerName localhost
    Simple drive space check using only computer name.
.EXAMPLE
    Get-DriveSpaceReport -ComputerName localhost -DriveType 2
    Check drive space using computer name and specifying a DriveType of 2.
.EXAMPLE
    Get-DriveSpaceReport -ComputerName dc2 -V2
    Check using PowerShell V2 compatibility.
.NOTES
    Version                 :  0.4
    Author                  :  @sukotto_san
    Disclaimer              :  If you run it, you take all responsibility for it.
#>

[CmdletBinding()]
param(
    [parameter(mandatory=$true)][string[]]$ComputerName,
    [int]$DriveType=3,
    [switch]$V2=$false,
    [switch]$NoSql=$false,
    [string]$SqlConnectionString
)

Begin {

    Write-Verbose "Begin Block"
    
    # Simply enumerating the computers to run against. May remove this loop in the future.
    foreach ($c in $ComputerName) {
        Write-Verbose "Function will run against this computer  :  $c"
    } # End foreach loop
    
} # End Begin block

Process {
    # Enumerate each computer in $ComputerName and get the required info.
    Write-Verbose "Process Block"
    foreach ($computer in $ComputerName) {
        Write-Verbose "Processing $computer"
        try{
            if ( $V2 -eq $True ) {
                Write-Verbose "Using Get-WmiObject calls; -V2 switch was used."
                $os = Get-WmiObject -ComputerName $computer -Class Win32_OperatingSystem -ErrorAction Stop -ErrorVariable OSError
                $disk = Get-WmiObject -ComputerName $computer -Class Win32_LogicalDisk -Filter "drivetype=$DriveType" -ErrorAction Stop -ErrorVariable DiskError
            }
            else {
                Write-Verbose "Using Get-CimInstance via PSRemoting; -V2 switch was NOT used."
                $os = Invoke-Command -ComputerName $computer -ScriptBlock { Get-CimInstance -ClassName Win32_OperatingSystem } -ErrorAction Stop -ErrorVariable OSError
                $disk = Invoke-Command -ComputerName $computer -ScriptBlock { param($dt) Get-CimInstance -ClassName Win32_LogicalDisk -Filter "drivetype=$dt" } -ArgumentList $DriveType -ErrorAction Stop -ErrorVariable DiskError
            }
       
        # Enumerate each drive in $disk. Specifically, this allows for the details on each drive if a computer has more than one.
        foreach ($drive in $disk){
            Write-Verbose "Processing $drive"
            $prop = @{
                'ComputerName' = $computer
                'Drive' = $drive.DeviceID
                'PctFree' = $drive.FreeSpace / $drive.Size
                'Free' = $drive.FreeSpace
                'Size' = $drive.Size
                'OSName' = $os.Caption
                'Date' = (Get-Date)
            }
            $object = New-Object -TypeName PSObject -Property $prop
            $object.PSObject.TypeNames.Insert(0,'Report.DriveSpaceInfo')
            if (-not $NoSql) { Write-Output $object  | Save-ReportData -ConnectionString $SqlConnectionString }
            Write-Output $object
        }
        }
     catch{
            Write-Warning "You done screwed up.  $computer is no con permiso."
     } # End Catch block
    }

} # End Process Block

End { Write-Verbose "End Block" } # End End Block

} # End Get-DriveSpaceReport function
 