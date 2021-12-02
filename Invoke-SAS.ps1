<#
.SYNOPSIS
Invoke-SAS runs a SAS program in a background job

.DESCRIPTION
Invoke-SAS runs a SAS program in a background job, providing regular feedback on its status.

The provided SAS program is submitted to a remote server.  After the program completes, the SAS log and listing are downloaded and written to a custom location.  The log and listing files have a datestamp appended to the filename to prevent overwriting existing log and listing files.

The intended behavior is akin to submitting a SAS program from the command line in a Linux environment.

In order for a SAS program to be submitted, the user needs to authenticate with SAS.  This is accomplished by passing the username and a password file location to the function.  In order to create the needed password file, the following can be run from a PowerShell prompt.

Read-Host "Enter Password:" -AsSecureString | ConvertFrom-SecureString | Out-File "C:\path\to\mypwd.txt"

Finally, the function utilizes SAS Integration Technologies which provides access to the SAS Integrated Object Model (IOM) through Microsoft Component Object Model (COM).  The easiest way to ensure SAS Integration Technologies are installed on the client machine is to have SAS Enterprise Guide installed.

.PARAMETER userName
The username to connect to the SAS server

.PARAMETER pwPath
The path to the password file (see DESCRIPTION if a password file needs to be created)

.PARAMETER sasWorkspaceServer
The fully qualified name of the SAS workspace server (not the metadata server)

.PARAMETER sasWorkspaceServerPort
The SAS workspace server port (default is 8591)

.PARAMETER sasCodePath
The path to the SAS code to be run

.PARAMETER outDirPath
The path to the directory to save the resulting SAS log and listing files

.PARAMETER reportEverySecs
The time in seconds to provide feedback to the console as to whether the SAS program is continuing to run (default is 60 seconds)

.EXAMPLE
Invoke-SAS -userName myusername -pwPath "C:\path\to\mypwd.txt" -sasCodePath "C:\path\to\sasprogram.sas" -outDirPath "C:\path\to\out" -reportEverySecs 30

.LINK
http://support.sas.com/rnd/itech/doc/dist-obj/winclnt/index.html

.NOTES
Version:    1.0
Author:     Curtis Alexander
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, HelpMessage = "The username to connect to the SAS server")]
    [Alias("user")]
    [string]$userName,
    [ValidateScript( {
            if ([System.IO.Path]::IsPathRooted($_)) {
                $pathToTest = $_
            }
            else {
                $pathToTest = [System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath $_))
            }

            if (-Not ($pathToTest | Test-Path)) {
                throw "Password file does not exist.  Please select a different file path for pwPath."
            }
            if (-Not ($pathToTest | Test-Path -PathType Leaf)) {
                throw "The pwPath argument must be a file."
            }
            return $true
        })]
    [Parameter(Mandatory = $true, HelpMessage = "The path to the password file (see DESCRIPTION if a password file needs to be created)")]
    [Alias("pw")]
    [System.IO.FileInfo]$pwPath,
    [Parameter(Mandatory = $true, HelpMessage = "The fully qualified name of the SAS workspace server (not the metadata server)")]
    [Alias("server")]
    [string]$sasWorkspaceServer,
    [Parameter(Mandatory = $false, HelpMessage = "The SAS workspace server port (default is 8591)")]
    [Alias("port")]
    [int]$sasWorkspaceServerPort,
    [ValidateScript( {
            if ([System.IO.Path]::IsPathRooted($_)) {
                $pathToTest = $_
            }
            else {
                $pathToTest = [System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath $_))
            }

            if (-Not ($pathToTest | Test-Path)) {
                throw "SAS file does not exist.  Please select a different file path for sasCodePath."
            }
            if (-Not ($pathToTest | Test-Path -PathType Leaf)) {
                throw "The sasCodePath argument must be a file."
            }
            return $true
        })]
    [Parameter(Mandatory = $true, HelpMessage = "The path to the SAS code to be run")]
    [Alias("sas")]
    [System.IO.FileInfo]$sasCodePath,
    [ValidateScript( {
            if ([System.IO.Path]::IsPathRooted($_)) {
                $pathToTest = $_
            }
            else {
                $pathToTest = [System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath $_))
            }

            if (-Not ($pathToTest | Test-Path)) {
                throw "Out directory does not exist.  Please select a different directory path for outDirPath."
            }
            if (-Not ($pathToTest | Test-Path -PathType Container)) {
                throw "The outDirPath argument must be a directory."
            }
            return $true
        })]
    [Parameter(Mandatory = $true, HelpMessage = "The path to the directory to save the resulting SAS log and listing files")]
    [Alias("out")]
    [System.IO.FileInfo]$outDirPath,
    [Parameter(Mandatory = $false, HelpMessage = "The time in seconds to provide feedback to the console as to whether the SAS program is continuing to run (default is 60 seconds)")]
    [Alias("secs")]
    [int]$reportEverySecs = 60
)

$sasCodeFile = Split-Path $sasCodePath -Leaf

$dt = (Get-Date).ToString("yyyyMMddHHmmss")
$logPath = Join-Path $outDirPath "${sasCodeFile}__${dt}.log"
$logPath = Join-Path $outDirPath "${sasCodeFile}__${dt}.log"
$lstPath = Join-Path $outDirPath "${sasCodeFile}__${dt}.lst"

# Record all output from PowerShell
# $pwshLogPath = Join-Path $outDirPath "${sasCodeFile}__${dt}.pwsh.log"
# Start-Transcript -Path $pwshLogPath -append

# Inner block to be used with Start-Job to submit actual SAS code
$sasSubmitBlock = {
    Param(
        [Parameter(Mandatory = $true, HelpMessage = "username")]
        [string]$userName,
        [Parameter(Mandatory = $true, HelpMessage = "password path")]
        [string]$pwPath,
        [Parameter(Mandatory = $true, HelpMessage = "SAS code path")]
        [string]$sasCodePath,
        [Parameter(Mandatory = $true, HelpMessage = "log path")]
        [string]$logPath,
        [Parameter(Mandatory = $true, HelpMessage = "lst path")]
        [string]$lstPath
    )

    # Setup new COM objects that are necessary for connecting to a SAS workspace server
    $objFactory = New-Object -ComObject SASObjectManager.ObjectFactoryMulti2
    $objServerDef = New-Object -ComObject SASObjectManager.ServerDef

    # Object Keeper is only necessary if use Asynchronous mode in CreateObjectByServer - again, that only makes the connection asynchronous and not the running of the code
    # $objKeeper = New-Object -ComObject SASObjectManager.ObjectKeeper

    # Assign the attributes of your workspace server
    # Be aware of potential load balancers if there are multiple workspace servers
    $objServerDef.MachineDNSName = $sasWorkspaceServer
    $objServerDef.Port = $sasWorkspaceServerPort
    $objServerDef.Protocol = 2      # 2 = IOM protocol

    # TODO: Allow for writing password to XML file.  As an example, below is how to write a password to an XML file.
    <# 
    # Prompt for username and password
    $user = Read-Host "Enter username"
    # Returns a System.Security.SecureString object
    $pw = Read-Host "Enter password" -AsSecureString
    $confirmpw = Read-Host "Re-enter password" -AsSecureString

    $pw_bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pw)
    $confirmpw_bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($confirmpw)

    try {
        $pw_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto($pw_bstr)
        $confirmpw_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto($confirmpw_bstr)

        if (-not ($pw_text -eq $confirmpw_text)) {
            Write-Host "The passwords do not match, please re-run the script"
            break
        }
    } finally {
        # even with the break statement above, the finally block is run to cleanup
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pw_bstr)
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($confirmpw_bstr)
        Remove-Variable pw_text
        Remove-Variable confirmpw_text
        Remove-Variable confirmpw
    }

    # Create a credential object
    $cred = New-Object System.Management.Automation.PSCredential ($user, $pw)

    # Write the credential object to disk
    $cred | Export-Clixml -Path "C:\some\path\cred.xml"
    #>

    # To create password, run the following from a PowerShell prompt
    # Read-Host "Enter Password:" -AsSecureString | ConvertFrom-SecureString | Out-File "C:\path\to\mypwd.txt"
    $password = Get-Content $pwPath | ConvertTo-SecureString
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password

    # Connect to the server
    # Return value is an OMI (Open Metadata Interface) handle
    try {
        # Create and connect to the SAS session
        $objSAS = $objFactory.CreateObjectByServer(
            "SASApp", # server name
            $true, # $true = Synchronous, $false = Aynchronous - refers to the connection and not the running of code
            $objServerDef, # server definition for Workspace
            $cred.GetNetworkCredential().UserName, # username
            $cred.GetNetworkCredential().Password # password
        )

        Write-Host "`nSAS code completed running on:" -NoNewline
        Write-Host "$($objServerDef.MachineDNSName)`n" -ForegroundColor Yellow
    }
    catch [System.Exception] {
        Write-Host "`nCould not connect to SAS Workspace server.  Error message received:" -NoNewline
        Write-Host "$($_.Exception)" -ForegroundColor Red
        exit -1
    }

    # Get the SAS object from the Object Keeper
    # Only necessary if use Asynchronous mode in CreateObjectByServer - again, that only makes the connection asynchronous and not the running of the code
    # $objSAS = $objKeeper.GetObjectByName("SASApp")

    # Get the Language Service object
    $objSASLS = $objSAS.LanguageService

    # Technically, the below COULD be used to make code submission asynchronous
    # BUT the recommendation at this time is to not use the Async property
    # INSTEAD, recommendation is to send code synchronously (blocking) but 
    #   sending the synchronous call to a different process (Start-Job) or
    #   a different thread (Runspace)
    # $objSASLS.Async = $true

    # Read the SAS code
    $sasCodeFileAsTxt = Get-Content $sasCodePath -Raw

    # Set SAS options
    #   validvarname = any  ==>  useful when batch submitting, as this sometimes is not the default option
    #   pagesize = max  ==>  simply removes page numbers from the log making it much easier to read
    $sasCodePreamble = "options validvarname=any pagesize=max ;"

    # Combine the preamble with the actual code
    $sasCodeFinal = "${sasCodePreamble}`n`n${sasCodeFileAsTxt}"

    # Submit SAS code for synchronous (blocking) execution
    $objSASLS.Submit($sasCodeFinal)

    # Flush log
    Write-Host "SAS log located at: " -NoNewline
    Write-Host "${logPath}" -ForegroundColor Yellow

    $log = ""
    do {
        $log = $objSASLS.FlushLog(10000) # flush 10,000 lines
        # Note that cannot use Out-File in PowerShell 5.1 due to lack of Encoding parameter
        # Out-File $log -FilePath $logPath -Append
        [System.IO.File]::AppendAllLines([string]$logPath, [string[]]$log)
    } while ($log.Length -gt 0)

    # Flush lst
    Write-Host "SAS lst located at: " -NoNewline
    Write-Host "${lstPath}" -ForegroundColor Yellow

    $lst = ""
    do {
        $lst = $objSASLS.FlushList(10000) # flush 10,000 lines
        # Note that cannot use Out-File in PowerShell 5.1 due to lack of Encoding parameter
        # Out-File $log -FilePath $lstPath -Append
        [System.IO.File]::AppendAllLines([string]$lstPath, [string[]]$lst)
    } while ($lst.Length -gt 0)

    $objSAS.Close()
}

# Submit the SAS code to run in the background using PowerShell jobs
# The SAS code is submitted synchronously within the script block BUT the job is submitted in a background job
# This allows for monitoring of the job
# Note that PowerShell jobs are quite heavy and could be substituted for other forms of concurrency
#   - Runspaces
#     - PoshRSJob
#     - Invoke-Parallel
#     - Split-Pipeline
#     - ForEach-Object -Parallel (available in PowerShell Core 7)
#   - Threads
#     - PSThreadJob

# An alternative API would return an actual job and then allow the user to manage the job
Write-Host "`nSubmitting SAS code located at: " -NoNewline
Write-Host "${sasCodePath}" -ForegroundColor Yellow

$job = Start-Job -Name "${sasCodeFile}" -ScriptBlock $sasSubmitBlock -ArgumentList $userName, $pwPath, $sasCodePath, $logPath, $lstPath

Write-Host "`nCode submitted, waiting for code to complete..."

# Wait until job completes, alerting the user it is still running every reportEverySecs
$stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
do {
    $totalSecs = [Math]::Round($stopWatch.Elapsed.TotalSeconds, 0)

    $randomColor = Get-Random -InputObject ([Enum]::GetValues([System.ConsoleColor]))
    Write-Host "...running ${sasCodeFile} for ${totalSecs} seconds" -ForegroundColor $randomColor

    $job | Wait-Job -Timeout $reportEverySecs | Receive-Job
} while (@("Completed", "Failed") -notcontains $job.State)

# Stop-Transcript
