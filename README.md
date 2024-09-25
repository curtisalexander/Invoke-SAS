# Invoke-SAS

Run a SAS program from PowerShell.

## Description
Invoke-SAS runs a SAS program in a [background job](https://learn.microsoft.com/en-us/powershell/scripting/developer/cmdlet/background-jobs), providing regular feedback on its status.

The provided SAS program is submitted to a remote server.  After the program completes, the SAS log and listing are downloaded and written to a custom location.  The log and listing files have a datestamp appended to the filename to prevent overwriting existing log and listing files.

The intended behavior is akin to submitting a SAS program from the command line in a Linux environment.

## Usage
In order for a SAS program to be submitted, the user needs to authenticate with SAS.  This is accomplished by passing the username and a password file location to the function.  In order to create the needed password file, the following can be run from a PowerShell prompt.

```powershell
Read-Host "Enter Password:" -AsSecureString | ConvertFrom-SecureString | Out-File "C:\path\to\mypwd.txt"
```

## Requirements
The function utilizes [SAS Integration Technologies](https://support.sas.com/downloads/browse.htm?fil=&cat=56) which provides access to the SAS Integrated Object Model (IOM) through Microsoft Component Object Model (COM).  The easiest way to ensure SAS Integration Technologies are installed on the client machine is to have SAS Enterprise Guide installed.

:pencil: Note that even if SAS Enterprise Guide is installed, it does not need to be opened in order to run a SAS program.
