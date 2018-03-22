# Fetching Email in PowerShell

Dominik Deák


## 1 Introduction

This repository contains sample code demonstrating how to take advantage of .NET libraries in PowerShell to facilitate email communications without using an actual email client.

**Warning:** This is a demo only. These scripts are not intended to be used directly in a production environment.


## 2 Instructions

### 2.1 Prerequisites

Before using the demo code, the following tools and libraries are required. You may skip this section if these prerequisites are already installed.

* PowerShell 6: <https://github.com/PowerShell/PowerShell/releases>
* OpenPOP.NET (SourceForge): <https://sourceforge.net/projects/hpop/files/>
* OpenPOP.NET (NuGet): <https://www.nuget.org/packages/OpenPop.NET/>

### 2.2 Configuration

1. Create a new email address dedicated for experimentation purposes. This way you don't need to worry about accidentally deleting important messages.

2. Download the OpenPOP.NET library, either from SourceForge, or NuGet (see links above).

3. Extract `OpenPop.dll` binary file from the `.zip` package into the `.\Binaries` subdirectory.

4. Edit the PowerShell script file `Source\FetchMail.ps1` and supply the incoming POP3 server configuration, including the credentials needed for authentication:

	```powershell
	# Incoming email configuration used for fetching messages
	$incomingUsername  = ""
	$incomingPassword  = ""
	$incomingServer    = ""
	$incomingPortPOP3  = 995   # Normally 110 (not secure), or 995 (SSL)
	$incomingEnableSSL = $true
	```

5. The `.\Temp` subdirectory is where the attachments will be saved in this demo. You may change the output location by editing `Source\FetchMail.ps1` and specifying a different path for `$tempBaseURL`.

## 3 Supporting Resources

* [Fetching Email in PowerShell](https://deaksoftware.com.au/articles/fetching_email_in_powershell/) - Main article


## 4 Legal and Copyright

Released under the [MIT License](./license.md).

Copyright 2018, DEAK Software