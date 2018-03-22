<#---------------------------------------------------------------------------
   Using POP3 in PowerShell

   Dominik Deak <deak.software@gmail.com>, DEAK Software

   This sample code demonstrates how to take advantage of .NET libraries in
   PowerShell to facilitate email communications without using an actual
   email client.

   Prerequisites:

   Before using the demo code, the following tools and libraries are
   required:

   * PowerShell 6: <https://github.com/PowerShell/PowerShell/releases>
   * OpenPOP.NET (SourceForge): <https://sourceforge.net/projects/hpop/files/>
   * OpenPOP.NET (NuGet): <https://www.nuget.org/packages/OpenPop.NET/>

   Instructions:

   1. Download the OpenPOP.NET library, either from SourceForge, or NuGet.

   2. Extract OpenPop.dll binary file from the .zip package into the
      .\Binaries subdirectory.

   3. Supply the incoming email server configuration below, including the
      credentials needed for authentication.

   Supporting Resources:

   * Article: <https://deaksoftware.com.au/articles/using_pop3_in_powershell/>
   * GitHub: <https://github.com/DEAKSoftware/Using-POP3-in-PowerShell/>

   Legal and Copyright:

   Released under the MIT License. Copyright 2018, DEAK Software.

   Permission is hereby granted, free of charge, to any person obtaining a
   copy of this software and associated documentation files (the "Software"),
   to deal in the Software without restriction, including without limitation
   the rights to use, copy, modify, merge, publish, distribute, sublicense,
   and/or sell copies of the Software, and to permit persons to whom the
   Software is furnished to do so, subject to the following conditions:

   The above copyright notice and this permission notice shall be included
   in all copies or substantial portions of the Software.

   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
   OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
   THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
   LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
   FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
   DEALINGS IN THE SOFTWARE.
  ---------------------------------------------------------------------------#>

Write-Host "Using POP3 in PowerShell - Dominik Deak <deak.software@gmail.com>, DEAK Software" -ForegroundColor Yellow

<#---------------------------------------------------------------------------
   Configuration data.
  ---------------------------------------------------------------------------#>
# Path configurations
$openPopLibraryURL = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath( "..\Binaries\OpenPop.dll" )
$tempBaseURL       = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath( "..\Temp\" ) # Or use system temp dir via [System.IO.Path]::GetTempPath()

# Incoming email configuration used for fetching messages
$incomingUsername  = ""
$incomingPassword  = ""
$incomingServer    = ""
$incomingPortPOP3  = 995   # Normally 110 (not secure), or 995 (SSL)
$incomingEnableSSL = $true

<#---------------------------------------------------------------------------
   Construct an POP3 client for the specified host name and credentials.
  ---------------------------------------------------------------------------#>
function makePOP3Client
   {
   Param
      (
      [string] $server,
      [int] $port,
      [bool] $enableSSL,
      [string] $username,
      [string] $password
      )

   $pop3Client = New-Object OpenPop.Pop3.Pop3Client

   $pop3Client.connect( $server, $port, $enableSSL )

   if (!$pop3Client.connected)
      {
      throw "Unable to create POP3 client. Connection failed with server $server"
      }

   $pop3Client.authenticate( $username, $password )

   return $pop3Client
   }

<#---------------------------------------------------------------------------
   Fetch $count messages from the POP3 server and list them in the console.
  ---------------------------------------------------------------------------#>
function fetchAndListMessages
   {
   Param
      (
      [OpenPop.Pop3.Pop3Client] $pop3Client,
      [int] $count
      )

   $messageCount = $pop3Client.getMessageCount()

   Write-Host "Messages available:" $messageCount

   $messagesEnd = [math]::max( $messageCount - $count, 0 )

   Write-Host "Fetching messages" $messageCount "to" ($messagesEnd + 1) "..."

   for ( $messageIndex = $messageCount; $messageIndex -gt $messagesEnd; $messageIndex-- )
      {
      $incomingMessage = $pop3Client.getMessage( $messageIndex )

      Write-Host "Message $messageIndex`:"
      Write-Host "`tFrom:" $incomingMessage.headers.from
      Write-Host "`tSubject:" $incomingMessage.headers.subject
      Write-Host "`tDate:" $incomingMessage.headers.dateSent
      }
   }

<#---------------------------------------------------------------------------
   Save a single OpenPop.Mime.Message message class to the specified URL.
  ---------------------------------------------------------------------------#>
function saveMessage
   {
   Param
      (
      [OpenPop.Mime.Message] $incomingMessage,
      [string] $outURL
      )

   New-Item -Path $outURL -ItemType "File" -Force | Out-Null

   $outStream = New-Object IO.FileStream $outURL, "Create"

   $incomingMessage.save( $outStream )

   $outStream.close()
   }

<#---------------------------------------------------------------------------
   Fetch $count messages from the POP3 server and save them to the specified
   $tempURL.
  ---------------------------------------------------------------------------#>
function fetchAndSaveMessages
   {
   Param
      (
      [OpenPop.Pop3.Pop3Client] $pop3Client,
      [int] $count,
      [string] $tempURL
      )

   $messageCount = $pop3Client.getMessageCount()

   Write-Host "Messages available:" $messageCount

   $messagesEnd = [math]::max( $messageCount - $count, 0 )

   Write-Host "Saving messages" $messageCount "to" ($messagesEnd + 1) "..."

   for ( $messageIndex = $messageCount; $messageIndex -gt $messagesEnd; $messageIndex-- )
      {
      $uid = $pop3Client.getMessageUid( $messageIndex )

      $emailURL = Join-Path -Path $tempURL -ChildPath ($uid + ".eml")

      Write-Host "Saving message $messageIndex to:" $emailURL

      $incomingMessage = $pop3Client.getMessage( $messageIndex )

      saveMessage $incomingMessage $emailURL
      }
   }

<#---------------------------------------------------------------------------
   Save an attachment to the specified URL.
  ---------------------------------------------------------------------------#>
function saveAttachment
   {
   Param
      (
      [System.Net.Mail.Attachment] $attachment,
      [string] $outURL
      )

   New-Item -Path $outURL -ItemType "File" -Force | Out-Null

   $outStream = New-Object IO.FileStream $outURL, "Create"

   $attachment.contentStream.copyTo( $outStream )

   $outStream.close()
   }

<#---------------------------------------------------------------------------
   Fetch $count messages from the POP3 server and save their attachments to
   the specified $tempURL.
  ---------------------------------------------------------------------------#>
function fetchAndSaveAttachments
   {
   Param
      (
      [OpenPop.Pop3.Pop3Client] $pop3Client,
      [int] $count,
      [string] $tempURL
      )

   $messageCount = $pop3Client.getMessageCount()

   Write-Host "Messages available:" $messageCount

   $messagesEnd = [math]::max( $messageCount - $count, 0 )

   Write-Host "Saving attachments for messages" $messageCount "to" ($messagesEnd + 1) "..."

   for ( $messageIndex = $messageCount; $messageIndex -gt $messagesEnd; $messageIndex-- )
      {
      $uid = $pop3Client.getMessageUid( $messageIndex )

      $incomingMessage = $pop3Client.getMessage( $messageIndex ).toMailMessage()

      Write-Host "Processing message $messageIndex..."

      foreach ( $attachment in $incomingMessage.attachments )
         {
         $attachmentURL = Join-Path -Path $tempURL -ChildPath $uid | Join-Path -ChildPath $attachment.name

         Write-Host "`tSaving attachment to:" $attachmentURL

         saveAttachment $attachment $attachmentURL
         }
      }
   }


<#---------------------------------------------------------------------------
   Run the demo.
  ---------------------------------------------------------------------------#>
[Reflection.Assembly]::LoadFile( $openPopLibraryURL )

try {
   Write-Host "Connecting to POP3 server: $incomingServer`:$incomingPortPOP3"

   $pop3Client = makePOP3Client `
      $incomingServer $incomingPortPOP3 $incomingEnableSSL `
      $incomingUsername $incomingPassword

   Remove-Variable -Name incomingPassword

   fetchAndListMessages $pop3Client 10
   fetchAndSaveMessages $pop3Client 10 $tempBaseURL
   fetchAndSaveAttachments $pop3Client 10 $tempBaseURL

   Write-Host "Disconnecting from POP3 server: $incomingServer`:$incomingPortPOP3"

   if ($pop3Client.connected)
      {
      $pop3Client.disconnect()
      }

   $pop3Client.dispose()

   Remove-Variable -Name pop3Client
   }

catch { Write-Error "Caught exception:`n`t$PSItem" }
