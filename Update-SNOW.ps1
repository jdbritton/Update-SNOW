#==============================================================================
# Update-SNOW.ps1
# 
# Version: 1.1
# Created: 25/10/2021
# Modified: 29/12/2021
# Author: James D. Britton. [JDB/JAMBRI]
# Purpose: For, but not limited to, the Kinetic IT desktop
# team for ______ account. Intended to aid with common updates to
# many assets on SNOW.
#==============================================================================

<# 
==============================
Start Changelog
==============================
V0.1 JDB
-   Created initial script
V0.9 JDB
-   Removed hard-coded credentials, opting for the use of a stored credentials file
    "safely" stored in a user's Documents folder.
-   Added more functionality, and corrected a problem where location was changed
    before user, and when user was changed, location was set to whatever SNOW
    thought it should be.
V0.99 JDB
-   Switched to using AES encrypted/decrypted string for the credentials, reliant
    on a keyfile kept in the user's Documents folder.
V1.0 JDB
-   Added error-checking in case user makes a mistake entering the location.
-   Added query-snow functionality, can just query for serial, status, etc. now.
-   Ready for EUT.
V1.1 JDB
-   Quarantine Date calculation now ignores weekends.
-   Added stopwatch timer and notification that the operations have completed.
-   Amended script to use the Department's proxy, authenticated. Due to a change
    by the network team, the script breaks unless it goes through the proxy.
==============================
End Changelog
============================== 
#>

# Params.
[CmdletBinding()]
PARAM (
    [Parameter(Mandatory = $false)] [String] $AssetID,
    [Parameter(Mandatory = $false)] [Switch] $Quarantine,
    [Parameter(Mandatory = $false)] [String] $Location,
    [Parameter(Mandatory = $false)] [String] $AssignedUser,
    [Parameter(Mandatory = $false)] [Switch] $Installed,
    [Parameter(Mandatory = $false)] [Switch] $InStock,
    [Parameter(Mandatory = $false)] [Switch] $PendingDisposal,
    [Parameter(Mandatory = $false)] [Switch] $Query,
    [Parameter(Mandatory = $false)] [Switch] $Help,
    [Parameter(Mandatory = $false)] [String] $InputFile # If user gives neither -InputFile nor -AssetID, script will exit.
)

#================================================ MESSAGES =================================================#

$NoParameterError = `
    "No switch parameter provided. Must select an operation to perform.
Please specify either an asset ID, or the full path to a text file for bulk input. Examples:
EG: .\Update-SNOW -AssetID DC101101A -Quarantine
EG: .\Update-SNOW -InputFile C:\temp\assets.txt -Quarantine
EG: .\Update-SNOW -AssetID DC101101A -AssignedUser `"example.person@example.ab.cde.fg`"
EG: .\Update-SNOW -AssetID DC101101A -Location `"EAST Town - 2nd Floor, 123 Fake Street`"
EG: .\Update-SNOW -AssetID DC101101A -Installed
EG: .\Update-SNOW -AssetID DC101101A -PendingDisposal
EG: .\Update-SNOW -AssetID DC101101A -Query
NOTE! You can run: '.\Update-SNOW.ps1 -Help' for quick-help OR 'Get-Help .\Update-SNOW.ps1' to see detailed help.`n"

$HelpMessage = `
    "Usage:
To query SNOW for some information about an asset --
Replacing DC101101A with the asset in question:
EG: .\Update-SNOW -AssetID DC101101A -Query

To set a single asset to quarantined --
Replacing DC101101A with the asset in question:
EG: .\Update-SNOW -AssetID DC101101A -Quarantine

To assign a single asset to a user --
Replacing DC101101A with the asset in question and the EMAIL ADDRESS with the user in question:
EG: .\Update-SNOW -AssetID DC101101A -AssignedUser `"example.person@example.ab.cde.fg`"

To assign a single asset to a location --
Replacing DC101101A with the asset in question and the LOCATION NAME AS IT APPEARS IN SNOW in question:
EG: .\Update-SNOW -AssetID DC101101A -Location `"EAST Town - 2nd Floor, 123 Fake Street`"
Note: a spelling mistake will BLANK the Location field, so, again: TYPE THE LOCATION NAME AS IT APPEARS IN SNOW

To set a single asset to Installed and In Use --
Replacing DC101101A with the asset in question:
EG: .\Update-SNOW -AssetID DC101101A -Installed

To set a single asset to Pending Disposal and clear all the fields required to mark an asset as pending disposal --
Replacing DC101101A with the asset in question:
EG: .\Update-SNOW -AssetID DC101101A -PendingDisposal

To apply the same operations to all assets in an input file --
Replacing the path to the input file to your file, and the operations you want:
EG: .\Update-SNOW -InputFile C:\temp\assets.txt -Quarantine
Note: the contents of an input file should look like this. No blank lines.
DC101101A
DCP103102A
H023456
(etc)`n"

$MissingCredentialsMsg = `
    "`aNo credentials file found, or the file exists, but is invalid.
Please place the correct credentials file in the Documents folder of the account
with which you are running this script. Such as:

C:\users\person.adm\Documents\[keyfile]

Retrieve the credentials file from the location outlined in the KB. You will only have to do
this once per computer you run this script on, unless you delete or modify the credentials file.`n"

$ConflictingOperationsMsg = `
    "`aYou have selected parameters for conflicting operations. -Installed, -Quarantine, -PendingDisposal 
and -InStock overwrite each other, and cannot be selected together. However, -Location and -AssignedUser
do not conflict. You can call -Location, -AssignedUser and one of the others together, too.
Please try again with non-conflicting parameters.`n"

$ErrorMessage = `
    "`aSomething went wrong! If you are sure the asset's ID is spelled correctly, it exists in SNOW, and that 
you've used this tool correctly and there's no other obvious reason for this error, contact the developer
with information and a sufficiently generous bribe (I will accept cookies) for assistance.`n"

$LocationInvalid = `
    "`a`nResponse content suggests there was an error when entering the location field, 
please double-check and ensure you input the location field exactly as it appears in SNOW, 
enclosed in double-quotes: `"EAST Town - 123 Fake Street`"
NOT: `"East Town - 123 Fake Road`" OR `"EAST Town - 123 Fake Stret`". 

Failure to get this right will blank out the location field for Service-NOW.`n"

$PossibleLocationErrorMsg = `
    "`nWarning: a possibly incorrect location string was detected during the running of this script. 
Review the output and correct this if so. Otherwise, a CI will have its location blanked out, and the CI record will be incorrect!

Please fix this ASAP: re-run this script with the location exactly as it appears in SNOW!`n"

$SpokenWarning = `
    "Location Warning: You may have entered an invalid location.
Please double check, and re-run the command with the location as it appears in Service-NOW!"

#================================================ FUNCTIONS =================================================#
# Here we define the function that does the actual work of invoke REST API and updating the CI.
# Ignore the "unused variable errors" if you're using VS Code, those variables are used. They're in scope. 


Function Get-ConfigurationItem() {
    Add-LogEntry -LogLevel Info -LogEntry "$env:USERNAME submitted a query for $AssetID."
    # Get the CI's sysid
    $GetQuery_Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $GetQuery_Headers.Add('Authorization', ('Basic {0}' -f $base64AuthInfo))
    $GetQuery_Headers.Add('Accept', 'application/json')
    $GetQuery_Uri = "https://example.service-now.com/api/now/table/cmdb_ci_computer?sysparm_query=nameIN$($assetID)&sysparm_fields=sys_id&sysparm_limit=10"
    
    $GetQuery_Method = "GET"
    $GetQuery_Response = Invoke-RestMethod -Headers $GetQuery_Headers -Method $GetQuery_Method -Uri $GetQuery_Uri -Proxy $ProxyURI -ProxyUseDefaultCredentials
    $sysid = $GetQuery_response.result.sys_id
    
    $GetReqHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $GetReqHeaders.Add('Authorization', ('Basic {0}' -f $base64AuthInfo))
    $GetReqHeaders.Add('Accept', 'application/json')
    $GetReqHeaders.Add('Content-Type', 'application/json')
    
    # Endpoint uri - Pass the "sys_id" retrieved from GET command to the highlighted area. This needs to be run for each CI.
    $GetReq_URI = "https://example.service-now.com/api/now/table/cmdb_ci_computer/$($sysid)?sysparm_fields=serial_number%2Cname%2Cram%2Cwarranty_expiration%2Ccpu_name%2Ccomments%2Chardware_status%2Chardware_substatus"
    
    # Specify HTTP method
    $GetReq_Method = "GET"

    Try {
        Write-Host -Object "Returning information for: $AssetID"

        $Response = Invoke-RestMethod -Headers $GetReqHeaders -Method $GetReq_Method -Uri $GetReq_URI -Proxy $ProxyURI -ProxyUseDefaultCredentials
        $AssetInfo = [PSCustomObject]@{
            Name               = $Response.result.name
            Serial             = $Response.result.serial_number
            WarrantyExpiration = $($Response.result.warranty_expiration | Get-Date -format "dd/MM/yy" -ErrorAction SilentlyContinue)
            HardwareStatus     = $Response.result.hardware_status
            HardwareSubstatus  = $Response.result.hardware_substatus
            Comments           = $Response.result.comments
        }
    }
    Catch {
        Write-Host -Object "An error occurred while processing the request for asset: $AssetID" -BackgroundColor Black -ForegroundColor Red
        Write-Host -Object "Was attempting to query." -BackgroundColor Black -ForegroundColor Red
        Write-Host -Object "$ErrorMessage" -BackgroundColor Black -ForegroundColor Magenta
        Add-LogEntry -LogLevel Error -LogEntry "$env:USERNAME suffered an error when querying for asset: $AssetID | $($error[0])"
        Try { $Response.result } Catch {}
    }

    Return $AssetInfo

}


Function Send-PatchRequest() {
    Try {
        Write-Output "Sending patch request for asset: $AssetID"
        $PatchReq_Response = Invoke-RestMethod -Headers $PatchReq_Headers -Method $PatchReq_Method -Uri $PatchReq_URI -Body $Body -Proxy $ProxyURI -ProxyUseDefaultCredentials
        # Print response result. Comment this out if we don't want the response result to appear.
        # $PatchReq_Response.result

        # Only if we're updating the location: if the Response contains part of the exact string the user entered, chances are they stuffed up.
        # A correctly entered location returns a LINK to the site record, not the string the user entered.
        # Said string is truncated to 33 characters in the response, hence the use of Substring(). We check for the first 8 characters.
        If ($Location) {
            $PatchReq_Response.result
            If ($PatchReq_Response.result.location | Select-String $Location.Substring(0, 7)) {
                Write-Warning -Message $LocationInvalid
                Write-Host -Object "Possible incorrect location string.`nLocation entered: $Location -- is this correct?" -ForegroundColor Black -BackgroundColor Red
                $Script:PossibleLocationError = $True # Must be <Script> scope to trigger the warning at the end.
                Add-LogEntry -LogLevel Warning -LogEntry "$env:USERNAME set the location for $AssetID to: $Location | Response content indicates this is not a valid location."
            }
        }
    }
    Catch {
        Write-Host -Object "An error occurred while processing the request for asset: $AssetID" -BackgroundColor Black -ForegroundColor Red
        Write-Host -Object "Body of attempted request: $Body" -BackgroundColor Black -ForegroundColor Red
        Write-Host -Object "$ErrorMessage" -BackgroundColor Black -ForegroundColor Magenta
        Add-LogEntry -LogLevel Error -LogEntry "$env:USERNAME suffered an error running an operation for asset: $AssetID | Body of request: $Body. | $($error[0])"
        Try { $PatchReq_Response.result } Catch {}
    }
}

Function Update-ConfigurationItem() {
    $GetQuery_Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $GetQuery_Headers.Add('Authorization', ('Basic {0}' -f $Base64AuthInfo))
    $GetQuery_Headers.Add('Accept', 'application/json')
    $GetQuery_URI = "https://example.service-now.com/api/now/table/cmdb_ci_computer?sysparm_query=nameIN$($AssetID)&sysparm_fields=sys_id&sysparm_limit=10"

    # Specify HTTP method - this is to get the sys_id from the asset name, so it's a "GET".
    $GetQuery_Method = "GET"
    $GetQuery_Response = Invoke-RestMethod -Headers $GetQuery_Headers -Method $GetQuery_Method -Uri $GetQuery_URI -Proxy $ProxyURI -ProxyUseDefaultCredentials
    $sysid = $GetQuery_response.result.sys_id

    # Set headers to update the CI
    $PatchReq_Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $PatchReq_Headers.Add('Authorization', ('Basic {0}' -f $Base64AuthInfo))
    $PatchReq_Headers.Add('Accept', 'application/json')
    $PatchReq_Headers.Add('Content-Type', 'application/json')

    # Endpoint URI - Pass the "sys_id" retrieved from GET command to the highlighted area. This needs to be run for each CI.
    $PatchReq_URI = "https://example.service-now.com/api/now/table/cmdb_ci_computer/$($sysid)?sysparm_fields=serial_number%2Cname%2Cwarranty_expiration%2Ccomments%2Chardware_status%2Chardware_substatus"

    # Specify HTTP method (should be "Patch")
    $PatchReq_Method = "PATCH"

    # Specify request body and get it done! There's an if statement and code-block for each operation we want to perform.
    # Doing it this way sends more patch reqs, but it's pretty reliable.
    If ($Quarantine) {
        $Body = "{`"hardware_status`":`"Unavailable`",`"hardware_substatus`":`"Quarantined`",`"comments`":`"Asset quarantined until $QuarantineDate. (Updated by: $UserName)`"}"
        Add-LogEntry -LogLevel Info -LogEntry "$env:USERNAME submitted the following patch request for $AssetID | Body of request: $Body."
        Write-Host -Object "Setting $AssetID to: Unavailable, Quarantined, and setting a comment with date: $QuarantineDate."
        Send-PatchRequest
        Start-Sleep -Milliseconds 75 # Avoid triggering rate limitations with a delay.
    }

    If ($PendingDisposal) {
        $Body = "{`"hardware_status`":`"Unavailable`",`"assigned_to`":`"`",`"hardware_substatus`":`"Pending_Disposal`",`"comments`":`"Asset pending disposal. (Updated by: $UserName)`"}"
        Add-LogEntry -LogLevel Info -LogEntry "$env:USERNAME submitted the following patch request for $AssetID | Body of request: $Body."
        Write-Host -Object "Setting $AssetID to: Unavailable, Pending Disposal, clearing assigned user, and setting a comment indicating it is pending disposal."
        Send-PatchRequest
        Start-Sleep -Milliseconds 75
    }

    If ($AssignedUser) {
        $Body = "{`"assigned_to`":`"$AssignedUser`"}"
        Add-LogEntry -LogLevel Info -LogEntry "$env:USERNAME submitted the following patch request for $AssetID | Body of request: $Body."
        [Console]::WriteLine("Assigning asset to: $AssignedUser") # Why this instead of Write-Host/Output? Because I can.
        Send-PatchRequest
        Start-Sleep -Milliseconds 75
    } # User must be updated before location, or else SNOW will overwrite the location field.
        
    If ($Installed) {
        $Body = "{`"hardware_status`":`"Installed`",`"hardware_substatus`":`"In_Use`"}"
        Add-LogEntry -LogLevel Info -LogEntry "$env:USERNAME submitted the following patch request for $AssetID | Body of request: $Body."
        [Console]::WriteLine("Assigning $AssetID to: Installed, In Use.")
        Send-PatchRequest
        Start-Sleep -Milliseconds 75
    }

    If ($InStock) {
        $Body = "{`"assigned_to`":`"none`",`"hardware_status`":`"In_Stock`",`"hardware_substatus`":`"Used`"}"
        Add-LogEntry -LogLevel Info -LogEntry "$env:USERNAME submitted the following patch request for $AssetID | Body of request: $Body."
        Write-Host -Object "Setting $AssetID to: In stock, Used."
        Send-PatchRequest
        Start-Sleep -Milliseconds 75
    }
        
    If ($Location) {
        $Body = "{`"location`":`"$Location`"}"
        $Original_PatchReq_URI = $PatchReq_URI
        $PatchReq_URI = $PatchReq_URI + "%2Clocation" # We don't always want/need the location field in the patch/response.
        Write-Host -Object "Setting $AssetID's location to $Location."
        Add-LogEntry -LogLevel Info -LogEntry "$env:USERNAME submitted the following patch request for $AssetID | Body of request: $Body."
        Send-PatchRequest
        $PatchReq_URI = $Original_PatchReq_URI
        Start-Sleep -Milliseconds 75
    }
}

Function Add-LogEntry {
    Param([ValidateSet("Error", "Info", "Warning")][String]$LogLevel, [String]$LogEntry)
    If ($LogFileAvailable) {
        $TimeStamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
        $StreamWriter = New-Object System.IO.StreamWriter -ArgumentList ([IO.File]::Open($LogFile, "Append"))
        $StreamWriter.WriteLine("[$TimeStamp] - $LogLevel - $LogEntry")
        $StreamWriter.Close()
    }
}

#================================================ MAIN =================================================#

#TODO: Disables logging if log file unavailable; maybe it should exit instead and not allow unlogged operations?
$Logfile = "\\example.xyz.abc\Scripts\Logs\update_snow.log"
$Script:LogFileAvailable = Test-Path -Path $Logfile

If (!$LogFileAvailable) {
    Write-Warning -Message "Warning, unable to find log file. Logging disabled."
}

# No parameters.
If (
    !($Quarantine) -and 
    !($PendingDisposal) -and 
    !($Location) -and 
    !($Installed) -and 
    !($AssignedUser) -and 
    !($Query) -and 
    !($InStock)
) {
    Write-Warning -Message $NoParameterError
    Break
}

# No asset ID or input file.
If (
    !($AssetID) -and 
    !($InputFile)
) {
    Write-Warning -Message $NoParameterError
    Break
}

# User wants short help message.
If ($Help) {
    Write-Host -Object $HelpMessage -BackgroundColor Black -ForegroundColor Yellow
    Break
}

# Incompatible operations selected.
If (
    ($Quarantine -and $PendingDisposal) -or 
    ($Quarantine -and $Installed) -or 
    ($PendingDisposal -and $Installed) -or 
    ($PendingDisposal -and $InStock) -or 
    ($InStock -and $Installed) -or 
    ($InStock -and $Quarantine)
) {
    Write-Warning -Message "Conflicting operation parameters provided."
    Write-Host -Object $ConflictingOperationsMsg -BackgroundColor Black -ForegroundColor Yellow
    Break
}

# User has not forgotten an operation or asset ID/input file, and hasn't used the "Help" switch.
If (!$Help) {
    [void]([System.Reflection.Assembly]::LoadWithPartialName('System.Diagnostics.Stopwatch'))
    $StopWatch = [Diagnostics.Stopwatch]::StartNew()
    $ProxyURI = "http://proxy.abc.de.fgh.ijk:8080"
    
    #$ProxyCredentials = Get-Credential # Uncomment if the ProxyDefaultCredentials parameter fails to work in future.
    # $QuarantineLength = 14 #How long should the quarantine period be?
    # $QuarantineDate = (Get-Date).AddDays($QuarantineLength) | Get-Date -format "dd/MM/yy"

    # More advanced way of getting the quarantine date (10 working days instead of just 14 days)
    # All weekends start with the letter "s"; count forward from today + 1 day, ignore all days starting with "s"
    $CurrentDate = Get-Date; $QuarantineDate = $CurrentDate.AddDays(1).ToOADate()..$CurrentDate.AddDays(14).ToOADate() `
    | Where-Object { [DateTime]::FromOADate($_).ToString("ddd") -notmatch "^S" } `
    | Select-Object -First 10 `
    | ForEach-Object { [DateTime]::FromOADate($_).ToString("yyyy-MM-dd") } `
    | Select-Object -last 1 `
    | Get-Date -format "dd/MM/yy"

    $KeyFile = "$home\Documents\random_name_for_obscurity_&_security.hex" # Our decryption key. Change nothing if you don't know what you're doing.

    # Does the key-file exist?
    If (![System.IO.File]::Exists($KeyFile)) {
        Write-Warning -Message "No valid credentials file exists. Breaking."
        Write-Host -Object $MissingCredentialsMsg -BackgroundColor Black -ForegroundColor Magenta
        Write-Error -Message "Credentials file either isn't where it should be, or contains wrong data." -Category InvalidData
        Add-LogEntry -LogLevel Error -LogEntry "$env:USERNAME attempted to use this utility when no credentials key-file was present for them."
        Break
    }

    # Is it the right file?
    # For the curious: Changing this expected hash achieves nothing but wastes time. You need the right key + ciphertext.
    $ValidationHash = "[KEYFILE HASH]" # We expect this hash from the key file.
    $HashGenerator = [System.Security.Cryptography.HashAlgorithm]::Create("sha256")
    $KeyFileStream = [System.IO.File]::OpenRead((Resolve-Path $KeyFile)) # Pop open the file and read the contents.
    $FileHash = $HashGenerator.ComputeHash($KeyFileStream); $FileHash = [System.BitConverter]::ToString($FileHash) -replace '-', ''
    $KeyFileStream.Close(); $KeyFileStream.Dispose()

    If (!($FileHash -eq $ValidationHash)) {
        Write-Warning -Message "Credentials file exists, but has been modified."
        Write-Host -Object $MissingCredentialsMsg -BackgroundColor Black -ForegroundColor Magenta
        Write-Error -Message "Credentials file failed hash validation. The contents have been modified." -Category InvalidData
        Add-LogEntry -LogLevel Error -LogEntry "$env:USERNAME attempted to use this utility with an invalid key-file."
        Break
    }
    
    # And the ciphertext our AES key will decrypt. DO NOT CHANGE.
    $CredentialsCipherText = "[EXTREMELY LONG STRING REMOVED FOR SECURITY - CREATE YOUR OWN + THE KEY VIA MY AES SNIPPETS]"

    # Decrypt the password string - still change nothing if you don't know what you're doing!
    $Key = Get-Content $KeyFile
    $CipherToSecureString = Write-Output $CredentialsCipherText | ConvertTo-SecureString -Key $Key
    $Passthrough = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CipherToSecureString)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($Passthrough)
    $Credentials = $PlainPassword # Well, okay, it's actually base64 encoded still.
    $Base64AuthInfo = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($Credentials)) # Still b64...

    $UserName = $env:USERNAME

    If ($AssignedUser) {
        If (!($AssignedUser -imatch '@[\w].*')) {
            Write-Warning -Message "Assigned user should be the user's current email address to avoid`
        ambiguity between things like old accounts, or admin accounts. Are you sure you'd like to continue?"
            
            While (($ContinueAssignedUserNotEmail = Read-Host "Are you sure you'd like to continue? `n1 -- Continue using $AssignedUser `n2 -- Change to something else?`nYour choice") -ne "1") {
                Switch ($ContinueAssignedUserNotEmail) {
                    1 {  }
                    2 {
                        "Supply new (must be email address):"; $AssignedUser = (Read-Host -Prompt "Input preferred VALID email address (spaces, !()*=\#, will be removed!)") -replace '[!()*=\s\\#]', ''
                        Write-Host -Object "Proceeding with $AssignedUser."
                    }
                    Default { Write-Host "`n`nInvalid entry. Please choose again. Options are 1 or 2.`n`n" -ForegroundColor Red }
                }
            }
            
        }
    }

    # No input file:
    If (!$InputFile) {
        If ($Query) {
            Get-ConfigurationItem -AssetID $AssetID
        }
        Update-ConfigurationItem -AssetID $AssetID
    }

    # Input file has been specified:
    If ($InputFile) {
        $BulkAssets = Get-Content -Path $InputFile
        [Console]::WriteLine("Processing patch requests for $($BulkAssets.Count) assets.")
        If ($Query) {
            ForEach ($IndividualAsset in $BulkAssets) {
                $AssetID = $IndividualAsset
                Get-ConfigurationItem -AssetID $AssetID
            }
        }
        ForEach ($IndividualAsset in $BulkAssets) {
            $AssetID = $IndividualAsset
            Update-ConfigurationItem -AssetID $AssetID
        } #TODO: Parallelize this? Would probably only save a few seconds in normal use though.
    }

    # Did the script detect a possible screw-up earlier when handling a location change?
    # We let the user know once everything is processed and halt until they acknowledge it.
    If ($PossibleLocationError -eq $True) {
        [void]([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")) 
        [Microsoft.VisualBasic.Interaction]::Beep()
        Start-Job -Name LocationWarningSpeech -ScriptBlock {
            [void]([System.Reflection.Assembly]::LoadWithPartialName('System.Speech'))
            $SpeechObject = New-Object System.Speech.Synthesis.SpeechSynthesizer
            $SpeechObject.Speak($using:SpokenWarning)
        } | Out-Null
        Write-Host -BackgroundColor Black -ForegroundColor Red -Object $PossibleLocationErrorMsg
        $UserAcknowledged = [Microsoft.VisualBasic.Interaction]::MsgBox($PossibleLocationErrorMsg, "MsgBoxSetForeground,Exclamation", "Error: Invalid Location")
        If ($UserAcknowledged -eq "Ok") { $PossibleLocationError = $False; "`a"; Add-LogEntry -LogLevel Warning -LogEntry "$env:USERNAME acknowledged error message for possibly incorrect location: `"$Location`"" }
    }
    $StopWatch.Stop()
    [Console]::WriteLine("Completed. Operation time: $($StopWatch.Elapsed.TotalMilliseconds) milliseconds."); 
}

<#
.SYNOPSIS
Updates Service-NOW via REST API. Intended to be used for speeding up repetitive tasks related to
updating CIs for assets, such as during the warranty replacement/disposal project.
As it takes some time to search for an asset, then load its page, then make changes, then save,
using this instead will speed such mundane tasks up considerably especially if making changes to 
dozens or hundreds of assets simultaneously.

.DESCRIPTION
Updates Service-NOW via REST API. Intended to be used for speeding up repetitive tasks.
As it takes some time to search for an asset, then load its page, then make changes, then save,
this will speed such mundane tasks up considerably especially if making changes to dozens or hundreds
of assets simultaneously.

Can be used to update a single asset or, using the -InputFile switch parameter, take input from a file and
iterate through, updating each asset in turn.

.EXAMPLE
Update-SNOW -AssetID DC101101A -Query
To query SNOW for some information about an asset, replacing DC101101A with the asset in question.

.EXAMPLE
Update-SNOW -AssetID DC101101A -Quarantine
To set a single asset to quarantined, replacing DC101101A with the asset in question.

.EXAMPLE
Update-SNOW -AssetID DC101101A -AssignedUser "example.person@example.ab.cde.fg"
To assign a single asset to a user, replacing DC101101A with the asset in question and the EMAIL ADDRESS with the user in question.

.EXAMPLE
Update-SNOW -AssetID DC101101A -Location "EAST Town - 2nd Floor, 123 Fake Street"
To assign a single asset to a location, replacing DC101101A with the asset in question and the LOCATION NAME AS IT APPEARS IN SNOW in question.

.EXAMPLE
Update-SNOW -AssetID DC101101A -Installed
To set a single asset to Installed and In Use, replacing DC101101A with the asset in question.

.EXAMPLE
Update-SNOW -AssetID DC101101A -PendingDisposal
To set a single asset to Pending Disposal and clear all the fields required to mark an asset as pending 
disposal, replacing DC101101A with the asset in question.

.EXAMPLE
Update-SNOW -AssetID DC101101A -PendingDisposal -Location "EAST Town - 2nd Floor, 123 Fake Street"
To set a single asset to Pending Disposal and clear all the fields required to mark an asset as pending 
disposal, and also changing the location, replacing DC101101A with the asset in question.

.EXAMPLE
Update-SNOW -InputFile C:\temp\assets_to_modify.txt -PendingDisposal
To process many assets from an input file. Setting them all to "Pending Disposal"
Replacing C:\temp\assets_to_modify.txt with the location of the input file in question.

.INPUTS
The -InputFile parameter allows you to specify the path to a folder with multiple assets.
Update-SNOW -InputFile C:\temp\assets_to_modify.txt -Location "EAST Town - 2nd Floor, 123 Fake Street"
Update-SNOW -InputFile C:\temp\assets_to_modify.txt -AssignedUser "example.person@example.ab.cde.fg"

The input file should be a text file that just looks like this, no blank lines or spaces before or in 
the asset names:
DC101101A
DCP103102A
H023456
Etc.

.LINK
LinkedIn:
https://www.linkedin.com/in/james-britton-476481123/
GitHub:
https://github.com/jdbritton

.NOTES
This script was originally created by James Duncan Britton (JAMBRI),
who had a lot of fun while he did it. I like using Write-Host, I like the colours. Haters gonna hate.
#>
