<#
Related project: script that runs at the end of a deployment task sequence (actually, a disposal task sequence?)
Placed here because I have nowhere else for it aside from another repository.
#>

# Params required
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]
    $assetID
)

# Document Variables
$OutputDirectory = "L:\"
$OutputFile = "$OutputDirectory\\DiskWipe-$assetID.txt"
$AggregatedOutputFile = "$OutputDirectory\\DiskWipeResults.txt"
$LocalLog = "X:\Windows\JobComplete.txt"
$SerialNo = Get-CimInstance win32_bios | Select-Object -ExpandProperty SerialNumber
$Vendor = Get-CimInstance win32_bios | Select-Object -ExpandProperty Manufacturer
$Memory = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1mb
$CPUName = Get-CIMInstance -Class Win32_Processor |Select-Object -ExpandProperty Name
$ComputerName = Get-CIMInstance -Class Win32_ComputerSystem |Select-Object -ExpandProperty Model
$DiskSize = [Math]::Round((Get-CIMInstance -Class Win32_DiskDrive | Measure-Object -Property Size -sum).sum /1gb)
$TimeStamp_Date = Get-Date -Format "dd/MM/yy"
$TimeStamp_Time = Get-Date -Format "hh:mm tt"

$user = "username.example"
$pass = "examplepassword1"

# Build auth header 
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $user, $pass)))

# Set proper headers
$getQuery_Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$getQuery_Headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
$getQuery_Headers.Add('Accept','application/json')

# Endpoint uri - Pass all the Ci names under "sysparam_query"
$getQuery_Uri = "https://example.service-now.com/api/now/table/cmdb_ci_computer?sysparm_query=nameIN$assetID&sysparm_fields=sys_id&sysparm_limit=10"

# Specify HTTP method
$getQuery_Method = "get"

# Send HTTP request
$getQuery_Response = Invoke-RestMethod -Headers $getQuery_Headers -Method $getQuery_Method -Uri $getQuery_Uri

# Save the result (the sys_id) to an variable
$sysid = $getQuery_response.result.sys_id

# Receipt and messages
$Receipt = `
"Vendor:                $Vendor
Model:                  $ComputerName
Serial:                 $SerialNo
Asset Number:           $assetID
CPU Type\Speed:         $CPUName
Memory:                 $($Memory)MB
HDD Size:               $($DiskSize)GB
"

$Message = `
"DoD 5220.22-M sanitization Wipe using MS SDELETE - 7 Passes.
Date Sanitized:     $TimeStamp_Date
Time Completed:     $TimeStamp_Time
"

# Set headers for the Update CI
$patchReq_Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$patchReq_Headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
$patchReq_Headers.Add('Accept','application/json')
$patchReq_Headers.Add('Content-Type','application/json')

# Endpoint uri - Pass the "sys_id" retrieved from GET command to the highlighted area. This needs to be run for each CI.
$patchReq_Uri = "https://example.service-now.com/api/now/table/cmdb_ci_computer/$($sysid)?sysparm_fields=sys_id%2Cname%2Chardware_status%2Ccomments"

# Specify HTTP method
$patchReq_Method = "patch"

# Specify request body
$bodyObj = @{hardware_status = "unavailable"
    substatus = "Pending Disposal"
    comments = $Receipt + "" + "" + "" + $Message }

$body = $bodyObj | ConvertTo-Json

# Send HTTP request
$patchReq_Response = Invoke-RestMethod -Headers $patchReq_Headers -Method $patchReq_Method -Uri $patchReq_Uri -Body $body

# Print response
$patchReq_Response.result

#======== Create Logs ========#

# Add entry to the CSV.
Try {
    Add-Content -Path "$OutputDirectory\AssetsWiped.csv" -Value "$assetID,$SerialNo,$TimeStamp_Date,$TimeStamp_Time"
    Add-Content -Path "$OutputDirectory\AssetsWiped_backup.csv" -Value "$assetID,$SerialNo,$TimeStamp_Date,$TimeStamp_Time"
}
Catch {
    Start-Sleep -Seconds 5 #try again after 5 seconds.
    Add-Content -Path "$OutputDirectory\AssetsWiped.csv" -Value "$assetID,$SerialNo,$TimeStamp_Date,$TimeStamp_Time"
    Add-Content -Path "$OutputDirectory\AssetsWiped_backup.csv" -Value "$assetID,$SerialNo,$TimeStamp_Date,$TimeStamp_Time"
}

# Creates Network Log File
Add-Content -Path $AggregatedOutputFile -Value "$Receipt
____________________________________________________________
"

# Creates Network Label for Machine
Add-Content -Path $OutputFile -Value "$Receipt
$Message"

# Creates Local Log file that displays at end of Process
Add-Content -Path $LocalLog -Value "$Receipt
$Message"

# Start a job that displays the receipt. We need to pass the $LocalLog variable with $using:LocalLog, it seems.
Start-Job -ScriptBlock {cmd /c notepad.exe $using:LocalLog} -Name "DisplayReceipt"

# Check if entry in CSV, if not, try again once more:
If (!((Get-Content -Path "$OutputDirectory\AssetsWiped.csv") -match $assetID)) {
    Try {
        Add-Content -Path "$OutputDirectory\AssetsWiped.csv" -Value "$assetID,$SerialNo,$TimeStamp_Date,$TimeStamp_Time"
        Add-Content -Path "$OutputDirectory\AssetsWiped_backup.csv" -Value "$assetID,$SerialNo,$TimeStamp_Date,$TimeStamp_Time"
    }
    Catch {
        Start-Sleep -Seconds 5
        Add-Content -Path "$OutputDirectory\AssetsWiped.csv" -Value "$assetID,$SerialNo,$TimeStamp_Date,$TimeStamp_Time"
        Add-Content -Path "$OutputDirectory\AssetsWiped_backup.csv" -Value "$assetID,$SerialNo,$TimeStamp_Date,$TimeStamp_Time"
    }
}

Start-Sleep -Seconds 5
Wait-Job -Name "DisplayReceipt"
Start-Sleep -Seconds 10 # Grace period before script finishes and asset restarts.
