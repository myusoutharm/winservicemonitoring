<#
    The purpose of this script is to monitor a specific Windows event, triggering a notification (email) upon detection.
    A Windows task scheduler will be configured to execute this script in response to the event.
    
    By default, the script monitors the following event:

    -Event ID: 1026
    -Event source: .NET Runtime
    -Keyword: Xpo.Svc.Agent (in Event Viewer, this is in the message field)

    THIS SCRIPT IS INTENDTED TO OPERATE WITHIN A TRUSTED, INTERNAL, AND NON-SHARED ENVIRONMENT. THEREFORE, CREDENTIALS ARE NOT ENCRYPTED.

    Copyright: Minghui Yu (myu@southarm.ca), 2024, South Arm Technology Services Ltd.
    License: GPL-3.0
#>

param (
    [Parameter(Mandatory=$false)]
    [bool]$setup = $false
)

# Construct the log filename
$logFile = "ServiceMonitor.log"
# Start a transcript to log all output
Start-Transcript -Path $logFile -Append

# Not in use at this moment
# Write a message to the log file, retry if necessary.
function Write-Log {
    param (
        [string]$Message,
        [string]$Path
    )
    $retryCount = 0
    $maxRetries = 5
    $delay = 1000 # 1 second

    while ($true) {
        try {
            Add-Content -Path $Path -Value $Message -ErrorAction Stop
            break
        } catch {
            if ($retryCount -ge $maxRetries) {
                Write-Error "Failed to write to log after several retries: $_"
                break
            } else {
                Start-Sleep -Milliseconds $delay
                $retryCount++
            }
        }
    }
}

# Send email notification using SMTP2GO API
# To-do: add parameters to send SMS notifications as well
function SendHttpPostRequest {
    param (
        [string]$url,
        [string]$body
    )

    try {
        $headers=@{}
        $headers.Add("accept", "application/json")
        $headers.Add("Content-Type", "application/json")
        #$response = Invoke-WebRequest -Uri $url -Method Post -Headers $headers -ContentType 'application/json' -Body $body
        #Changed to basic parsing as regular pasing throws an error that is Internet Explorer related (how come in 2024 IE is still a thing?)
        $response = Invoke-RestMethod -Uri $url -Method Post -Body $body -Headers $headers -UseBasicParsing
        if ($response.StatusCode -ge 400) {
            Write-Error "HTTP request failed with status code $($response.StatusCode)."
        }
    } catch {
        Write-Error "Failed to send HTTP POST request: $_"
    }
}

# Read and parse config file
<#
    Sample configuration file (config-servicemonitor.txt):
    
    to=["email1@yourdomain.com","email2@yourdomain.com"]
    sender=sender@yourdomain.com
    template_id=TEMPLATE_ID
    API_KEY=API_KEY
    eventID=1026
    eventSource=.NET Runtime
    keyword=xpo.svc.agent
#>
$configFile = "config-servicemonitor.txt"
$config = @{}
if (Test-Path $configFile) {
    try {
        Get-Content $configFile | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $config[$key.Trim()] = $value.Trim()
        }

        # Ensure all necessary keys are present
        $requiredKeys = @('to', 'sender', 'template_id', 'API_KEY', 'eventID', 'eventSource', 'keyword')
        $missingKeys = $requiredKeys | Where-Object { -not $config.ContainsKey($_) }

        if ($missingKeys.Count -gt 0) {
            throw "Missing configuration keys: $($missingKeys -join ', ')"
        }
         # Validate that eventID is an integer
        if (![int]::TryParse($config['eventID'], [ref]0)) {
            throw "Event ID is not an integer: $($config['eventID'])"
        }  
        $eventID = [int]$config['eventID']   
        } catch {
            Write-Error "Failed to parse config file or invalid configuration value: $_"
            return
    }
} else {
    Write-Error "Configuration file not found."
    return
}

# Define the SMTP2GO API key and JSON template
$apiKey = $config["API_KEY"] ##smtp2go API KEY. 
$to = $config["to"]
$emailSender = $config["sender"]
$templateId = $config["template_id"]
$eventid = $config["eventID"]
$eventsource = $config["eventSource"]
$keyword = $config["keyword"]
$sanitizedKeyword = $keyword -replace '[^a-zA-Z]', ''
# Need more testing on variable assignment; single vs double quotes.
$jsonTemplate = @"
{
    "api_key": "$apiKey",
    "to": $to,
    "sender": "$emailSender",
    "template_id": "$templateId"
}
"@

# Setup mode: create a scheduled task to monitor the event
if ($setup) {
    try {
        # Task name based on sanitized keyword
        $taskName = "${sanitizedKeyword}_Monitor"

        # Delete existing task if it exists
        Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false

        # Define the custom XML for the event filter
        # Unable to filter events by keyword (always throws an error), only by event ID and source. 
        $eventFilterXml = @"
<QueryList>
  <Query Id='0' Path='Application'>
    <Select Path='Application'>*[System[Provider[@Name='$eventsource'] and (EventID=$eventid) and TimeCreated[timediff(@SystemTime) &lt;= 3600000]]]</Select>
  </Query>
</QueryList>
"@

        # Create event trigger using CIM
        $CIMTriggerClass = Get-CimClass -ClassName MSFT_TaskEventTrigger -Namespace Root/Microsoft/Windows/TaskScheduler:MSFT_TaskEventTrigger
        $trigger = New-CimInstance -CimClass $CIMTriggerClass -ClientOnly
        $trigger.Subscription = $eventFilterXml
        $trigger.Enabled = $True

        # Get the directory of the currently executing executable
        # The assumption is the script will be complied into an executable. If running this ps1 script directly via PowerShell, the exeDirectory will be powershell.exe's directory.
        $exeDirectory = Split-Path -Parent -Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
        $exePath = Join-Path -Path $exeDirectory -ChildPath "ServiceMonitor.exe"

        # Create a task action to execute the script when triggered        
        $action = New-ScheduledTaskAction -Execute $exePath  -WorkingDirectory $exeDirectory
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable

        # Register the task
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
    }
    catch {
            Write-Error "Failed to set up the monitoring task: $_"
        }
    # Return or Exit? Perhaps Exit is better. But for the time being, return is used.
    return
}

# Retrieve the most recent event for the specified source and ID
# Only retrieve the most recent event: MaxEvents 1
$events = Get-WinEvent -LogName Application -MaxEvents 1 | Where-Object {
    $_.ProviderName -eq $eventsource -and
    $_.Id -eq $eventid
}

if ($events) {
    # Only one event is expected because this script is trigged when and only when the specific event occured.
    foreach ($event in $events) {
        $logMessage = "Time: $($event.TimeCreated), Event ID: $($event.Id), Message: $($event.Message)"
        # We want to know the time of the event, the event ID, and the message.
        Write-Error $logMessage

    }
} else {
    $noEventMessage = "No events found that match the keyword '$keyword'."
    # An event with the same ID and source may not tbe the event we are looking for.
    Write-Error $noEventMessage
}

# Filter these events to find one that contains the keyword in its message
try {
    $matchedEvent = $events | Where-Object { $_.Message -match $keyword }
} catch {
    Write-Error "Failed to filter events: $_"
    return
}

if ($null -eq $matchedEvent) {
    Write-Error "No events found that match the keyword '$keyword'."
    return
}

# Found a matching event
if ($null -ne $matchedEvent) {
    # Construct the filename for the last event ID
    $lastEventFile = "last_${sanitizedKeyword}-event.txt"   
    try {
        # Load the last event ID from the file
        # An integer is expected, but if the file doesn't exist or doesn't contain an integer, set $lastEventId to $null
        $lastEventId = [int] (Get-Content -Path $lastEventFile -ErrorAction Stop)
    }
    catch {
        # If the file doesn't exist or doesn't contain an integer, set $lastEventId to $null
        $lastEventId = $null
    }
    
    # If the event is not the same as the last event, send a notification
    if ($matchedEvent.Id -ne $lastEventId) {
        # Send the HTTP request
        SendHttpPostRequest -url "https://api.smtp2go.com/v3/email/send/" -body $jsonTemplate
        # Store the ID of the event in the file
        $matchedEvent.Id | Out-File -FilePath $lastEventFile
        # to-do: Restart the service. Will need the service name as a parameter..
    }
}

Stop-Transcript