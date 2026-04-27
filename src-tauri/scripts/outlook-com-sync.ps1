param(
    [Parameter(Mandatory = $true)]
    [string]$WindowStartUtc,

    [Parameter(Mandatory = $true)]
    [string]$WindowEndUtc,

    [switch]$IncludePrivateDetails
)

$ErrorActionPreference = "Stop"
$WarningPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

function Convert-InputUtcToLocal([string]$Value) {
    return ([datetimeoffset]::Parse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)).UtcDateTime.ToLocalTime()
}

function Format-OutlookRestrictDate([datetime]$Value) {
    return $Value.ToString("g", [Globalization.CultureInfo]::CurrentCulture)
}

function Get-SafeProperty($Object, [string]$Name) {
    if ($null -eq $Object) { return $null }
    try {
        return $Object.$Name
    } catch {
        return $null
    }
}

function Convert-ToIsoUtc($Value, [bool]$AlreadyUtc) {
    if ($null -eq $Value) { return $null }
    $date = [datetime]$Value
    if ($AlreadyUtc) {
        $date = [datetime]::SpecifyKind($date, [DateTimeKind]::Utc)
    } elseif ($date.Kind -eq [DateTimeKind]::Unspecified) {
        $date = [datetime]::SpecifyKind($date, [DateTimeKind]::Local)
    }
    return $date.ToUniversalTime().ToString("o", [Globalization.CultureInfo]::InvariantCulture)
}

function Convert-Sensitivity($Value) {
    if ($null -eq $Value) { return $null }
    switch ([int]$Value) {
        0 { return "normal" }
        1 { return "personal" }
        2 { return "private" }
        3 { return "confidential" }
        default { return [string]$Value }
    }
}

function Split-Categories($Value) {
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return @() }
    return ([string]$Value).Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

function Release-ComObject($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

function Join-UnicodeText([string[]]$HexCodePoints) {
    return -join ($HexCodePoints | ForEach-Object { [char][Convert]::ToInt32($_, 16) })
}

$UntitledEventText = Join-UnicodeText @("65E0", "6807", "9898", "65E5", "7A0B")
$PrivateEventText = Join-UnicodeText @("79C1", "5BC6", "65E5", "7A0B")
$ClassicOutlookPrefixText = Join-UnicodeText @("672C", "673A", "7ECF", "5178")

$windowStartLocal = Convert-InputUtcToLocal $WindowStartUtc
$windowEndLocal = Convert-InputUtcToLocal $WindowEndUtc
$outlook = $null
$session = $null
$calendar = $null
$items = $null
$restricted = $null

try {
    $outlook = New-Object -ComObject Outlook.Application
    $session = $outlook.GetNamespace("MAPI")
    $calendar = $session.GetDefaultFolder(9)
    $items = $calendar.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true

    $startText = Format-OutlookRestrictDate $windowStartLocal
    $endText = Format-OutlookRestrictDate $windowEndLocal
    $filter = "[Start] < '$endText' AND [End] > '$startText'"
    $restricted = $items.Restrict($filter)

    $profileName = [string](Get-SafeProperty $session "CurrentProfileName")
    $currentUser = Get-SafeProperty $session "CurrentUser"
    $currentUserName = [string](Get-SafeProperty $currentUser "Name")
    $currentUserAddress = [string](Get-SafeProperty $currentUser "Address")
    $currentUserAddressEntry = Get-SafeProperty $currentUser "AddressEntry"
    $exchangeUser = $null
    if ($null -ne $currentUserAddressEntry) {
        try {
            $exchangeUser = $currentUserAddressEntry.GetExchangeUser()
        } catch {
            $exchangeUser = $null
        }
    }
    $primarySmtpAddress = [string](Get-SafeProperty $exchangeUser "PrimarySmtpAddress")
    if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
        $currentUserAddress = $primarySmtpAddress
    }
    $storeId = [string](Get-SafeProperty $calendar "StoreID")
    $calendarName = [string](Get-SafeProperty $calendar "Name")

    if ([string]::IsNullOrWhiteSpace($currentUserAddress)) {
        $currentUserAddress = if ([string]::IsNullOrWhiteSpace($profileName)) { "classic-outlook-local" } else { $profileName }
    }
    if ([string]::IsNullOrWhiteSpace($currentUserName)) {
        $currentUserName = $currentUserAddress
    }
    if ([string]::IsNullOrWhiteSpace($calendarName)) {
        $calendarName = "Calendar"
    }

    $providerUserIdSeed = if ([string]::IsNullOrWhiteSpace($storeId)) { "$profileName|$currentUserAddress" } else { $storeId }
    $providerUserId = "classic-outlook:$providerUserIdSeed"
    $providerCalendarId = if ([string]::IsNullOrWhiteSpace($storeId)) { "default-calendar" } else { "default-calendar:$storeId" }
    $events = New-Object System.Collections.Generic.List[object]

    foreach ($item in $restricted) {
        $messageClass = [string](Get-SafeProperty $item "MessageClass")
        if ($messageClass -notlike "IPM.Appointment*") { continue }

        $startUtc = Convert-ToIsoUtc (Get-SafeProperty $item "StartUTC") $true
        if ($null -eq $startUtc) {
            $startUtc = Convert-ToIsoUtc (Get-SafeProperty $item "Start") $false
        }
        $endUtc = Convert-ToIsoUtc (Get-SafeProperty $item "EndUTC") $true
        if ($null -eq $endUtc) {
            $endUtc = Convert-ToIsoUtc (Get-SafeProperty $item "End") $false
        }
        if ($null -eq $startUtc -or $null -eq $endUtc) { continue }

        $sensitivity = Convert-Sensitivity (Get-SafeProperty $item "Sensitivity")
        $isPrivate = $sensitivity -eq "private"
        $subject = [string](Get-SafeProperty $item "Subject")
        if ([string]::IsNullOrWhiteSpace($subject)) {
            $subject = $UntitledEventText
        }
        if ($isPrivate -and -not $IncludePrivateDetails) {
            $subject = $PrivateEventText
        }

        $globalAppointmentId = [string](Get-SafeProperty $item "GlobalAppointmentID")
        $entryId = [string](Get-SafeProperty $item "EntryID")
        $isRecurring = [bool](Get-SafeProperty $item "IsRecurring")
        $recurrenceState = Get-SafeProperty $item "RecurrenceState"
        if (-not [string]::IsNullOrWhiteSpace($globalAppointmentId)) {
            $eventKey = "global:$globalAppointmentId"
        } elseif (-not [string]::IsNullOrWhiteSpace($entryId)) {
            $eventKey = "entry:$entryId"
        } else {
            $eventKey = "fallback:$subject"
        }
        if ($isRecurring -or ($null -ne $recurrenceState -and [int]$recurrenceState -ne 0)) {
            $eventKey = "$eventKey|start:$startUtc"
        }

        $hidePrivateDetails = $isPrivate -and -not $IncludePrivateDetails
        $body = $null
        if (-not $hidePrivateDetails) {
            $body = [string](Get-SafeProperty $item "Body")
            if ([string]::IsNullOrWhiteSpace($body)) { $body = $null }
        }

        $attendees = $null
        $organizer = $null
        $location = $null
        $onlineMeetingUrl = $null
        $categories = @()
        if (-not $hidePrivateDetails) {
            $attendees = [ordered]@{
                required = [string](Get-SafeProperty $item "RequiredAttendees")
                optional = [string](Get-SafeProperty $item "OptionalAttendees")
                resources = [string](Get-SafeProperty $item "Resources")
            }
            $organizer = [ordered]@{
                name = [string](Get-SafeProperty $item "Organizer")
            }
            $location = [string](Get-SafeProperty $item "Location")
            $onlineMeetingUrl = [string](Get-SafeProperty $item "OnlineMeetingUrl")
            $categories = @(Split-Categories (Get-SafeProperty $item "Categories"))
        }
        if ([string]::IsNullOrWhiteSpace($location)) { $location = $null }
        if ([string]::IsNullOrWhiteSpace($onlineMeetingUrl)) { $onlineMeetingUrl = $null }

        $events.Add([pscustomobject]@{
            providerEventId = $eventKey
            title = $subject
            bodyContentType = if ($null -eq $body) { $null } else { "text" }
            bodyContent = $body
            startUtc = $startUtc
            endUtc = $endUtc
            startTimezone = [System.TimeZoneInfo]::Local.Id
            endTimezone = [System.TimeZoneInfo]::Local.Id
            isAllDay = [bool](Get-SafeProperty $item "AllDayEvent")
            location = $location
            attendees = $attendees
            organizer = $organizer
            webLink = $null
            onlineMeetingUrl = $onlineMeetingUrl
            categories = $categories
            reminderMinutesBeforeStart = Get-SafeProperty $item "ReminderMinutesBeforeStart"
            isReminderOn = [bool](Get-SafeProperty $item "ReminderSet")
            sensitivity = $sensitivity
            lastModifiedUtc = Convert-ToIsoUtc (Get-SafeProperty $item "LastModificationTime") $false
        }) | Out-Null
    }

    [pscustomobject]@{
        account = [pscustomobject]@{
            providerUserId = $providerUserId
            email = $currentUserAddress
            displayName = "$ClassicOutlookPrefixText Outlook - $currentUserName"
            calendarId = $providerCalendarId
            calendarName = $calendarName
        }
        windowStartUtc = $WindowStartUtc
        windowEndUtc = $WindowEndUtc
        events = $events
    } | ConvertTo-Json -Depth 8 -Compress
} finally {
    Release-ComObject $restricted
    Release-ComObject $items
    Release-ComObject $calendar
    Release-ComObject $session
    Release-ComObject $outlook
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
