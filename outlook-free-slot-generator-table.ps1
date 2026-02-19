# Carica la configurazione dal file JSON
$configPath = Join-Path $PSScriptRoot "config.json"
if (-not (Test-Path $configPath)) {
    Write-Error "File di configurazione config.json non trovato!"
    return
}
# Forza la lettura in UTF8 per gestire correttamente gli accenti nel JSON
$config = Get-Content $configPath -Encoding UTF8 | ConvertFrom-Json

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$today = Get-Date

# ---- CONFIGURAZIONE CARICATA DA JSON ----
$calendarsToRead = $config.CalendarsToRead
$workingDayStart = $config.WorkingHours.Start
$workingDayEnd   = $config.WorkingHours.End
$lunchStart      = $config.LunchBreak.Start
$lunchEnd        = $config.LunchBreak.End
$workingDaysForward      = $config.Preferences.DaysForward
$minSlotDurationMinutes  = $config.Preferences.MinSlotDurationMinutes
$tableHeaderColor        = $config.Preferences.TableHeaderColor
$rowAlternateColor       = $config.Preferences.RowAlternateColor

# ---- LOCALIZZAZIONE ----
$txt = $config.Localization
$culture = [System.Globalization.CultureInfo]::GetCultureInfo($txt.Culture)
# -----------------------------------------

# ---- FUNZIONI ----
function Round-Up5($date) {
    $minutes = [math]::Ceiling($date.Minute / 5) * 5
    if ($minutes -eq 60) { return [datetime]::new($date.Year, $date.Month, $date.Day, $date.Hour, 0, 0).AddHours(1) }
    return [datetime]::new($date.Year, $date.Month, $date.Day, $date.Hour, $minutes, 0)
}

function Round-Down5($date) {
    $minutes = [math]::Floor($date.Minute / 5) * 5
    return [datetime]::new($date.Year, $date.Month, $date.Day, $date.Hour, $minutes, 0)
}

function Get-Appointments($folder, $from, $to) {
    $items = $folder.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    $filter = "[Start] < '" + $to.ToString("g") + "' AND [End] > '" + $from.ToString("g") + "'"
    return $items.Restrict($filter)
}

function Merge-Appointments($appointments) {
    $sorted = $appointments | Sort-Object Start
    $merged = @()
    foreach ($appt in $sorted) {
        if ($merged.Count -eq 0) { $merged += $appt }
        else {
            $last = $merged[-1]
            if ($appt.Start -le $last.End) { if ($appt.End -gt $last.End) { $last.End = $appt.End } }
            else { $merged += $appt }
        }
    }
    return $merged
}

function Exclude-Lunch($slots, $day) {
    $lunchStartTime = [datetime]::new($day.Year, $day.Month, $day.Day, $lunchStart, 0, 0)
    $lunchEndTime   = [datetime]::new($day.Year, $day.Month, $day.Day, $lunchEnd, 0, 0)
    $result = @()
    foreach ($slot in $slots) {
        if ($slot.End -le $lunchStartTime -or $slot.Start -ge $lunchEndTime) { $result += $slot }
        elseif ($slot.Start -lt $lunchStartTime -and $slot.End -gt $lunchStartTime -and $slot.End -le $lunchEndTime) {
            $result += [PSCustomObject]@{ Start=$slot.Start; End=$lunchStartTime }
        }
        elseif ($slot.Start -ge $lunchStartTime -and $slot.Start -lt $lunchEndTime -and $slot.End -gt $lunchEndTime) {
            $result += [PSCustomObject]@{ Start=$lunchEndTime; End=$slot.End }
        }
        elseif ($slot.Start -lt $lunchStartTime -and $slot.End -gt $lunchEndTime) {
            $result += [PSCustomObject]@{ Start=$slot.Start; End=$lunchStartTime }
            $result += [PSCustomObject]@{ Start=$lunchEndTime; End=$slot.End }
        }
    }
    return $result | Where-Object { $_.End -gt $_.Start }
}

# ---- GENERAZIONE HTML ----
$htmlHeader = @"
<div style="font-family: 'Segoe UI', Helvetica, Arial, sans-serif; color: #1e293b; line-height: 1.4;">
    <p>$($txt.Greeting)</p>
    <p>$($txt.IntroText)</p>
    
    <table style="display: inline-table; border-collapse: separate; border-spacing: 0; margin-top: 5px; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
        <thead>
            <tr style="background-color: $tableHeaderColor; color: white;">
                <th style="padding: 6px 15px; text-align: center; vertical-align: middle; font-weight: 600; font-size: 11px; white-space: nowrap; text-transform: uppercase; letter-spacing: 0.05em;">$($txt.TableHeaderDay)</th>
                <th style="padding: 6px 15px; text-align: left; vertical-align: middle; font-weight: 600; font-size: 11px; white-space: nowrap; text-transform: uppercase; letter-spacing: 0.05em;">$($txt.TableHeaderAvailability)</th>
            </tr>
        </thead>
        <tbody>
"@

$tableRows = ""
$workdaysAdded = 0
$i = 0

while ($workdaysAdded -lt ($workingDaysForward + 1)) {
    $day = $today.Date.AddDays($i)
    $i++
    if ($day.DayOfWeek -in @("Saturday","Sunday")) { continue }
    $workdaysAdded++

    $dayStart = $day.AddHours($workingDayStart)
    $dayEnd = $day.AddHours($workingDayEnd)
    if ($day.Date -eq $today.Date -and $today -gt $dayStart) { $dayStart = Round-Up5($today) }

    $appointmentsAll = @()
    foreach ($store in $namespace.Stores) {
        if ($calendarsToRead -contains $store.DisplayName) {
            try {
                $folder = $store.GetDefaultFolder(9)
                $appointments = Get-Appointments $folder $dayStart $dayEnd
                foreach ($appt in $appointments) {
                    if ($appt.BusyStatus -ne 0) {
                        $start = if ($appt.Start -lt $dayStart) { $dayStart } else { $appt.Start }
                        $end   = if ($appt.End -gt $dayEnd) { $dayEnd } else { $appt.End }
                        $start = Round-Down5 $start
                        $end   = Round-Up5 $end
                        if ($end -gt $start) { $appointmentsAll += [PSCustomObject]@{ Start = $start; End = $end } }
                    }
                }
            } catch {}
        }
    }

    if ($appointmentsAll.Count -eq 0) { $freeSlots = @([PSCustomObject]@{ Start=$dayStart; End=$dayEnd }) }
    else {
        $merged = Merge-Appointments $appointmentsAll
        $freeSlots = @()
        $current = $dayStart
        foreach ($appt in $merged) {
            if ($appt.Start -gt $current) { $freeSlots += [PSCustomObject]@{ Start=$current; End=$appt.Start } }
            if ($appt.End -gt $current) { $current = $appt.End }
        }
        if ($current -lt $dayEnd) { $freeSlots += [PSCustomObject]@{ Start=$current; End=$dayEnd } }
    }

    $freeSlots = Exclude-Lunch $freeSlots $day
    
    $badges = @()
    foreach ($slot in $freeSlots) {
        $duration = ($slot.End - $slot.Start).TotalMinutes
        if ($duration -ge $minSlotDurationMinutes) {
            $timeRange = $slot.Start.ToString("HH:mm") + " - " + $slot.End.ToString("HH:mm")
            if ($slot.End -ge $dayEnd) { $timeRange = $slot.Start.ToString("HH:mm") + " " + $txt.EndOfDaySuffix }
            $badges += "<span style='display: inline-block; background-color: #eff6ff; color: #1d4ed8; border: 1px solid #dbeafe; padding: 3px 10px; border-radius: 6px; margin: 2px; font-size: 12px; font-weight: 600; white-space: nowrap;'>$timeRange</span>"
        }
    }

    if ($badges.Count -gt 0) {
        $dayStringRaw = $day.ToString("dddd dd/MM", $culture)
        $dayString = $dayStringRaw.Substring(0,1).ToUpper() + $dayStringRaw.Substring(1)
        $rowBg = if ($workdaysAdded % 2 -eq 0) { $rowAlternateColor } else { "#ffffff" }
        $tableRows += "<tr style='background-color: $rowBg;'><td style='padding: 6px 15px; border-bottom: 1px solid #e2e8f0; font-weight: bold; font-size: 13px; text-align: center; vertical-align: middle; white-space: nowrap;'>$dayString</td><td style='padding: 6px 15px; border-bottom: 1px solid #e2e8f0; text-align: left; vertical-align: middle;'>$($badges -join ' ')</td></tr>"
    }
}

$htmlFooter = "</tbody></table><p style='margin-top: 15px;'>$($txt.Closing)</p></div>"
$mailBody = $htmlHeader + $tableRows + $htmlFooter

$mail = $outlook.CreateItem(0)
$mail.Subject = $txt.MailSubject
$mail.HTMLBody = $mailBody
$mail.Display()