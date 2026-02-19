<#
.SYNOPSIS
    Genera le disponibilità Outlook con temi di colore selezionabili da config.json.
#>
Param(
    [ValidateSet("Tabella", "Testo")]
    [String]$Formato = "Tabella"
)

# ---- 1. CONFIGURAZIONE ----
$configPath = Join-Path $PSScriptRoot "config.json"
if (-not (Test-Path $configPath)) {
    Write-Error "Errore: File config.json non trovato!"
    return
}
$config = Get-Content $configPath -Encoding UTF8 | ConvertFrom-Json
$txt = $config.Localization
$culture = [System.Globalization.CultureInfo]::GetCultureInfo($txt.Culture)

# Selezione del Tema
$themeName = $config.Preferences.SelectedTheme
$theme = $config.ColorThemes.$themeName
if (-not $theme) { $theme = $config.ColorThemes.Grigio } # Fallback se il nome non esiste

# ---- 2. OUTLOOK ----
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
} catch {
    Write-Error "Impossibile aprire Outlook."
    return
}
$today = Get-Date

# ---- 3. FUNZIONI (Logica Originale) ----
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
    $lStart = [datetime]::new($day.Year, $day.Month, $day.Day, $config.LunchBreak.Start, 0, 0)
    $lEnd   = [datetime]::new($day.Year, $day.Month, $day.Day, $config.LunchBreak.End, 0, 0)
    $result = @()
    foreach ($slot in $slots) {
        if ($slot.End -le $lStart -or $slot.Start -ge $lEnd) { $result += $slot }
        elseif ($slot.Start -lt $lStart -and $slot.End -gt $lStart -and $slot.End -le $lEnd) { $result += [PSCustomObject]@{ Start=$slot.Start; End=$lStart } }
        elseif ($slot.Start -ge $lStart -and $slot.Start -lt $lEnd -and $slot.End -gt $lEnd) { $result += [PSCustomObject]@{ Start=$lEnd; End=$slot.End } }
        elseif ($slot.Start -lt $lStart -and $slot.End -gt $lEnd) {
            $result += [PSCustomObject]@{ Start=$slot.Start; End=$lStart }
            $result += [PSCustomObject]@{ Start=$lEnd; End=$slot.End }
        }
    }
    return $result | Where-Object { $_.End -gt $_.Start }
}

# ---- 4. ELABORAZIONE ----
$daysProcessed = @()
$workdaysAdded = 0
$i = 0
while ($workdaysAdded -lt ($config.Preferences.DaysForward + 1)) {
    $day = $today.Date.AddDays($i); $i++
    if ($day.DayOfWeek -in @("Saturday","Sunday")) { continue }
    $workdaysAdded++
    $dayStart = $day.AddHours($config.WorkingHours.Start)
    $dayEnd = $day.AddHours($config.WorkingHours.End)
    if ($day.Date -eq $today.Date -and $today -gt $dayStart) { $dayStart = Round-Up5($today) }
    $apptsAll = @()
    foreach ($store in $namespace.Stores) {
        if ($config.CalendarsToRead -contains $store.DisplayName) {
            try {
                $folder = $store.GetDefaultFolder(9)
                foreach ($appt in (Get-Appointments $folder $dayStart $dayEnd)) {
                    if ($appt.BusyStatus -ne 0) {
                        $s = if ($appt.Start -lt $dayStart) { $dayStart } else { $appt.Start }
                        $e = if ($appt.End -gt $dayEnd) { $dayEnd } else { $appt.End }
                        $apptsAll += [PSCustomObject]@{ Start = Round-Down5 $s; End = Round-Up5 $e }
                    }
                }
            } catch {}
        }
    }
    $free = if ($apptsAll.Count -eq 0) { @([PSCustomObject]@{ Start=$dayStart; End=$dayEnd }) }
            else {
                $m = Merge-Appointments $apptsAll; $sols = @(); $curr = $dayStart
                foreach ($a in $m) {
                    if ($a.Start -gt $curr) { $sols += [PSCustomObject]@{ Start=$curr; End=$a.Start } }
                    if ($a.End -gt $curr) { $curr = $a.End }
                }
                if ($curr -lt $dayEnd) { $sols += [PSCustomObject]@{ Start=$curr; End=$dayEnd } }
                $sols
            }
    $valid = Exclude-Lunch $free $day | Where-Object { ($_.End - $_.Start).TotalMinutes -ge $config.Preferences.MinSlotDurationMinutes }
    if ($valid) { $daysProcessed += [PSCustomObject]@{ Day = $day; Slots = $valid } }
}

# ---- 5. OUTPUT (DINAMICO) ----
$htmlHeader = "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
$mailBody = ""

if ($Formato -eq "Tabella") {
    $rows = ""
    $count = 0
    foreach ($d in $daysProcessed) {
        $dayStrRaw = $d.Day.ToString("dddd dd/MM", $culture)
        $dayStr = $dayStrRaw.Substring(0,1).ToUpper() + $dayStrRaw.Substring(1)
        
        $badges = foreach ($s in $d.Slots) {
            $tS = $s.Start.ToString("HH:mm")
            $range = if ($s.End -ge $d.Day.AddHours($config.WorkingHours.End)) { "$tS $($txt.EndOfDaySuffix)" } else { "$tS - $($s.End.ToString('HH:mm'))" }
            "<span style='display:inline-block;background:$($theme.BadgeBg);color:$($theme.BadgeText);border:1px solid $($theme.BadgeBorder);padding:3px 10px;border-radius:6px;margin:2px;font-size:12px;font-weight:600;'>$range</span>"
        }
        
        $bg = if ($count++ % 2 -eq 0) { "#ffffff" } else { $theme.RowAlt }
        $rows += "<tr style='background:$bg;'><td style='padding:8px 15px;font-weight:bold;border-bottom:1px solid #e2e8f0;color:#374151;'>${dayStr}</td><td style='padding:8px 15px;border-bottom:1px solid #e2e8f0;'>$($badges -join ' ')</td></tr>"
    }

    $mailBody = @"
    $htmlHeader
    <div style="font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color:#111827; max-width:650px;">
        <p>$($txt.Greeting)</p>
        <p>$($txt.IntroText)</p>
        <table style="border-collapse:collapse; border:1px solid #e2e8f0; border-radius:8px; overflow:hidden; min-width:400px;">
            <thead>
                <tr style="background:$($theme.HeaderBg); color:$($theme.HeaderText);">
                    <th style="padding:10px 15px; text-align:left; font-size:11px; text-transform:uppercase; letter-spacing:0.05em;">$($txt.TableHeaderDay)</th>
                    <th style="padding:10px 15px; text-align:left; font-size:11px; text-transform:uppercase; letter-spacing:0.05em;">$($txt.TableHeaderAvailability)</th>
                </tr>
            </thead>
            <tbody>
                $rows
            </tbody>
        </table>
        <p>$($txt.Closing)</p>
    </div>
"@
} else {
    $lines = @("$htmlHeader<div style='font-family:Segoe UI, sans-serif;'><div>$($txt.Greeting)</div><br><div>$($txt.IntroTextList)</div>")
    foreach ($d in $daysProcessed) {
        $dayStrRaw = $d.Day.ToString("dddd dd/MM", $culture)
        $dayStr = $dayStrRaw.Substring(0,1).ToUpper() + $dayStrRaw.Substring(1)
        $parts = foreach ($s in $d.Slots) {
            $tS = $s.Start.ToString("HH.mm")
            if ($s.End -ge $d.Day.AddHours($config.WorkingHours.End)) { "$($txt.FromTimeText)$tS$($txt.InPoiText)" }
            else { "$($txt.FromTimeText)$tS$($txt.ToTimeText)$($s.End.ToString('HH.mm'))" }
        }
        $lines += "<div style='margin-bottom:5px;'><b>${dayStr}:</b> $($parts -join $txt.OrText)</div>"
    }
    $lines += "<br><div>$($txt.Closing)</div></div>"
    $mailBody = $lines -join ""
}

# ---- 6. OUTPUT ----
$mail = $outlook.CreateItem(0)
$mail.Subject = $txt.MailSubject
$mail.HTMLBody = $mailBody
$mail.Display()
