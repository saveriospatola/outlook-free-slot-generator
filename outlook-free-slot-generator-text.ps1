<#
.SYNOPSIS
    Genera un elenco testuale delle disponibilità Outlook basato su file di configurazione JSON.
    
.DESCRIPTION
    Lo script legge uno o più calendari, esclude i weekend e la pausa pranzo, 
    e crea una bozza di email in Outlook con le fasce orarie libere.
#>

# ---- CARICAMENTO CONFIGURAZIONE ----
$configPath = Join-Path $PSScriptRoot "config.json"
if (-not (Test-Path $configPath)) {
    Write-Error "Errore: File config.json non trovato in $PSScriptRoot"
    return
}

# Utilizziamo UTF8 per leggere il JSON ed evitare errori sui caratteri speciali
$config = Get-Content $configPath -Encoding UTF8 | ConvertFrom-Json
$txt = $config.Localization
$culture = [System.Globalization.CultureInfo]::GetCultureInfo($txt.Culture)

# ---- INIZIALIZZAZIONE OUTLOOK ----
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
} catch {
    Write-Error "Impossibile aprire Outlook. Assicurati che sia installato."
    return
}

$today = Get-Date

# ---- FUNZIONI DI SUPPORTO ----
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
        elseif ($slot.Start -lt $lStart -and $slot.End -gt $lStart -and $slot.End -le $lEnd) {
            $result += [PSCustomObject]@{ Start=$slot.Start; End=$lStart }
        }
        elseif ($slot.Start -ge $lStart -and $slot.Start -lt $lEnd -and $slot.End -gt $lEnd) {
            $result += [PSCustomObject]@{ Start=$lEnd; End=$slot.End }
        }
        elseif ($slot.Start -lt $lStart -and $slot.End -gt $lEnd) {
            $result += [PSCustomObject]@{ Start=$slot.Start; End=$lStart }
            $result += [PSCustomObject]@{ Start=$lEnd; End=$slot.End }
        }
    }
    return $result | Where-Object { $_.End -gt $_.Start }
}

# ---- COSTRUZIONE CORPO EMAIL ----
$bodyLines = @()
$bodyLines += "<div>$($txt.Greeting)</div>"
$bodyLines += "<div><br></div>"
$bodyLines += "<div>$($txt.IntroTextList)</div>"

$workdaysAdded = 0
$i = 0

while ($workdaysAdded -lt ($config.Preferences.DaysForward + 1)) {
    $day = $today.Date.AddDays($i)
    $i++

    if ($day.DayOfWeek -in @("Saturday","Sunday")) { continue }
    $workdaysAdded++

    $dayStart = $day.AddHours($config.WorkingHours.Start)
    $dayEnd = $day.AddHours($config.WorkingHours.End)
    if ($day.Date -eq $today.Date -and $today -gt $dayStart) { $dayStart = Round-Up5($today) }

    $appointmentsAll = @()
    foreach ($store in $namespace.Stores) {
        if ($config.CalendarsToRead -contains $store.DisplayName) {
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

    $freeSlots = @()
    if ($appointmentsAll.Count -eq 0) { $freeSlots = @([PSCustomObject]@{ Start=$dayStart; End=$dayEnd }) }
    else {
        $merged = Merge-Appointments $appointmentsAll
        $current = $dayStart
        foreach ($appt in $merged) {
            if ($appt.Start -gt $current) { $freeSlots += [PSCustomObject]@{ Start=$current; End=$appt.Start } }
            if ($appt.End -gt $current) { $current = $appt.End }
        }
        if ($current -lt $dayEnd) { $freeSlots += [PSCustomObject]@{ Start=$current; End=$dayEnd } }
    }

    $freeSlots = Exclude-Lunch $freeSlots $day

    $parts = @()
    foreach ($slot in $freeSlots) {
        if (($slot.End - $slot.Start).TotalMinutes -ge $config.Preferences.MinSlotDurationMinutes) {
            $sTime = $slot.Start.ToString("HH.mm")
            $eTime = $slot.End.ToString("HH.mm")
            
            if ($slot.End -ge $dayEnd) { 
                $parts += "$($txt.FromTimeText)$sTime$($txt.InPoiText)" 
            } else { 
                $parts += "$($txt.FromTimeText)$sTime$($txt.ToTimeText)$eTime" 
            }
        }
    }

    if ($parts.Count -gt 0) {
        $dayStrRaw = $day.ToString("dddd dd/MM", $culture)
        $dayStr = $dayStrRaw.Substring(0,1).ToUpper() + $dayStrRaw.Substring(1)
        $bodyLines += "<div><b>${dayStr}:</b> " + ($parts -join $txt.OrText) + "</div>"
    }
}

$bodyLines += "<div><br></div>"
$bodyLines += "<div>$($txt.Closing)</div>"

# Generazione finale della mail
$mail = $outlook.CreateItem(0)
$mail.Subject = $txt.MailSubject
$mail.HTMLBody = ($bodyLines -join "")
$mail.Display()