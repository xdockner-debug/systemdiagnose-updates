#Requires -Version 5.1
<#
.SYNOPSIS
    Bereinigungsteil der Systemdiagnose - fuehrt NUR vordefinierte, gepruefte Reparaturen aus.

.DESCRIPTION
    Wird von der App mit Adminrechten gestartet und bekommt einen Plan (JSON) aus dem Ordner
    "Bereinigung". Sicherheitsmechanismen:
      1. Fester Katalog: Ausgefuehrt werden nur die unten fest programmierten Aktionen. Der Plan
         liefert ausschliesslich eine Aktions-Kennung und Parameter - niemals Befehle.
      2. Strenge Parameterpruefung (Format) UND erneute Pruefung am laufenden System direkt vor
         jeder Aktion: Besteht das Problem wirklich noch? Sonst wird abgelehnt bzw. uebersprungen.
      3. Harte Grenzen: Microsoft-Dienste werden nie abgeschaltet; Treiberpakete werden nur
         entfernt, wenn eine aktive Ersatzversion gebunden ist und ein hoechstens 24 h alter
         Wiederherstellungspunkt existiert.
      4. Umkehrbare Aktionen speichern den alten Zustand -> "Rueckgaengig".
      5. Einstellungen des Benutzers (HKCU) werden nur geaendert, wenn der Plan vom selben
         Benutzerkonto stammt, unter dem die Bereinigung laeuft.
      6. Plan-, Status- und Ergebnisdateien nur im eigenen Ordner; alles wird protokolliert.
      7. Reinigung (Caches, Temp, Protokolle): Der Plan nennt nur eine Kategorie aus Reinigung.psd1,
         niemals einen Pfad. Die Ordner werden hier selbst aufgeloest und hart geprueft (nie Laufwerks-
         wurzel, Windows-/Programmordner oder Benutzerordner). Geloescht wird Datei fuer Datei ueber ein
         Handle, dessen tatsaechlicher Pfad vorher geprueft wird - Verknuepfungen (Junctions, Symlinks)
         koennen so nicht nach aussen umleiten. Laufende Programme der Kategorie => uebersprungen.

.PARAMETER Plan
    Plan-Datei (Plan-JJJJMMTT_HHMMSS.json im Ordner "Bereinigung").
.PARAMETER Rueckgaengig
    Ergebnis-Datei (Ergebnis-JJJJMMTT_HHMMSS.json), deren umkehrbare Aktionen zurueckgenommen werden.
.PARAMETER Simulation
    Prueft alles wie im Ernstfall, aendert aber nichts.
#>
param(
    [string] $Plan,
    [string] $Rueckgaengig,
    [switch] $Simulation,
    [ValidatePattern('^\d{8}_\d{6}$')]
    [string] $Stempel = (Get-Date).ToString('yyyyMMdd_HHmmss')
)

$ErrorActionPreference = 'Stop'
# Wie in der Engine: externe Programme (powercfg) in UTF-8 lesen - sonst wird aus „Ryzen™“ je nach Startart „RyzenT“
# und der exakte Vergleich mit dem eingelernten Energieplan schlaegt fehl
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
$Script:ToolOrdner    = Split-Path -Parent $PSCommandPath
# Datenordner wie in Systemdiagnose.ps1 (Store-Fassung: LocalState des Pakets statt schreibgeschuetztem Programmordner)
$Script:DatenOrdner   = $Script:ToolOrdner
if ($Script:ToolOrdner -match '\\WindowsApps\\([^\\_]+)_[^\\_]*_[^\\_]*_[^\\_]*_([^\\_]+)$') { $Script:DatenOrdner = Join-Path $env:LOCALAPPDATA ('Packages\' + $Matches[1] + '_' + $Matches[2] + '\LocalState') }
$Script:Arbeitsordner = Join-Path $Script:DatenOrdner 'Bereinigung'
$Script:IstSimulation = [bool]$Simulation
$Script:StatusDatei   = $null
$Script:GleicherBenutzer = $false
$Script:WiederherstellungOk = $false
$Script:Cache = @{}

# Dienste, die abgeschaltet bzw. gestartet werden duerfen, OBWOHL ihr Programm existiert (PC-Profil).
# Alle anderen Dienste werden nur abgeschaltet, wenn ihr Programm nachweislich fehlt.
$Script:DeaktivierbareDienste = @('NahimicService')
$Script:StartbareDienste      = @('logi_lamparray_service')

# Fester Katalog: Umkehrbar = alter Zustand wird gespeichert; Sicherung = braucht Wiederherstellungspunkt
$Script:Katalog = @{
    'TreiberpaketEntfernen'          = @{ Umkehrbar = $false; Sicherung = $true }
    'TreiberdienstDeaktivieren'      = @{ Umkehrbar = $true;  Sicherung = $false }
    'DienstDeaktivieren'             = @{ Umkehrbar = $true;  Sicherung = $false }
    'DienstStarten'                  = @{ Umkehrbar = $true;  Sicherung = $false }
    'AufgabeDeaktivieren'            = @{ Umkehrbar = $true;  Sicherung = $false }
    'AutostartDeaktivieren'          = @{ Umkehrbar = $true;  Sicherung = $false }
    'Reinigen'                       = @{ Umkehrbar = $false; Sicherung = $false }
    'WiederherstellungspunktAnlegen' = @{ Umkehrbar = $false; Sicherung = $false }
    'DismReparieren'                 = @{ Umkehrbar = $false; Sicherung = $false }
    'SfcReparieren'                  = @{ Umkehrbar = $false; Sicherung = $false }
    'EnergieplanSetzen'              = @{ Umkehrbar = $true;  Sicherung = $false }
    'SchnellstartAus'                = @{ Umkehrbar = $true;  Sicherung = $false }
    'GameDvrAus'                     = @{ Umkehrbar = $true;  Sicherung = $false }
    'DynamicLightingAus'             = @{ Umkehrbar = $true;  Sicherung = $false }
    'AbbildTypKlein'                 = @{ Umkehrbar = $true;  Sicherung = $false }
    'UacEinschalten'                 = @{ Umkehrbar = $true;  Sicherung = $false }
    'UacNachfrageAn'                 = @{ Umkehrbar = $true;  Sicherung = $false }
    'DefenderSignaturen'             = @{ Umkehrbar = $false; Sicherung = $false }
    'FirewallEinschalten'            = @{ Umkehrbar = $true;  Sicherung = $false }
    'ZeitSynchronisieren'            = @{ Umkehrbar = $false; Sicherung = $false }
    'DnsCacheLeeren'                 = @{ Umkehrbar = $false; Sicherung = $false }
    'DoppeltesSkriptBeenden'         = @{ Umkehrbar = $false; Sicherung = $false }
}

$Script:AutostartOrte = @{
    'Run (Benutzer)'         = @{ Hive = 'HKCU'; Run = 'Software\Microsoft\Windows\CurrentVersion\Run';             Freigabe = 'Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' }
    'Run (Computer)'         = @{ Hive = 'HKLM'; Run = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Run';             Freigabe = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' }
    'Run (Computer, 32 Bit)' = @{ Hive = 'HKLM'; Run = 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; Freigabe = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32' }
}

# ============================================================================
#  Hilfsfunktionen
# ============================================================================

function Test-IstAdmin {
    $pr = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-SichereDatei {
    param([string] $Pfad, [string] $Muster)
    $voll = [IO.Path]::GetFullPath($Pfad)
    $ordner = [IO.Path]::GetFullPath($Script:Arbeitsordner)
    if ((Split-Path -Parent $voll) -ne $ordner) { throw 'Die Datei liegt nicht im Bereinigungsordner.' }
    if ((Split-Path -Leaf $voll) -notmatch $Muster) { throw 'Unzulässiger Dateiname.' }
    if (-not (Test-Path -LiteralPath $voll -PathType Leaf)) { throw 'Datei nicht gefunden.' }
    return $voll
}

function Write-Status {
    param([string] $Text)
    if ($Script:StatusDatei) { try { [IO.File]::WriteAllText($Script:StatusDatei, $Text, (New-Object Text.UTF8Encoding($false))) } catch { } }
}

function Write-Protokoll {
    param([string] $Zeile)
    try { [IO.File]::AppendAllText((Join-Path $Script:Arbeitsordner 'Protokoll.txt'), ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '  ' + $Zeile + "`r`n"), (New-Object Text.UTF8Encoding($true))) } catch { }
}

function New-Ergebnis {
    param([string] $Status, [string] $Meldung, [object[]] $Rueckgaengig = @(), [int64] $Bytes = -1)
    return [pscustomobject]@{ Status = $Status; Meldung = $Meldung; Rueckgaengig = @($Rueckgaengig | Where-Object { $_ }); Bytes = $Bytes }
}

# Fuer Fehlerberichte von anderen PCs: HRESULT, Ausnahmetyp, Befehl und Zeile - damit laesst sich die Ursache ohne Sitzung dort finden
function Get-Fehlercode {
    param($Fehler)
    $ex = $Fehler.Exception
    while ($ex.InnerException -and $ex -is [System.Management.Automation.MethodInvocationException]) { $ex = $ex.InnerException }
    $teile = @(('0x{0:X8}' -f $ex.HResult), $ex.GetType().Name)
    if ($Fehler.FullyQualifiedErrorId) { $teile += [string]$Fehler.FullyQualifiedErrorId }
    if ($Fehler.InvocationInfo -and $Fehler.InvocationInfo.ScriptLineNumber) { $teile += ('Zeile {0}: {1}' -f $Fehler.InvocationInfo.ScriptLineNumber, (Format-Kurz ([string]$Fehler.InvocationInfo.Line).Trim() 160)) }
    return ($teile -join ' | ')
}

function Format-Groesse {
    param([int64] $Bytes)
    if ($Bytes -ge 1GB) { return '{0:N1} GB' -f ($Bytes / 1GB) }
    return '{0:N0} MB' -f ($Bytes / 1MB)
}

function Format-Kurz {
    param([string] $Text, [int] $Laenge = 200)
    $t = ([string]$Text -replace '\s+', ' ').Trim()
    if ($t.Length -le $Laenge) { return $t }
    return $t.Substring(0, $Laenge) + ' ...'
}

function Get-ExePath {
    param([string] $Befehl)
    if ([string]::IsNullOrWhiteSpace($Befehl)) { return $null }
    $c = [Environment]::ExpandEnvironmentVariables($Befehl.Trim())
    $c = $c -replace '^\\\?\?\\', ''
    $c = $c -replace '^(?i)\\SystemRoot\\', ($env:SystemRoot + '\')
    if ($c -match '^(?i)system32\\') { $c = Join-Path $env:SystemRoot $c }
    if ($c.StartsWith('"')) {
        $schluss = $c.IndexOf('"', 1)
        if ($schluss -gt 1) { return $c.Substring(1, $schluss - 1) }
        return $c.Trim('"')
    }
    if ($c -match '^(?i)(.+?\.(exe|sys|dll|cmd|bat|vbs|ps1|com|scr))(?=\s|,|$)') { return $Matches[1] }
    return $c
}

function Test-DateiDa {
    param([string] $Pfad)
    if ([string]::IsNullOrWhiteSpace($Pfad)) { return $false }
    if ($Pfad -notmatch '\\') { return $true }
    return [bool](Test-Path -LiteralPath $Pfad -PathType Leaf -ErrorAction SilentlyContinue)
}

function ConvertTo-TreiberVersion {
    param([string] $Text)
    $teil = @(([string]$Text).Trim() -split '\s+')[-1]
    $v = $null
    if ([version]::TryParse($teil, [ref]$v)) { return $v }
    return [version]'0.0'
}

function Invoke-System32 {
    param([string] $Exe, [string[]] $Argumente)
    $voll = Join-Path ([Environment]::GetFolderPath('System')) $Exe
    $ausgabe = (& $voll @Argumente 2>&1 | Out-String) -replace "`0", ''
    return [pscustomobject]@{ Code = $LASTEXITCODE; Text = $ausgabe }
}

# --- Registry mit Sicherung des alten Werts ---
function Get-RegBasis {
    param([string] $Hive)
    $h = if ($Hive -eq 'HKLM') { [Microsoft.Win32.RegistryHive]::LocalMachine } else { [Microsoft.Win32.RegistryHive]::CurrentUser }
    return [Microsoft.Win32.RegistryKey]::OpenBaseKey($h, [Microsoft.Win32.RegistryView]::Registry64)
}

function Set-RegDwordMitSicherung {
    param([string] $Hive, [string] $Pfad, [string] $Name, [int] $Wert)
    $basis = Get-RegBasis $Hive
    $vorher = $null; $art = $null
    $k = $basis.OpenSubKey($Pfad, $false)
    if ($k) {
        $vorher = $k.GetValue($Name, $null)
        if ($null -ne $vorher) { $art = [string]$k.GetValueKind($Name) }
        $k.Close()
    }
    if ($null -ne $vorher -and [string]$vorher -eq [string]$Wert) { return $null }
    if (-not $Script:IstSimulation) {
        $w = $basis.CreateSubKey($Pfad)
        $w.SetValue($Name, $Wert, [Microsoft.Win32.RegistryValueKind]::DWord)
        $w.Close()
    }
    return [ordered]@{ Typ = 'Registry'; Hive = $Hive; Pfad = $Pfad; Name = $Name; Existierte = ($null -ne $vorher); Vorher = $(if ($vorher -is [byte[]]) { [Convert]::ToBase64String($vorher) } else { $vorher }); Art = $art }
}

function Restore-RegWert {
    param($Eintrag)
    $basis = Get-RegBasis $Eintrag.Hive
    $w = $basis.CreateSubKey($Eintrag.Pfad)
    try {
        if (-not $Eintrag.Existierte) {
            $w.DeleteValue($Eintrag.Name, $false)
        } else {
            $art = if ($Eintrag.Art) { [Microsoft.Win32.RegistryValueKind]$Eintrag.Art } else { [Microsoft.Win32.RegistryValueKind]::DWord }
            $wert = $Eintrag.Vorher
            if ($art -eq [Microsoft.Win32.RegistryValueKind]::Binary) { $wert = [Convert]::FromBase64String([string]$wert) }
            elseif ($art -eq [Microsoft.Win32.RegistryValueKind]::DWord) { $wert = [int]$wert }
            $w.SetValue($Eintrag.Name, $wert, $art)
        }
    } finally { $w.Close() }
}

# --- Dienste ---
function ConvertTo-ScStart {
    param([string] $StartMode)
    switch ($StartMode) { 'Auto' { 'auto' } 'Manual' { 'demand' } 'Disabled' { 'disabled' } 'Boot' { 'boot' } 'System' { 'system' } default { 'demand' } }
}

function Set-DienstStart {
    param([string] $Name, [string] $ScStart)
    if ($Name -notmatch '^[\w\.\-]{1,120}$' -or $ScStart -notmatch '^(auto|demand|disabled|boot|system)$') { throw 'Ungültiger Dienstname oder Starttyp' }
    $r = Invoke-System32 'sc.exe' @('config', $Name, 'start=', $ScStart)
    if ($r.Code -ne 0) { throw "sc config meldet Fehler $($r.Code): $(Format-Kurz $r.Text 120)" }
}

# --- Treiberspeicher (einmal pro Lauf lesen) ---
function Get-TreiberSpeicher {
    if ($Script:Cache.ContainsKey('Speicher')) { return $Script:Cache['Speicher'] }
    $t = (Invoke-System32 'pnputil.exe' @('/enum-drivers')).Text
    $liste = foreach ($block in ($t -split '(?:\r?\n){2,}')) {
        if ($block -notmatch '(oem\d+\.inf)') { continue }
        $inf = $Matches[1]
        [pscustomobject]@{
            Inf = $inf
            Original = $(if ($block -match '(?m)^\s*(?:Originalname|Original Name)\s*:\s*(\S+)') { $Matches[1] } else { '' })
            Anbieter = $(if ($block -match '(?m)^\s*(?:Anbietername|Provider Name)\s*:\s*(.+?)\s*$') { $Matches[1] } else { '' })
            Klasse   = $(if ($block -match '(?m)^\s*(?:Klassenname|Class Name)\s*:\s*(.+?)\s*$') { $Matches[1] } else { '' })
            Version  = $(if ($block -match '(?m)^\s*(?:Treiberversion|Driver Version)\s*:\s*(.+?)\s*$') { $Matches[1] } else { '' })
        }
    }
    $Script:Cache['Speicher'] = @($liste)
    return $Script:Cache['Speicher']
}

function Get-GebundeneInfs {
    if ($Script:Cache.ContainsKey('Gebunden')) { return $Script:Cache['Gebunden'] }
    $g = @{}
    foreach ($d in @(Get-CimInstance -ClassName Win32_PnPSignedDriver -ErrorAction Stop)) { if ($d.InfName) { $g[([string]$d.InfName).ToLowerInvariant()] = $true } }
    $Script:Cache['Gebunden'] = $g
    return $g
}

# --- Wiederherstellungspunkt ---
function Confirm-Wiederherstellungspunkt {
    if ($Script:IstSimulation) { return [pscustomobject]@{ Ok = $true; Text = 'würde einen Wiederherstellungspunkt anlegen (Simulation)' } }
    $fehler = ''
    try { Checkpoint-Computer -Description 'Systemdiagnose: vor Bereinigung' -RestorePointType 'MODIFY_SETTINGS' -WarningAction SilentlyContinue -ErrorAction Stop } catch { $fehler = $_.Exception.Message }
    $letzter = @(Get-ComputerRestorePoint -ErrorAction SilentlyContinue | Sort-Object SequenceNumber -Descending)[0]
    if ($letzter) {
        $zeit = [Management.ManagementDateTimeConverter]::ToDateTime($letzter.CreationTime)
        if (((Get-Date) - $zeit).TotalHours -le 24) {
            $neu = ((Get-Date) - $zeit).TotalMinutes -le 10
            return [pscustomobject]@{ Ok = $true; Text = $(if ($neu) { "angelegt ($($zeit.ToString('dd.MM.yyyy HH:mm')))" } else { "vorhanden vom $($zeit.ToString('dd.MM.yyyy HH:mm')) - Windows erlaubt nur einen pro 24 Stunden" }) }
        }
    }
    return [pscustomobject]@{ Ok = $false; Text = 'kein Wiederherstellungspunkt möglich' + $(if ($fehler) { ": $fehler" } else { ' (Computerschutz aus oder zu wenig Platz?)' }) }
}

# ============================================================================
#  Aktionen (fester Katalog)
# ============================================================================

function Invoke-TreiberpaketEntfernen {
    param($P)
    $inf = [string]$P.Inf
    if ($inf -notmatch '^oem\d{1,4}\.inf$') { return New-Ergebnis 'abgelehnt' "Ungültiger Paketname '$inf'" }
    $speicher = @(Get-TreiberSpeicher)
    $paket = $speicher | Where-Object { $_.Inf -eq $inf } | Select-Object -First 1
    if (-not $paket) { return New-Ergebnis 'übersprungen' "$inf ist nicht mehr vorhanden" }
    $geb = Get-GebundeneInfs
    if ($geb.ContainsKey($inf.ToLowerInvariant())) { return New-Ergebnis 'abgelehnt' "$inf wird von einem Gerät benutzt" }
    $istGrafik = $paket.Klasse -match '(?i)^(Display|Grafik)'
    $ersatz = @($speicher | Where-Object { $_.Inf -ne $inf -and $geb.ContainsKey($_.Inf.ToLowerInvariant()) -and (($_.Original -eq $paket.Original) -or ($istGrafik -and $_.Klasse -match '(?i)^(Display|Grafik)' -and $_.Anbieter -eq $paket.Anbieter)) })
    if ($ersatz.Count -eq 0) { return New-Ergebnis 'abgelehnt' "Für $inf ist keine aktive Ersatzversion gebunden - wird nicht angefasst" }
    $max = @($ersatz | ForEach-Object { ConvertTo-TreiberVersion $_.Version } | Sort-Object -Descending)[0]
    if ((ConvertTo-TreiberVersion $paket.Version) -gt $max) { return New-Ergebnis 'abgelehnt' "$inf ist neuer als die aktive Version - wird nicht angefasst" }
    if (-not $Script:WiederherstellungOk) { return New-Ergebnis 'übersprungen' 'Kein Wiederherstellungspunkt verfügbar - Treiberpakete werden ohne Sicherung nicht entfernt.' }
    $text = "$inf ($($paket.Original) $($paket.Version))"
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' "würde $text entfernen" }
    $r = Invoke-System32 'pnputil.exe' @('/delete-driver', $inf)
    if ($r.Code -eq 0) { return New-Ergebnis 'erledigt' "$text entfernt" }
    return New-Ergebnis 'fehlgeschlagen' "Windows hat das Entfernen von $text verweigert (wird evtl. doch gebraucht): $(Format-Kurz $r.Text 140)"
}

function Invoke-TreiberdienstDeaktivieren {
    param($P)
    $name = [string]$P.Name
    if ($name -notmatch '^[\w\.\-]{1,120}$') { return New-Ergebnis 'abgelehnt' 'Ungültiger Treibername' }
    $d = Get-CimInstance -ClassName Win32_SystemDriver -Filter "Name='$name'" -ErrorAction SilentlyContinue
    if (-not $d) { return New-Ergebnis 'übersprungen' "Treiber $name existiert nicht mehr" }
    if (Test-DateiDa (Get-ExePath $d.PathName)) { return New-Ergebnis 'abgelehnt' "Die Datei von $name existiert (wieder) - der Treiber ist nicht verwaist" }
    if ($d.StartMode -eq 'Disabled') { return New-Ergebnis 'übersprungen' "$name ist bereits deaktiviert" }
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' "würde Treiber $name deaktivieren" }
    Set-DienstStart $name 'disabled'
    return New-Ergebnis 'erledigt' "Treiber $name deaktiviert" @([ordered]@{ Typ = 'Start'; Name = $name; Vorher = $d.StartMode; LiefVorher = $false })
}

function Invoke-DienstDeaktivieren {
    param($P)
    $name = [string]$P.Name
    if ($name -notmatch '^[\w\.\-]{1,120}$') { return New-Ergebnis 'abgelehnt' 'Ungültiger Dienstname' }
    $s = Get-CimInstance -ClassName Win32_Service -Filter "Name='$name'" -ErrorAction SilentlyContinue
    if (-not $s) { return New-Ergebnis 'übersprungen' "Dienst $name existiert nicht mehr" }
    $pfad = Get-ExePath $s.PathName
    $da = Test-DateiDa $pfad
    if ($da -and $pfad -match '\\') {
        $firma = ''
        try { $firma = [string](Get-Item -LiteralPath $pfad -Force).VersionInfo.CompanyName } catch { }
        if ($firma -match '^Microsoft') { return New-Ergebnis 'abgelehnt' "$name ist ein Windows-Dienst - wird nie abgeschaltet" }
    }
    if ($da -and $Script:DeaktivierbareDienste -notcontains $name) { return New-Ergebnis 'abgelehnt' "$name ist weder verwaist noch im PC-Profil zum Abschalten freigegeben" }
    if ($s.StartMode -eq 'Disabled') { return New-Ergebnis 'übersprungen' "$name ist bereits deaktiviert" }
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' "würde Dienst $name deaktivieren" }
    $lief = ($s.State -eq 'Running')
    if ($lief) { try { Stop-Service -Name $name -Force -ErrorAction Stop } catch { } }
    Set-DienstStart $name 'disabled'
    return New-Ergebnis 'erledigt' "Dienst $name deaktiviert" @([ordered]@{ Typ = 'Start'; Name = $name; Vorher = $s.StartMode; LiefVorher = $lief })
}

function Invoke-DienstStarten {
    param($P)
    $name = [string]$P.Name
    if ($Script:StartbareDienste -notcontains $name) { return New-Ergebnis 'abgelehnt' "$name ist nicht zum Starten freigegeben" }
    $s = Get-CimInstance -ClassName Win32_Service -Filter "Name='$name'" -ErrorAction SilentlyContinue
    if (-not $s) { return New-Ergebnis 'fehlgeschlagen' "Dienst $name ist nicht installiert - zugehörige Software neu installieren" }
    if ($s.State -eq 'Running' -and $s.StartMode -eq 'Auto') { return New-Ergebnis 'übersprungen' "$name läuft bereits" }
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' "würde $name auf Automatisch stellen und starten" }
    # Sicherung immer (auch bei Auto/gestoppt): Rueckgaengig stellt Starttyp UND Laufzustand wieder her
    $undo = @([ordered]@{ Typ = 'Start'; Name = $name; Vorher = $s.StartMode; LiefVorher = ($s.State -eq 'Running') })
    if ($s.StartMode -ne 'Auto') { Set-DienstStart $name 'auto' }
    try { Start-Service -Name $name -ErrorAction Stop }
    catch { return New-Ergebnis 'fehlgeschlagen' "$name startet nicht: $($_.Exception.Message)$(if ($s.StartMode -ne 'Auto') { ' (Starttyp wurde schon auf Automatisch gestellt - rückgängig möglich)' })" $(if ($s.StartMode -ne 'Auto') { $undo } else { @() }) }
    return New-Ergebnis 'erledigt' "$name gestartet (Start: Automatisch). Falls das RGB nicht sofort zurückkommt: G HUB neu starten." $undo
}

function Invoke-AufgabeDeaktivieren {
    param($P)
    $pfad = [string]$P.Pfad
    if ($pfad -notmatch '^\\[^\r\n"<>|]{1,250}$' -or $pfad -match '^(?i)\\Microsoft\\') { return New-Ergebnis 'abgelehnt' 'Ungültiger Aufgabenpfad (Microsoft-Aufgaben werden nie angefasst)' }
    $i = $pfad.LastIndexOf('\')
    $tp = $pfad.Substring(0, $i + 1)
    $tn = $pfad.Substring($i + 1)
    $t = Get-ScheduledTask -TaskPath $tp -TaskName $tn -ErrorAction SilentlyContinue
    if (-not $t) { return New-Ergebnis 'übersprungen' "Aufgabe $pfad existiert nicht mehr" }
    if ([string]$t.State -eq 'Disabled') { return New-Ergebnis 'übersprungen' "Aufgabe $pfad ist bereits deaktiviert" }
    $a = @($t.Actions)[0]
    if ($a -and $a.Execute -and (Test-DateiDa (Get-ExePath ([string]$a.Execute)))) { return New-Ergebnis 'abgelehnt' "Das Programm der Aufgabe existiert (wieder) - wird nicht deaktiviert" }
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' "würde Aufgabe $pfad deaktivieren" }
    [void](Disable-ScheduledTask -TaskPath $tp -TaskName $tn)
    return New-Ergebnis 'erledigt' "Aufgabe $pfad deaktiviert" @([ordered]@{ Typ = 'Aufgabe'; TaskPath = $tp; TaskName = $tn })
}

function Invoke-AutostartDeaktivieren {
    param($P)
    $ort = [string]$P.Ort
    $name = [string]$P.Name
    if (-not $Script:AutostartOrte.ContainsKey($ort) -or $name.Length -lt 1 -or $name.Length -gt 260) { return New-Ergebnis 'abgelehnt' 'Ungültiger Autostart-Eintrag' }
    $o = $Script:AutostartOrte[$ort]
    if ($o.Hive -eq 'HKCU' -and -not $Script:GleicherBenutzer) { return New-Ergebnis 'übersprungen' 'Benutzer-Einstellung - die Bereinigung läuft unter einem anderen Konto als die App' }
    $basis = Get-RegBasis $o.Hive
    $k = $basis.OpenSubKey($o.Run, $false)
    $befehl = if ($k) { $k.GetValue($name, $null) } else { $null }
    if ($k) { $k.Close() }
    if ($null -eq $befehl) { return New-Ergebnis 'übersprungen' "Autostart '$name' existiert nicht mehr" }
    if (Test-DateiDa (Get-ExePath ([string]$befehl))) { return New-Ergebnis 'abgelehnt' "Das Programm von '$name' existiert (wieder) - wird nicht deaktiviert" }
    $f = $basis.OpenSubKey($o.Freigabe, $false)
    $vorher = if ($f) { $f.GetValue($name, $null) } else { $null }
    if ($f) { $f.Close() }
    if ($vorher -is [byte[]] -and $vorher.Length -gt 0 -and ($vorher[0] -band 1)) { return New-Ergebnis 'übersprungen' "'$name' ist bereits deaktiviert" }
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' "würde Autostart '$name' deaktivieren" }
    $neu = [byte[]](@(3, 0, 0, 0) + [BitConverter]::GetBytes((Get-Date).ToFileTime()))
    $w = $basis.CreateSubKey($o.Freigabe)
    $w.SetValue($name, $neu, [Microsoft.Win32.RegistryValueKind]::Binary)
    $w.Close()
    return New-Ergebnis 'erledigt' "Autostart '$name' deaktiviert (wie im Task-Manager)" @([ordered]@{ Typ = 'Registry'; Hive = $o.Hive; Pfad = $o.Freigabe; Name = $name; Existierte = ($null -ne $vorher); Vorher = $(if ($vorher -is [byte[]]) { [Convert]::ToBase64String($vorher) } else { $null }); Art = 'Binary' })
}

# ----------------------------------------------------------------------------
#  Reinigung (Kategorien aus Reinigung.psd1)
# ----------------------------------------------------------------------------

# Loescht ueber ein Handle: Die Datei wird geoeffnet, ohne einer Verknuepfung am Ende zu folgen, dann wird
# ihr TATSAECHLICHER Pfad geprueft (alle Junctions/Symlinks aufgeloest). Liegt sie in Wahrheit ausserhalb
# des Reinigungsordners, wird sie nicht angeruehrt. Erst dann wird ueber genau dieses Handle geloescht.
$Script:LoescherQuelle = @'
using System;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

public static class SdLoescher {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern SafeFileHandle CreateFileW(string name, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern uint GetFinalPathNameByHandleW(SafeFileHandle handle, StringBuilder path, uint length, uint flags);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool SetFileInformationByHandle(SafeFileHandle handle, int infoClass, ref uint info, uint size);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool SetFileInformationByHandle(SafeFileHandle handle, int infoClass, ref byte info, uint size);

    const uint DELETE = 0x00010000, READ_ATTRIBUTES = 0x80, SHARE_ALL = 7, OPEN_EXISTING = 3;
    const uint OPEN_REPARSE_POINT = 0x00200000, BACKUP_SEMANTICS = 0x02000000;

    static string Endpfad(SafeFileHandle h) {
        var sb = new StringBuilder(1024);
        uint n = GetFinalPathNameByHandleW(h, sb, (uint)sb.Capacity, 0);
        if (n == 0 || n >= sb.Capacity) return null;
        string p = sb.ToString();
        if (p.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase)) return null;
        if (p.StartsWith(@"\\?\")) p = p.Substring(4);
        return p;
    }

    // Tatsaechlicher Pfad eines Ordners (alle Verknuepfungen aufgeloest) oder null
    public static string EchterPfad(string pfad) {
        using (var h = CreateFileW(@"\\?\" + pfad, READ_ATTRIBUTES, SHARE_ALL, IntPtr.Zero, OPEN_EXISTING, BACKUP_SEMANTICS, IntPtr.Zero)) {
            return h.IsInvalid ? null : Endpfad(h);
        }
    }

    // 0 = geloescht, 1 = abgelehnt (liegt tatsaechlich ausserhalb der Wurzel), 2 = nicht moeglich (gesperrt, Ordner nicht leer ...)
    public static int Loeschen(string pfad, string echteWurzel) {
        using (var h = CreateFileW(@"\\?\" + pfad, DELETE | READ_ATTRIBUTES, SHARE_ALL, IntPtr.Zero, OPEN_EXISTING, OPEN_REPARSE_POINT | BACKUP_SEMANTICS, IntPtr.Zero)) {
            if (h.IsInvalid) return 2;
            string echt = Endpfad(h);
            if (echt == null) return 2;
            if (!echt.StartsWith(echteWurzel.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase)) return 1;
            uint flags = 0x1 | 0x10;                                           // DELETE | IGNORE_READONLY_ATTRIBUTE
            if (SetFileInformationByHandle(h, 21, ref flags, 4)) return 0;    // FileDispositionInfoEx
            byte loeschen = 1;
            if (SetFileInformationByHandle(h, 4, ref loeschen, 1)) return 0;  // FileDispositionInfo
            return 2;
        }
    }
}
'@

$Script:ReinigungsBasen = @('LocalAppData', 'AppData', 'Windows', 'ProgramData', 'Systemlaufwerk')
$Script:BenutzerBasen   = @('LocalAppData', 'AppData')

function Initialize-Loescher {
    if (-not ('SdLoescher' -as [type])) { Add-Type -TypeDefinition $Script:LoescherQuelle -Language CSharp }
}

function Get-ReinigungsKatalog {
    if (-not $Script:Cache.ContainsKey('Reinigung')) {
        $def = Import-PowerShellDataFile -LiteralPath (Join-Path $Script:ToolOrdner 'Reinigung.psd1')
        if ($def.Format -ne 'Systemdiagnose-Reinigung/1') { throw 'Reinigung.psd1 hat ein unbekanntes Format.' }
        $Script:Cache['Reinigung'] = @($def.Kategorien)
    }
    return $Script:Cache['Reinigung']
}

function Get-ReinigungsBasis {
    param([string] $Basis)
    switch ($Basis) {
        'LocalAppData'   { return [Environment]::GetFolderPath('LocalApplicationData') }
        'AppData'        { return [Environment]::GetFolderPath('ApplicationData') }
        'Windows'        { return [Environment]::GetFolderPath('Windows') }
        'ProgramData'    { return [Environment]::GetFolderPath('CommonApplicationData') }
        'Systemlaufwerk' { return [IO.Path]::GetPathRoot([Environment]::GetFolderPath('Windows')) }
    }
    return $null
}

# Gleiche Formatregeln wie in der Diagnose - '' = in Ordnung, sonst der Grund
function Test-ZielFormat {
    param($Ziel)
    $pfad = [string]$Ziel.Pfad
    $muster = [string]$Ziel.Muster
    if ($Script:ReinigungsBasen -notcontains [string]$Ziel.Basis) { return "unbekannte Basis '$($Ziel.Basis)'" }
    if ($pfad -and ($pfad -notmatch '^[\w \-\.\(\)\*]+(\\[\w \-\.\(\)\*]+)*$' -or @($pfad.Split('\') | Where-Object { $_ -match '^\.+$' }).Count)) { return "ungültiger Pfad '$pfad'" }
    if ($muster -notmatch '^[\w\-\.\*]+$' -or $muster -match '^\.+$') { return "ungültiges Muster '$muster'" }
    if (-not $pfad -and $muster.Contains('*')) { return 'direkt in der Basis nur mit genauem Dateinamen' }
    return ''
}

function Resolve-ZielOrdner {
    param($Ziel)
    $basis = Get-ReinigungsBasis ([string]$Ziel.Basis)
    if (-not $basis) { return }
    $aktuell = @($(if ($basis.Length -gt 3) { $basis.TrimEnd('\') } else { $basis }))
    $pfad = [string]$Ziel.Pfad
    if (-not $pfad) { return $aktuell }
    foreach ($teil in $pfad.Split('\')) {
        $naechste = New-Object System.Collections.Generic.List[string]
        foreach ($o in $aktuell) {
            try { foreach ($d in ([IO.DirectoryInfo]::new($o)).GetDirectories($teil)) { if (-not ($d.Attributes -band [IO.FileAttributes]::ReparsePoint)) { $naechste.Add($d.FullName) } } } catch { }
        }
        $aktuell = $naechste.ToArray()
        if ($aktuell.Count -eq 0) { return }
    }
    return $aktuell
}

# Harte Grenze unabhaengig von Reinigung.psd1: Diese Ordner werden nie gereinigt
function Get-GeschuetzteOrdner {
    if ($Script:Cache.ContainsKey('Geschuetzt')) { return $Script:Cache['Geschuetzt'] }
    $win = [Environment]::GetFolderPath('Windows')
    $profil = [Environment]::GetFolderPath('UserProfile')
    # System- und persoenliche Ordner: Ziel darf keiner davon sein, nicht darin liegen und keinen enthalten
    $nie = @((Join-Path $win 'System32'), (Join-Path $win 'SysWOW64'), (Join-Path $win 'WinSxS'), (Join-Path $win 'servicing'),
        (Join-Path $win 'Installer'), (Join-Path $win 'assembly'), (Join-Path $win 'Microsoft.NET'), (Join-Path $win 'Boot'))
    $nie += @(foreach ($f in 'ProgramFiles', 'ProgramFilesX86', 'Desktop', 'MyDocuments', 'MyPictures', 'MyMusic', 'MyVideos', 'Favorites',
            'CommonDocuments', 'CommonDesktopDirectory', 'StartMenu', 'CommonStartMenu') { [Environment]::GetFolderPath($f) })
    $nie += @('Downloads', 'OneDrive', 'Saved Games', 'Contacts', 'Links', 'Searches' | ForEach-Object { Join-Path $profil $_ })
    if ($env:OneDrive) { $nie += $env:OneDrive }
    # Sammelordner: Ziel darf keiner davon sein und keinen enthalten (Unterordner wie AppData\Local\Temp sind erlaubt)
    $kein = @(foreach ($f in 'UserProfile', 'LocalApplicationData', 'ApplicationData', 'CommonApplicationData', 'Windows') { [Environment]::GetFolderPath($f) })
    $kein += Split-Path -Parent $profil
    $Script:Cache['Geschuetzt'] = [pscustomobject]@{
        Nie  = @($nie | Where-Object { $_ } | ForEach-Object { $_.TrimEnd('\') })
        Kein = @($kein | Where-Object { $_ } | ForEach-Object { $_.TrimEnd('\') })
    }
    return $Script:Cache['Geschuetzt']
}

# '' = Ordner darf gereinigt werden, sonst der Grund. NurGenaueDatei: einzelne Datei mit festem Namen
# (z. B. MEMORY.DMP direkt im Windows-Ordner) - dann darf der Ordner selbst ein geschuetzter sein.
function Test-ReinigungsWurzel {
    param([string] $Pfad, [bool] $NurGenaueDatei = $false)
    if ($Pfad -notmatch '^[A-Za-z]:\\') { return 'kein lokaler Pfad' }
    $p = $Pfad.TrimEnd('\')
    if ($p -match '^[A-Za-z]:$') { return 'Laufwerkswurzel' }
    $g = Get-GeschuetzteOrdner
    $ic = [StringComparison]::OrdinalIgnoreCase
    foreach ($n in $g.Nie) {
        if ($p -ieq $n -or $p.StartsWith($n + '\', $ic)) { return "liegt in einem geschützten Ordner ($n)" }
        if (-not $NurGenaueDatei -and $n.StartsWith($p + '\', $ic)) { return "würde einen geschützten Ordner treffen ($n)" }
    }
    if (-not $NurGenaueDatei) {
        foreach ($k in $g.Kein) { if ($p -ieq $k -or $k.StartsWith($p + '\', $ic)) { return "würde einen geschützten Ordner treffen ($k)" } }
    }
    return ''
}

# Loescht (bzw. zaehlt) passende Dateien unterhalb der Wurzel, danach leere Unterordner von innen
# nach aussen - nie die Wurzel selbst. Verknuepfungen werden nie betreten.
function Invoke-OrdnerReinigung {
    param([string] $Wurzel, [string] $EchteWurzel, [string] $Muster, [bool] $Rekursiv, [int] $MinAlterStunden, [switch] $NurZaehlen)
    $zaehlen = $NurZaehlen -or $Script:IstSimulation
    $grenze = [DateTime]::UtcNow.AddHours(-$MinAlterStunden)
    $s = [pscustomobject]@{ Dateien = 0; Bytes = [int64]0; Gesperrt = 0; Abgelehnt = 0 }
    $ordner = New-Object System.Collections.Generic.List[object]
    $ordner.Add([IO.DirectoryInfo]::new($Wurzel))
    for ($i = 0; $i -lt $ordner.Count; $i++) {
        try { $eintraege = $ordner[$i].GetFileSystemInfos() } catch { $s.Gesperrt++; continue }
        foreach ($e in $eintraege) {
            try {
                if ($e.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
                if ($e -is [IO.DirectoryInfo]) { if ($Rekursiv) { $ordner.Add($e) }; continue }
                if ($e.Name -notlike $Muster -or $e.LastWriteTimeUtc -gt $grenze) { continue }
                $laenge = $e.Length
                if ($zaehlen) { $s.Dateien++; $s.Bytes += $laenge; continue }
                switch ([SdLoescher]::Loeschen($e.FullName, $EchteWurzel)) {
                    0       { $s.Dateien++; $s.Bytes += $laenge }
                    1       { $s.Abgelehnt++ }
                    default { $s.Gesperrt++ }
                }
            } catch { $s.Gesperrt++ }   # z. B. ueberlange Pfade
        }
    }
    # Kinder stehen in der Liste immer hinter ihren Eltern -> rueckwaerts = von innen nach aussen
    if ($Rekursiv -and -not $zaehlen) {
        for ($i = $ordner.Count - 1; $i -ge 1; $i--) {
            if ($ordner[$i].LastWriteTimeUtc -le $grenze) { [void][SdLoescher]::Loeschen($ordner[$i].FullName, $EchteWurzel) }
        }
    }
    return $s
}

function Test-NeustartAusstehend {
    return (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
        (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
}

function New-ReinigungsErgebnis {
    param([string] $Name, $Summe, [string[]] $Abgelehnt)
    $teile = @()
    if ($Summe.Gesperrt) { $teile += "$($Summe.Gesperrt) gesperrt bzw. in Benutzung - übersprungen" }
    if ($Summe.Abgelehnt) { $teile += "$($Summe.Abgelehnt) Datei(en) abgelehnt (lagen über eine Verknüpfung außerhalb)" }
    if ($Abgelehnt.Count) { $teile += 'abgelehnt: ' + ($Abgelehnt -join '; ') }
    $rest = if ($teile.Count) { ' - ' + ($teile -join ', ') } else { '' }
    if ($Summe.Dateien -eq 0) {
        if ($Abgelehnt.Count -and -not $Summe.Gesperrt) { return New-Ergebnis 'abgelehnt' ("$Name - nichts gelöscht" + $rest) }
        return New-Ergebnis 'übersprungen' ("$Name - nichts zu löschen" + $rest)
    }
    $kern = '{0} ({1:N0} Dateien)' -f (Format-Groesse $Summe.Bytes), $Summe.Dateien
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' ("würde $kern freigeben" + $rest) -Bytes $Summe.Bytes }
    return New-Ergebnis 'erledigt' ("$kern freigegeben" + $rest) -Bytes $Summe.Bytes
}

function Invoke-PapierkorbLeeren {
    param($K)
    if (-not $Script:GleicherBenutzer) { return New-Ergebnis 'übersprungen' 'Der Papierkorb gehört zu einem anderen Konto als die App' }
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $koerbe = @(foreach ($lw in @([IO.DriveInfo]::GetDrives() | Where-Object { [string]$_.DriveType -eq 'Fixed' -and $_.IsReady })) {
        $korb = $lw.RootDirectory.FullName + '$Recycle.Bin\' + $sid
        if ([IO.Directory]::Exists($korb)) { $korb }
    })
    # Nur die Eintraege des Papierkorbs ($I = Info, $R = Inhalt) - nie desktop.ini oder andere Konten
    $durchgang = {
        param([bool] $Zaehlen)
        $s = [pscustomobject]@{ Dateien = 0; Bytes = [int64]0; Gesperrt = 0; Abgelehnt = 0 }
        foreach ($korb in $koerbe) {
            $echt = [SdLoescher]::EchterPfad($korb)
            if (-not $echt -or $echt -notmatch '^[A-Za-z]:\\\$Recycle\.Bin\\S-1-5-[\d\-]+$' -or $echt -ine $korb) { $s.Abgelehnt++; continue }
            foreach ($e in ([IO.DirectoryInfo]::new($korb)).GetFileSystemInfos('$*')) {
                if ($e.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
                if ($e -is [IO.DirectoryInfo]) {
                    $r = Invoke-OrdnerReinigung $e.FullName $echt '*' $true 0 -NurZaehlen:$Zaehlen
                    $s.Dateien += $r.Dateien; $s.Bytes += $r.Bytes; $s.Gesperrt += $r.Gesperrt; $s.Abgelehnt += $r.Abgelehnt
                    if (-not $Zaehlen) { [void][SdLoescher]::Loeschen($e.FullName, $echt) }
                } elseif ($Zaehlen) {
                    $s.Dateien++; $s.Bytes += $e.Length
                } else {
                    $laenge = $e.Length
                    switch ([SdLoescher]::Loeschen($e.FullName, $echt)) { 0 { $s.Dateien++; $s.Bytes += $laenge } 1 { $s.Abgelehnt++ } default { $s.Gesperrt++ } }
                }
            }
        }
        return $s
    }
    $vorher = & $durchgang $true
    if ($vorher.Dateien -eq 0) { return New-Ergebnis 'übersprungen' 'Der Papierkorb ist bereits leer' }
    if ($Script:IstSimulation) { return New-ReinigungsErgebnis $K.Name $vorher @() }
    # 1. Windows-eigene Funktion (aktualisiert auch das Papierkorb-Symbol), 2. verwaiste Reste selbst loeschen
    try { Clear-RecycleBin -Force -ErrorAction Stop } catch { }
    [void](& $durchgang $false)
    $nachher = & $durchgang $true
    $frei = [pscustomobject]@{ Dateien = $vorher.Dateien - $nachher.Dateien; Bytes = $vorher.Bytes - $nachher.Bytes; Gesperrt = $nachher.Dateien; Abgelehnt = 0 }
    return New-ReinigungsErgebnis $K.Name $frei @()
}

function Get-LieferCache {
    # Echte Cache-Groesse (wie Datentraegerbereinigung); die Jobliste zaehlt laufende/angeheftete Downloads mit
    try { $s = Get-DeliveryOptimizationPerfSnap -ErrorAction Stop -WarningAction SilentlyContinue; if ($null -ne $s.CacheSizeBytes) { return [int64]$s.CacheSizeBytes } } catch { }
    return [int64]((@(Get-DeliveryOptimizationStatus -ErrorAction Stop -WarningAction SilentlyContinue) | Measure-Object -Property FileSizeInCache -Sum).Sum)
}

function Invoke-LieferoptimierungLeeren {
    param($K)
    $jobs = @(Get-DeliveryOptimizationStatus -ErrorAction Stop -WarningAction SilentlyContinue)
    if (@($jobs | Where-Object { [string]$_.Status -eq 'Downloading' }).Count) { return New-Ergebnis 'übersprungen' 'Windows lädt gerade Updates herunter - später erneut reinigen' }
    $bytes = Get-LieferCache
    if ($bytes -le 0) { return New-Ergebnis 'übersprungen' 'Der Zwischenspeicher ist bereits leer' }
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' ('würde bis zu {0} freigeben' -f (Format-Groesse $bytes)) -Bytes $bytes }
    Delete-DeliveryOptimizationCache -Force -ErrorAction Stop
    # PerfSnap zieht sofort nach (die Jobliste erst verzoegert) - freigegeben = gemessene Differenz
    $nachher = Get-LieferCache
    $frei = [Math]::Max([int64]0, $bytes - $nachher)
    $rest = if ($nachher -gt 0) { (' - {0} behält Windows (laufende oder angeheftete Downloads)' -f (Format-Groesse $nachher)) } else { '' }
    return New-Ergebnis 'erledigt' ('{0} freigegeben{1}' -f (Format-Groesse $frei), $rest) -Bytes $frei
}

function Invoke-Reinigen {
    param($P)
    $id = [string]$P.Kategorie
    if ($id -notmatch '^[A-Za-z]{2,40}$') { return New-Ergebnis 'abgelehnt' 'Ungültige Kategorie' }
    $k = @(Get-ReinigungsKatalog) | Where-Object { $_.Id -eq $id } | Select-Object -First 1
    if (-not $k) { return New-Ergebnis 'abgelehnt' "Kategorie '$id' steht nicht in Reinigung.psd1" }
    $art = if ($k.Art) { [string]$k.Art } else { 'Dateien' }
    if ($art -notin @('Dateien', 'Papierkorb', 'Lieferoptimierung')) { return New-Ergebnis 'abgelehnt' "Unbekannte Reinigungsart '$art'" }
    $laeuft = @(@($k.Prozesse) | Where-Object { [string]$_ -match '^[\w\-\.]{1,60}$' -and (Get-Process -Name ([string]$_) -ErrorAction SilentlyContinue) })
    if ($laeuft.Count) { return New-Ergebnis 'übersprungen' "$($k.Name): läuft gerade ($($laeuft -join ', ')) - bitte ganz schließen und erneut reinigen" }
    if ($k.Bedingung -eq 'KeinNeustartAusstehend' -and (Test-NeustartAusstehend)) { return New-Ergebnis 'übersprungen' 'Updates warten auf einen Neustart - danach erneut reinigen' }
    Initialize-Loescher
    if ($art -eq 'Papierkorb') { return Invoke-PapierkorbLeeren $k }
    if ($art -eq 'Lieferoptimierung') { return Invoke-LieferoptimierungLeeren $k }

    $summe = [pscustomobject]@{ Dateien = 0; Bytes = [int64]0; Gesperrt = 0; Abgelehnt = 0 }
    $abgelehnt = New-Object System.Collections.Generic.List[string]
    foreach ($z in @($k.Ziele)) {
        $grund = Test-ZielFormat $z
        if ($grund) { return New-Ergebnis 'abgelehnt' "Reinigung.psd1, Kategorie $id`: $grund - nichts gelöscht" }
        if ($Script:BenutzerBasen -contains [string]$z.Basis -and -not $Script:GleicherBenutzer) { return New-Ergebnis 'übersprungen' 'Benutzer-Daten - die Bereinigung läuft unter einem anderen Konto als die App' }
        $genau = (-not $z.Rekursiv) -and -not ([string]$z.Muster).Contains('*')
        foreach ($o in @(Resolve-ZielOrdner $z)) {
            $g = Test-ReinigungsWurzel $o $genau
            $echt = if (-not $g) { [SdLoescher]::EchterPfad($o) } else { $null }
            if (-not $g -and -not $echt) { $summe.Gesperrt++; continue }   # kein Zugriff auf den Ordner
            if (-not $g) { $g = Test-ReinigungsWurzel $echt $genau }
            if ($g) { $abgelehnt.Add("$o ($g)"); Write-Protokoll "REINIGUNG abgelehnt: $o - $g"; continue }
            $r = Invoke-OrdnerReinigung $o $echt ([string]$z.Muster) ([bool]$z.Rekursiv) ([int]$z.MinAlterStunden)
            $summe.Dateien += $r.Dateien; $summe.Bytes += $r.Bytes; $summe.Gesperrt += $r.Gesperrt; $summe.Abgelehnt += $r.Abgelehnt
        }
    }
    return New-ReinigungsErgebnis $k.Name $summe @($abgelehnt)
}

function Invoke-WiederherstellungspunktAnlegen {
    param($P)
    $r = Confirm-Wiederherstellungspunkt
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' $r.Text }
    if ($r.Ok) { $Script:WiederherstellungOk = $true; return New-Ergebnis 'erledigt' "Wiederherstellungspunkt $($r.Text)" }
    return New-Ergebnis 'fehlgeschlagen' $r.Text
}

function Invoke-DismReparieren {
    param($P)
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' 'würde DISM /RestoreHealth ausführen' }
    Write-Status 'DISM repariert den Komponentenspeicher - das kann 10-20 Minuten dauern ...'
    $r = Invoke-System32 'dism.exe' @('/Online', '/Cleanup-Image', '/RestoreHealth')
    if ($r.Code -eq 0) { return New-Ergebnis 'erledigt' 'DISM: Komponentenspeicher repariert bzw. in Ordnung' }
    return New-Ergebnis 'fehlgeschlagen' "DISM meldet Fehler $($r.Code): $(Format-Kurz $r.Text 160)"
}

function Invoke-SfcReparieren {
    param($P)
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' 'würde sfc /scannow ausführen' }
    Write-Status 'SFC prüft und repariert die Systemdateien - das kann 10-15 Minuten dauern ...'
    $r = Invoke-System32 'sfc.exe' @('/scannow')
    if ($r.Text -match '(?i)keine Integrit|did not find any integrity violations') { return New-Ergebnis 'erledigt' 'SFC: keine Integritätsverletzungen gefunden' }
    if ($r.Text -match '(?i)erfolgreich repariert|successfully repaired') { return New-Ergebnis 'erledigt' 'SFC: beschädigte Dateien wurden repariert' }
    return New-Ergebnis 'fehlgeschlagen' "SFC: $(Format-Kurz ($r.Text -split "`n" | Where-Object { $_.Trim() } | Select-Object -Last 2 | Out-String) 200)"
}

function Invoke-EnergieplanSetzen {
    param($P)
    # Entweder exakter Name (eingelernter Sollwert, Vergleich per Gleichheit - kein Regex) oder das eingebaute einfache Muster
    $exakt = [string]$P.Name
    $muster = [string]$P.Muster
    if ($exakt) {
        if ($exakt.Length -lt 2 -or $exakt.Length -gt 100 -or $exakt -match '[\x00-\x1F]') { return New-Ergebnis 'abgelehnt' 'Ungültiger Name für den Energieplan' }
        $muster = $exakt
    }
    # Muster ist ein einfacher regulaerer Ausdruck (nur Buchstaben, Ziffern, Leerzeichen, Punkt, Stern, Bindestrich)
    elseif ($muster -notmatch '^[\w \-\.\*]{2,60}$') { return New-Ergebnis 'abgelehnt' 'Ungültiges Suchmuster für den Energieplan' }
    $liste = (Invoke-System32 'powercfg.exe' @('/list')).Text
    $plaene = @(foreach ($z in ($liste -split "`r?`n")) {
        # Name bis zur LETZTEN Klammer (Plannamen duerfen selbst Klammern enthalten)
        if ($z -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\s+\((.+)\)\s*(\*)?\s*$') { [pscustomobject]@{ Guid = $Matches[1]; Name = $Matches[2]; Aktiv = [bool]$Matches[3] } }
    })
    $treffer = @(if ($exakt) { $plaene | Where-Object { $_.Name -eq $exakt } } else { $plaene | Where-Object { $_.Name -match $muster } })
    if ($treffer.Count -eq 0) { return New-Ergebnis 'fehlgeschlagen' "Kein Energieplan passt auf '$muster' - ggf. AMD-Chipsatztreiber neu installieren" }
    $aktiverTreffer = @($treffer | Where-Object { $_.Aktiv })
    if ($aktiverTreffer.Count) { return New-Ergebnis 'übersprungen' "'$($aktiverTreffer[0].Name)' ist bereits aktiv" }
    if ($treffer.Count -gt 1) { return New-Ergebnis 'abgelehnt' ("Mehrere Pläne passen auf '$muster' (" + (@($treffer.Name) -join ', ') + ') - nicht eindeutig, es wird keiner gesetzt') }
    $ziel = $treffer[0]
    $vorher = @($plaene | Where-Object { $_.Aktiv })[0]
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' "würde '$($ziel.Name)' aktivieren" }
    $r = Invoke-System32 'powercfg.exe' @('/setactive', $ziel.Guid)
    if ($r.Code -ne 0) { return New-Ergebnis 'fehlgeschlagen' "powercfg meldet Fehler $($r.Code)" }
    return New-Ergebnis 'erledigt' "Energieplan '$($ziel.Name)' aktiviert" @($(if ($vorher) { [ordered]@{ Typ = 'Energieplan'; Guid = $vorher.Guid; Name = $vorher.Name } }))
}

function Invoke-RegistrySollwerte {
    param([string] $Titel, [object[]] $Werte, [switch] $Benutzer)
    if ($Benutzer -and -not $Script:GleicherBenutzer) { return New-Ergebnis 'übersprungen' 'Benutzer-Einstellung - die Bereinigung läuft unter einem anderen Konto als die App' }
    # Wert fuer Wert: scheitert ein spaeterer Wert, bleiben die Sicherungen der schon geschriebenen erhalten (rueckgaengig moeglich)
    $undo = New-Object System.Collections.Generic.List[object]
    foreach ($w in $Werte) {
        try { $s = Set-RegDwordMitSicherung $w[0] $w[1] $w[2] $w[3] }
        catch { return New-Ergebnis 'fehlgeschlagen' "$Titel - $($w[2]) nicht gesetzt: $($_.Exception.Message)$(if ($undo.Count) { " ($($undo.Count) Wert(e) davor schon geändert - rückgängig möglich)" })" $undo.ToArray() }
        if ($null -ne $s) { $undo.Add($s) }
    }
    $undo = $undo.ToArray()
    if ($undo.Count -eq 0) { return New-Ergebnis 'übersprungen' "$Titel - war bereits eingestellt" }
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' "würde: $Titel" }
    return New-Ergebnis 'erledigt' $Titel $undo
}

function Invoke-SchnellstartAus { param($P) Invoke-RegistrySollwerte 'Schnellstart ausgeschaltet' @(, @('HKLM', 'SYSTEM\CurrentControlSet\Control\Session Manager\Power', 'HiberbootEnabled', 0)) }
function Invoke-AbbildTypKlein { param($P) Invoke-RegistrySollwerte 'Kleine Absturzabbilder (Minidump) eingeschaltet' @(, @('HKLM', 'SYSTEM\CurrentControlSet\Control\CrashControl', 'CrashDumpEnabled', 3)) }
function Invoke-UacEinschalten { param($P) Invoke-RegistrySollwerte 'Benutzerkontensteuerung eingeschaltet - wirksam nach dem nächsten Neustart' @(, @('HKLM', 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System', 'EnableLUA', 1)) }
function Invoke-UacNachfrageAn {
    param($P)
    Invoke-RegistrySollwerte 'Benutzerkontensteuerung fragt wieder nach (Windows-Standardstufe)' @(
        @('HKLM', 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System', 'ConsentPromptBehaviorAdmin', 5),
        @('HKLM', 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System', 'PromptOnSecureDesktop', 1))
}
function Invoke-GameDvrAus {
    param($P)
    Invoke-RegistrySollwerte 'Xbox-Hintergrundaufnahme ausgeschaltet' @(
        @('HKCU', 'Software\Microsoft\Windows\CurrentVersion\GameDVR', 'AppCaptureEnabled', 0),
        @('HKCU', 'System\GameConfigStore', 'GameDVR_Enabled', 0)) -Benutzer
}

function Invoke-DynamicLightingAus {
    param($P)
    $e = Invoke-RegistrySollwerte 'Dynamic Lighting ausgeschaltet' @(
        @('HKCU', 'Software\Microsoft\Windows\CurrentVersion\Lighting', 'AmbientLightingEnabled', 0),
        @('HKCU', 'Software\Microsoft\Windows\CurrentVersion\Lighting', 'ControlledByForegroundApp', 0)) -Benutzer
    if ($e.Status -eq 'erledigt' -and (Get-Service -Name 'logi_lamparray_service' -ErrorAction SilentlyContinue)) {
        try { Restart-Service -Name 'logi_lamparray_service' -Force -ErrorAction Stop; $e.Meldung += ', LampArray-Dienst neu gestartet - bitte G HUB einmal neu starten' } catch { $e.Meldung += ' (LampArray-Dienst ließ sich nicht neu starten)' }
    }
    return $e
}

function Invoke-DefenderSignaturen {
    param($P)
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' 'würde die Virensignaturen aktualisieren' }
    Update-MpSignature -ErrorAction Stop
    return New-Ergebnis 'erledigt' 'Virensignaturen aktualisiert'
}

function Invoke-FirewallEinschalten {
    param($P)
    $profile = @($P.Profile | ForEach-Object { [string]$_ })
    if (-not $profile.Count -or @($profile | Where-Object { $_ -notin @('Domain', 'Private', 'Public') }).Count) { return New-Ergebnis 'abgelehnt' 'Ungültiges Firewall-Profil' }
    $aus = @($profile | Where-Object { -not (Get-NetFirewallProfile -Name $_).Enabled })
    if ($aus.Count -eq 0) { return New-Ergebnis 'übersprungen' 'Die Firewall ist bereits überall an' }
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' ('würde die Firewall einschalten für ' + ($aus -join ', ')) }
    foreach ($n in $aus) { Set-NetFirewallProfile -Name $n -Enabled True }
    return New-Ergebnis 'erledigt' ('Firewall eingeschaltet für ' + ($aus -join ', ')) @([ordered]@{ Typ = 'Firewall'; Profile = $aus })
}

function Invoke-ZeitSynchronisieren {
    param($P)
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' 'würde die Uhrzeit synchronisieren' }
    $dienst = Get-Service -Name 'W32Time'
    if ($dienst.Status -ne 'Running') { Start-Service -Name 'W32Time' }
    $r = Invoke-System32 'w32tm.exe' @('/resync', '/force')
    if ($r.Code -eq 0) { return New-Ergebnis 'erledigt' 'Uhrzeit synchronisiert' }
    return New-Ergebnis 'fehlgeschlagen' "w32tm: $(Format-Kurz $r.Text 140)"
}

function Invoke-DnsCacheLeeren {
    param($P)
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' 'würde den DNS-Zwischenspeicher leeren' }
    Clear-DnsClientCache
    return New-Ergebnis 'erledigt' 'DNS-Zwischenspeicher geleert'
}

function Invoke-DoppeltesSkriptBeenden {
    param($P)
    $skript = [string]$P.Skript
    if ($skript -notmatch '^[A-Za-z]:\\[^"<>|*?\r\n]+\.(ps1|vbs)$') { return New-Ergebnis 'abgelehnt' 'Ungültiger Skriptpfad' }
    $proz = @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop | Where-Object { $_.Name -in @('powershell.exe', 'pwsh.exe', 'wscript.exe', 'cscript.exe') -and $_.CommandLine -and $_.CommandLine.IndexOf($skript, [StringComparison]::OrdinalIgnoreCase) -ge 0 } | Sort-Object CreationDate -Descending)
    if ($proz.Count -le 1) { return New-Ergebnis 'übersprungen' 'Das Skript läuft nicht (mehr) doppelt' }
    $alt = @($proz | Select-Object -Skip 1)
    if ($Script:IstSimulation) { return New-Ergebnis 'simuliert' "würde $($alt.Count) ältere Instanz(en) beenden" }
    foreach ($p in $alt) { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue }
    return New-Ergebnis 'erledigt' "$($alt.Count) ältere Instanz(en) von $(Split-Path -Leaf $skript) beendet, die neueste (PID $($proz[0].ProcessId)) läuft weiter"
}

# Explizite Zuordnung - kein dynamischer Aufruf von Funktionsnamen aus dem Plan
function Invoke-Aktion {
    param($Aktion)
    $p = $Aktion.Parameter
    switch ([string]$Aktion.Id) {
        'TreiberpaketEntfernen'          { return Invoke-TreiberpaketEntfernen $p }
        'TreiberdienstDeaktivieren'      { return Invoke-TreiberdienstDeaktivieren $p }
        'DienstDeaktivieren'             { return Invoke-DienstDeaktivieren $p }
        'DienstStarten'                  { return Invoke-DienstStarten $p }
        'AufgabeDeaktivieren'            { return Invoke-AufgabeDeaktivieren $p }
        'AutostartDeaktivieren'          { return Invoke-AutostartDeaktivieren $p }
        'Reinigen'                       { return Invoke-Reinigen $p }
        'WiederherstellungspunktAnlegen' { return Invoke-WiederherstellungspunktAnlegen $p }
        'DismReparieren'                 { return Invoke-DismReparieren $p }
        'SfcReparieren'                  { return Invoke-SfcReparieren $p }
        'EnergieplanSetzen'              { return Invoke-EnergieplanSetzen $p }
        'SchnellstartAus'                { return Invoke-SchnellstartAus $p }
        'GameDvrAus'                     { return Invoke-GameDvrAus $p }
        'DynamicLightingAus'             { return Invoke-DynamicLightingAus $p }
        'AbbildTypKlein'                 { return Invoke-AbbildTypKlein $p }
        'UacEinschalten'                 { return Invoke-UacEinschalten $p }
        'UacNachfrageAn'                 { return Invoke-UacNachfrageAn $p }
        'DefenderSignaturen'             { return Invoke-DefenderSignaturen $p }
        'FirewallEinschalten'            { return Invoke-FirewallEinschalten $p }
        'ZeitSynchronisieren'            { return Invoke-ZeitSynchronisieren $p }
        'DnsCacheLeeren'                 { return Invoke-DnsCacheLeeren $p }
        'DoppeltesSkriptBeenden'         { return Invoke-DoppeltesSkriptBeenden $p }
        default                          { return New-Ergebnis 'abgelehnt' "Unbekannte Aktion '$($Aktion.Id)' - steht nicht im Katalog" }
    }
}

# ============================================================================
#  Rueckgaengig
# ============================================================================

function Undo-Eintrag {
    param($E)
    switch ([string]$E.Typ) {
        'Registry'    { if ($E.Hive -eq 'HKCU' -and -not $Script:GleicherBenutzer) { throw 'Benutzer-Einstellung eines anderen Kontos' }; Restore-RegWert $E; return "Einstellung $($E.Name) zurückgesetzt" }
        'Start'       {
            # Starttyp und Laufzustand wie vorher; Fehler werden gemeldet statt verschluckt
            $name = [string]$E.Name
            Set-DienstStart $name (ConvertTo-ScStart ([string]$E.Vorher))
            $jetzt = Get-CimInstance -ClassName Win32_Service -Filter "Name='$name'" -ErrorAction SilentlyContinue
            if ($E.LiefVorher -and $jetzt -and $jetzt.State -ne 'Running') {
                try { Start-Service -Name $name -ErrorAction Stop } catch { throw "Starttyp von $name zurückgesetzt, aber der Dienst startet nicht: $($_.Exception.Message)" }
                return "Starttyp von $name auf '$($E.Vorher)' zurückgesetzt und Dienst wieder gestartet"
            }
            if (-not $E.LiefVorher -and $jetzt -and $jetzt.State -eq 'Running') {
                try { Stop-Service -Name $name -Force -ErrorAction Stop } catch { throw "Starttyp von $name zurückgesetzt, aber der Dienst lässt sich nicht stoppen: $($_.Exception.Message)" }
                return "Starttyp von $name auf '$($E.Vorher)' zurückgesetzt und Dienst wieder gestoppt"
            }
            return "Starttyp von $name auf '$($E.Vorher)' zurückgesetzt"
        }
        'Aufgabe'     { [void](Enable-ScheduledTask -TaskPath ([string]$E.TaskPath) -TaskName ([string]$E.TaskName) -ErrorAction Stop); return "Aufgabe $($E.TaskPath)$($E.TaskName) wieder aktiviert" }
        'Energieplan' {
            if ([string]$E.Guid -notmatch '^[0-9a-fA-F\-]{36}$') { throw 'Ungültige Plan-ID' }
            $r = Invoke-System32 'powercfg.exe' @('/setactive', [string]$E.Guid)
            if ($r.Code -ne 0) { throw "powercfg meldet Fehler $($r.Code) - Energieplan '$($E.Name)' nicht aktiviert" }
            return "Energieplan '$($E.Name)' wieder aktiviert"
        }
        'Firewall'    { foreach ($n in @($E.Profile)) { if ($n -in @('Domain', 'Private', 'Public')) { Set-NetFirewallProfile -Name $n -Enabled False } }; return 'Firewall-Profile wie vorher' }
        default       { throw "Unbekannter Rückgängig-Typ '$($E.Typ)'" }
    }
}

# ============================================================================
#  Hauptablauf
# ============================================================================

$ergebnisPfad = $null
$protokoll = New-Object System.Collections.Generic.List[object]

# Gleiche Kennung wie Engine und App (Name|UUID): Ergebnisse und Rueckgaengig gehoeren zu genau diesem PC.
# Ein kopierter Ordner darf auf einem anderen PC keine alten Werte zuruecksetzen (HKLM, Dienste, Firewall).
function Get-MaschinenId {
    $uuid = ''
    try { $uuid = [string](Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop).UUID } catch { }
    if ($uuid -match '^(0{8}|F{8})-') { $uuid = '' }
    return "$env:COMPUTERNAME|$uuid"
}
$maschine = Get-MaschinenId

# Ergebnis samt Rueckgaengig-Daten nach JEDER Aktion schreiben - ein Abbruch mittendrin verliert sonst die Sicherungen
# der schon erledigten Aktionen. -Pflicht: vor der ersten Aenderung; scheitert das Schreiben, wird nichts geaendert.
function Save-Ergebnis {
    param([switch] $Pflicht)
    $kopf['Aktionen'] = $protokoll.ToArray()
    try { [IO.File]::WriteAllText($ergebnisPfad, ($kopf | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($true))) }
    catch {
        if ($Pflicht) { throw "Die Sicherungsdatei lässt sich nicht schreiben ($($_.Exception.Message)) - es wurde nichts verändert." }
        Write-Protokoll "WARNUNG: Ergebnisdatei nicht geschrieben: $($_.Exception.Message)"
    }
}
$kopf = [ordered]@{ Format = 'Systemdiagnose-Ergebnis/1'; Art = ''; Zeit = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Simulation = $Script:IstSimulation; Bezug = ''; Wiederherstellungspunkt = ''; Fehler = ''; Maschine = $maschine }
try {
    if (-not (Test-Path -LiteralPath $Script:Arbeitsordner)) { [void](New-Item -ItemType Directory -Path $Script:Arbeitsordner) }
    $istAdmin = Test-IstAdmin
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value

    if ($Rueckgaengig) {
        $quelle = Get-SichereDatei $Rueckgaengig '^Ergebnis-\d{8}_\d{6}\.json$'
        $ergebnisPfad = Join-Path $Script:Arbeitsordner "Rueckgaengig-$Stempel.json"
        $Script:StatusDatei = Join-Path $Script:Arbeitsordner "Status-$Stempel.txt"
        $kopf.Art = 'Rückgängig'; $kopf.Bezug = Split-Path -Leaf $quelle
        if (-not $istAdmin -and -not $Script:IstSimulation) { throw 'Rückgängig machen braucht Adminrechte.' }
        $alt = Get-Content -LiteralPath $quelle -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($alt.Format -ne 'Systemdiagnose-Ergebnis/1') { throw 'Unbekanntes Ergebnisformat - es wird nichts zurückgenommen.' }
        if ($alt.Maschine -and [string]$alt.Maschine -ne $maschine) { throw 'Diese Behebung stammt von einem anderen PC - Rückgängig würde dessen alte Werte hier einspielen. Es wird nichts verändert.' }
        $Script:GleicherBenutzer = ([string]$alt.BenutzerSid -eq $sid)
        if (-not $Script:IstSimulation) { Save-Ergebnis -Pflicht }
        # Auch fehlgeschlagene Aktionen mit Sicherung: sie haben evtl. schon einen Teil geaendert
        foreach ($a in @($alt.Aktionen | Where-Object { $_.Status -in 'erledigt', 'fehlgeschlagen' -and @($_.Rueckgaengig | Where-Object { $_ }).Count })) {
            Write-Status "Nehme zurück: $($a.Titel)"
            $meldungen = @(); $status = 'erledigt'
            foreach ($e in @($a.Rueckgaengig | Where-Object { $_ })) {
                try { if ($Script:IstSimulation) { $meldungen += "würde zurücknehmen ($($e.Typ))" } else { $meldungen += Undo-Eintrag $e } } catch { $status = 'fehlgeschlagen'; $meldungen += $_.Exception.Message }
            }
            $zeile = [pscustomobject]@{ Id = $a.Id; Titel = $a.Titel; BefundNr = $a.BefundNr; Status = $(if ($Script:IstSimulation) { 'simuliert' } else { $status }); Meldung = ($meldungen -join '; '); Rueckgaengig = @() }
            $protokoll.Add($zeile)
            Write-Protokoll "RÜCKGÄNGIG [$($zeile.Status)] $($a.Titel): $($zeile.Meldung)"
            if (-not $Script:IstSimulation) { Save-Ergebnis }
        }
    } elseif ($Plan) {
        $planPfad = Get-SichereDatei $Plan '^Plan-\d{8}_\d{6}\.json$'
        $planStempel = [regex]::Match((Split-Path -Leaf $planPfad), '\d{8}_\d{6}').Value
        $ergebnisPfad = Join-Path $Script:Arbeitsordner "Ergebnis-$planStempel.json"
        $Script:StatusDatei = Join-Path $Script:Arbeitsordner "Status-$planStempel.txt"
        $kopf.Art = 'Bereinigung'; $kopf.Bezug = Split-Path -Leaf $planPfad
        $planDaten = Get-Content -LiteralPath $planPfad -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($planDaten.Format -ne 'Systemdiagnose-Plan/1') { throw 'Unbekanntes Planformat.' }
        if (-not $istAdmin -and -not $Script:IstSimulation) { throw 'Die Bereinigung braucht Adminrechte.' }
        $Script:GleicherBenutzer = ([string]$planDaten.BenutzerSid -eq $sid)
        $kopf['BenutzerSid'] = $sid
        $aktionen = @($planDaten.Aktionen)
        if ($aktionen.Count -eq 0 -or $aktionen.Count -gt 60) { throw 'Der Plan enthält keine oder zu viele Aktionen.' }

        # Wiederherstellungspunkt vor Aktionen, die eine Systemsicherung brauchen (Treiberpakete)
        $braucht = @($aktionen | Where-Object { $Script:Katalog.ContainsKey([string]$_.Id) -and $Script:Katalog[[string]$_.Id].Sicherung })
        if ($braucht.Count) {
            Write-Status 'Lege einen Wiederherstellungspunkt an ...'
            $wp = Confirm-Wiederherstellungspunkt
            $Script:WiederherstellungOk = $wp.Ok
            $kopf.Wiederherstellungspunkt = $wp.Text
            Write-Protokoll "Wiederherstellungspunkt: $($wp.Text)"
        } else {
            $kopf.Wiederherstellungspunkt = 'nicht nötig - keine Aktion verändert Treiber oder Systemdateien'
        }

        if (-not $Script:IstSimulation) { Save-Ergebnis -Pflicht }
        $i = 0
        foreach ($a in $aktionen) {
            $i++
            Write-Status "($i/$($aktionen.Count)) $($a.Titel)"
            $code = ''
            try { $erg = Invoke-Aktion $a } catch { $erg = New-Ergebnis 'fehlgeschlagen' $_.Exception.Message; $code = Get-Fehlercode $_ }
            $zeile = [pscustomobject]@{ Id = [string]$a.Id; Titel = [string]$a.Titel; BefundNr = $a.BefundNr; Status = $erg.Status; Meldung = $erg.Meldung; Rueckgaengig = @($erg.Rueckgaengig); FreigegebenBytes = $(if ($erg.Bytes -ge 0) { $erg.Bytes } else { $null }); Fehlercode = $code }
            $protokoll.Add($zeile)
            Write-Protokoll "$(if ($Script:IstSimulation) { 'SIMULATION ' })[$($zeile.Status)] $($zeile.Titel): $($zeile.Meldung)"
            if (-not $Script:IstSimulation) { Save-Ergebnis }
        }
    } else {
        throw 'Kein Plan angegeben.'
    }
} catch {
    $kopf.Fehler = $_.Exception.Message
    $kopf.Fehlercode = Get-Fehlercode $_
    Write-Protokoll "FEHLER: $($_.Exception.Message) [$($kopf.Fehlercode)]"
}

if (-not $ergebnisPfad) { $ergebnisPfad = Join-Path $Script:Arbeitsordner ("Ergebnis-" + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.json') }
Save-Ergebnis
if ($Script:StatusDatei) { Remove-Item -LiteralPath $Script:StatusDatei -Force -ErrorAction SilentlyContinue }
Write-Output $ergebnisPfad
