#Requires -Version 5.1
<#
.SYNOPSIS
    Systemdiagnose - Diagnose-Engine. Prueft Windows tiefgehend und aendert dabei NICHTS.
    Erkennt den PC selbst: die eingebauten Sollwerte gelten nur fuer PC 1, auf anderen PCs eingelernte oder keine.

.DESCRIPTION
    Ablauf:
      1. Sicherheits-Selbstpruefung: Das Skript liest seinen eigenen Code und bricht ab,
         sobald darin etwas Schreibendes oder nicht Freigegebenes steht.
      2. Datenquellen ("Sammler"): Jede Quelle wird genau EINMAL gelesen und zwischengespeichert.
      3. Pruefungen: werten nur noch die zwischengespeicherten Daten aus. Jeder Befund erklaert,
         was das ist, woran es liegt und was zu tun ist - und beschreibt (nur als Daten!) moegliche
         Loesungen, die die App ueber den getrennten Bereinigungsteil anbieten kann.
      4. Vergleich mit dem letzten Lauf.
      5. Ausgabe im Ordner "Berichte": HTML-Bericht, Befunddatei (JSON, fuer die App) und
         Vergleichsdatei - der einzige Schreibvorgang.

    Sicherheitsmechanismen (voneinander unabhaengig):
      A. Selbstpruefung per PowerShell-Syntaxbaum (Erlaubt-Liste fuer Befehle, Sperrliste fuer
         schreibende .NET-Methoden, keine Datei-Umleitungen, keine dynamischen Aufrufe).
      B. $WhatIfPreference = $true: Jeder Befehl, der etwas aendern koennte, wuerde nur simulieren.
      C. Registry ausschliesslich ueber schreibgeschuetzte Handles (RegistryRights.ReadKey).
      D. Externe Programme nur aus System32 und nur mit exakt freigegebenen, lesenden Parametern.
      E. Einziger Schreibzugriff: neue Dateien im eigenen Ordner "Berichte", niemals ueberschreiben.
      F. Standardmaessig ohne Adminrechte; Adminmodus nur auf ausdruecklichen Wunsch.

.PARAMETER Modus
    Schnell     = ohne Adminrechte (Standard)
    Gruendlich  = braucht Admin: zusaetzlich SMART, TPM, Wiederherstellungspunkte, DISM (Pruefung, Analyse, /ScanHealth),
                  sfc /verifyonly, Dateisystem (chkdsk im Lesemodus, NTFS-Ereignisse, Reparaturstatus), Signaturen aller
                  Kernel-Treiber und Autostart-Dateien, Ergebnis der Windows-Speicherdiagnose (einige Minuten)
    Tief        = frueherer Name, wird wie Gruendlich behandelt (alte Aufrufe und Berichte)
    Reinigung   = misst nur, wie viel die Reinigungskategorien (Reinigung.psd1) freigeben wuerden
#>
[CmdletBinding()]
param(
    [ValidateSet('Schnell', 'Tief', 'Gruendlich', 'Reinigung')]
    [string] $Modus = 'Schnell',

    [ValidateRange(1, 90)]
    [int] $Tage = 14,

    [switch] $Selbsttest,
    [switch] $KeinBrowser,
    [switch] $Pause,

    # Nur fuer Tests: so tun, als liefe die Engine auf einem anderen PC (Kennung „Name|UUID“)
    [string] $Maschine
)

# Seit 28.10. nur noch zwei Stufen - die fruehere Mittelstufe „Tief“ ist in „Gruendlich“ aufgegangen
if ($Modus -eq 'Tief') { $Modus = 'Gruendlich' }

# Das CIM-Modul legt beim Laden Kurznamen (Aliase) an. Unter dem WhatIf-Netz wuerde das nur
# simuliert und Meldungen erzeugen - deshalb wird dieses eine Microsoft-Modul vorab geladen.
Import-Module -Name CimCmdlets -ErrorAction SilentlyContinue

# Einheitliche Zeichenkodierung fuer die Ausgabe externer Programme (pnputil, powercfg, ...).
# Sonst haengt das Ergebnis davon ab, wie das Skript gestartet wurde (z. B. wird aus "™" ein "T"),
# und der Vergleich mit dem letzten Lauf meldet Aenderungen, die keine sind. Wirkt nur auf dieses Konsolenfenster.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# Sicherheitsnetz B: Jeder Befehl mit Aenderungswirkung wuerde ab hier nur simulieren.
$WhatIfPreference = $true
$ErrorActionPreference = 'Continue'

$Script:Version       = '2.9.4 (02.11.2026)'
$Script:ToolOrdner    = Split-Path -Parent $PSCommandPath
# Datenordner (Berichte, Profile): neben den Programmdateien. Store-Fassung (MSIX): Programmordner unter WindowsApps ist
# schreibgeschuetzt -> LocalState des Pakets; echter Pfad, damit App und erhoehte Kindprozesse dieselben Dateien sehen.
# Gleiche Regel in Systemdiagnose-App.ps1 und Bereinigung.ps1.
$Script:DatenOrdner   = $Script:ToolOrdner
if ($Script:ToolOrdner -match '\\WindowsApps\\([^\\_]+)_[^\\_]*_[^\\_]*_[^\\_]*_([^\\_]+)$') { $Script:DatenOrdner = Join-Path $env:LOCALAPPDATA ('Packages\' + $Matches[1] + '_' + $Matches[2] + '\LocalState') }
$Script:BerichtOrdner = Join-Path $Script:DatenOrdner 'Berichte'
$Script:ReinigungsDatei = Join-Path $Script:ToolOrdner 'Reinigung.psd1'
$Script:Cache         = @{}
$Script:Befunde       = New-Object System.Collections.Generic.List[object]
$Script:Quellen       = New-Object System.Collections.Generic.List[object]
$Script:QuellenFehler = @{}
$Script:IstAdmin      = $false
$Script:AdminModus    = $false
$Script:StartZeit     = Get-Date

# ============================================================================
#  1. PC-PROFIL (Sollwerte) - hier anpassen, wenn sich am PC etwas aendert
# ============================================================================
# Dieses eingebaute Profil gilt NUR auf PC 1 (Rechnername + Mainboard-UUID). Auf jedem anderen PC prueft die Engine
# neutral; die App lernt dort beim ersten Lauf den aktuellen Zustand als Sollwert ein (Profile\<Kennung>.json).
$Script:Pc1Id = 'DESKTOP-620U4TE|17566DA4-794F-1A1F-84B3-B082E2460564'
$Script:MaschinenId  = ''
$Script:IstEigenerPc = $true
$Script:ProfilQuelle = 'eingebaut'
$Script:ProfilStand  = ''
$Script:Profil = [ordered]@{
    Name                  = 'PC 1 - Ryzen 5 3600 / ASUS ROG STRIX B550-F / RX 7800 XT'
    EnergieplanMuster     = 'AMD Ryzen.*High Performance'
    RamMhz                = 3200
    RamModule             = 2
    DiensteMuessenLaufen  = [ordered]@{
        'logi_lamparray_service' = 'Über diesen Dienst läuft das komplette Logitech-RGB (Maus und Headset). Ohne ihn bleibt die Beleuchtung dunkel.'
    }
    DiensteMuessenAusSein = [ordered]@{
        'NahimicService' = 'Nahimic/Sonic Studio ist im September wiederholt abgestürzt und wurde deshalb bewusst abgeschaltet - der Ton läuft ohne ihn einwandfrei.'
    }
    Monitore              = @(
        [ordered]@{ Edid = 'IVM661F'; Name = 'iiyama PL2783Q (Bildschirm 1)';    Aufloesung = '2560x1440'; Hz = 70;  Hdr = $false }
        [ordered]@{ Edid = 'PHL198B'; Name = 'Philips 43PUS8808 (Bildschirm 2)'; Aufloesung = '3840x2160'; Hz = 120; Hdr = $true }
    )
    MinFreiProzent        = 10
    LaufwerkAusnahmen     = [ordered]@{
        'D' = 'USB-Stick'
        'E' = 'Medienlaufwerk (Serien/Filme) - wenig freier Platz ist hier egal'
    }
    FirmwareHinweise      = [ordered]@{
        'EXT0BB0Q' = 'Samsung 840 EVO mit Firmware vor dem Fix für langsames Lesen alter Daten (aktuell wäre EXT0DB6Q). Daten auf dieser SSD gesichert halten.'
    }
    BekannteTiefenTreiber = @('MSIO', 'Asusgio3', 'IOMap', 'CorsairLLAccess*', 'cpuz*', 'AMDRyzenMasterDriver*', 'CTIAIO', 'vgk', 'HWiNFO*')
    UpdateMaxTage         = 45
}

# Bekanntes, harmloses Log-Rauschen auf diesem PC (Quelle + Id, optional Textmuster)
$Script:BekanntesRauschen = @(
    @{ Quelle = 'Microsoft-Windows-DistributedCOM';      Id = 10016; Muster = '';           Grund = 'DCOM-Berechtigungswarnung - Windows-Standardrauschen' }
    @{ Quelle = 'Microsoft-Windows-DistributedCOM';      Id = 10010; Muster = '';           Grund = 'DCOM-Zeitüberschreitung - meist Store-Apps oder bewusst abgeschaltete Dienste (GameDVR)' }
    @{ Quelle = 'Microsoft-Windows-TPM-WMI';             Id = 1796;  Muster = '';           Grund = 'Secure-Boot-SBAT-Update von der Firmware abgelehnt - auf reinem Windows-System bedeutungslos' }
    @{ Quelle = 'Microsoft-Windows-Time-Service';        Id = 158;   Muster = '';           Grund = 'Hyper-V-Zeitanbieter - auf einem echten PC nicht verfügbar' }
    @{ Quelle = 'Microsoft-Windows-WLAN-AutoConfig';     Id = 10002; Muster = 'mtkihvx';    Grund = 'WLAN-Erweiterungsmodul wird beim Herunterfahren beendet' }
    @{ Quelle = 'e2fnexpress';                           Id = 27;    Muster = '';           Grund = 'LAN-Anschluss (Intel I226-V) wird bewusst nicht genutzt' }
    @{ Quelle = 'Microsoft-Windows-WindowsUpdateClient'; Id = 20;    Muster = '0x80073D02'; Grund = 'Store-App war beim Update geöffnet - wird später nachgeholt' }
    @{ Quelle = 'Service Control Manager';               Id = 7000;  Muster = 'cpuz\d+';    Grund = 'CPU-Z registriert bei jedem Start einen neu benannten Temp-Treiber' }
    @{ Quelle = 'Microsoft-Windows-DNS-Client';          Id = 1014;  Muster = '';           Grund = 'einzelne DNS-Zeitüberschreitungen' }
    @{ Quelle = 'HidBth';                                Id = 2;     Muster = '';           Grund = 'Bluetooth-Gerät ausgeschaltet oder außer Reichweite' }
    @{ Quelle = 'BTHUSB';                                Id = 29;    Muster = '';           Grund = 'Bluetooth-Kopplungsversuch eines fremden Geräts' }
    @{ Quelle = '.NET Runtime';                          Id = 1022;  Muster = '';           Grund = '.NET-Profiler konnte nicht geladen werden - harmlos' }
    @{ Quelle = 'Universal Print';                       Id = 1;     Muster = '';           Grund = 'Microsoft-Cloud-Druckdienst findet keinen angemeldeten Benutzer - harmlos' }
    @{ Quelle = 'CodexSandboxService.OpenAI.Codex';      Id = 1003;  Muster = '';           Grund = 'Codex-Sandbox-Dienst (OpenAI) protokolliert beim Starten/Beenden - harmlos' }
    @{ Quelle = 'VSS';                                   Id = 13;    Muster = 'CEventSystem'; Grund = 'Schattenkopie-Dienst beim Herunterfahren - Ereignissystem war schon beendet' }
    @{ Quelle = 'VSS';                                   Id = 8193;  Muster = '0x8007045b'; Grund = 'Schattenkopie-Dienst beim Herunterfahren (0x8007045b = Herunterfahren läuft)' }
    @{ Quelle = 'Microsoft-Windows-Perflib';             Id = 1005;  Muster = 'OpenLsaPerformanceData'; Grund = 'Leistungszähler des LSA-Dienstes fehlt - bekanntes Windows-Rauschen' }
)

# Treiber, die direkt auf Hardware (SMBus, EC, MSR) zugreifen - Name oder Datei
$Script:TiefenTreiberMuster = '(?i)(msio|asio|iomap|llaccess|cpuz|ryzenmaster|ctiaio|winring|inpout|rtcore|gdrv|hwinfo|glckio|atillk|phymem|aida|nvoclock|speedfan|vgk)'

$Script:GeraeteFehlerCodes = @{
    1 = 'nicht richtig konfiguriert'; 3 = 'Treiber beschädigt'; 10 = 'kann nicht gestartet werden'; 12 = 'Ressourcenkonflikt'
    14 = 'Neustart nötig'; 18 = 'Treiber neu installieren'; 19 = 'Registry-Eintrag beschädigt'; 22 = 'vom Benutzer deaktiviert'
    24 = 'nicht vorhanden oder fehlerhaft'; 28 = 'kein Treiber installiert'; 29 = 'von der Firmware deaktiviert'; 31 = 'Treiber lädt nicht'
    32 = 'Treiberdienst deaktiviert'; 37 = 'Treiber-Initialisierung fehlgeschlagen'; 39 = 'Treiber beschädigt oder fehlt'
    41 = 'Treiber geladen, Gerät nicht gefunden'; 43 = 'Gerät hat ein Problem gemeldet'; 48 = 'Treiber blockiert'; 52 = 'Treibersignatur ungültig'
}
$Script:LiveKernelCodes = @{ '141' = 'Grafikkarte hing (Engine-Timeout)'; '117' = 'Grafiktreiber-Timeout (TDR)'; '193' = 'Grafik-Kernel-Livedump'; '124' = 'Hardwarefehler (WHEA)' }
$Script:DienstEreignisse = @{
    7000 = 'Start fehlgeschlagen'; 7001 = 'abhängiger Dienst fehlgeschlagen'; 7009 = 'Start-Zeitüberschreitung'; 7011 = 'Antwort-Zeitüberschreitung'
    7022 = 'hängt beim Start'; 7023 = 'mit Fehler beendet'; 7024 = 'mit dienstspezifischem Fehler beendet'
    7031 = 'unerwartet beendet (wird neu gestartet)'; 7034 = 'unerwartet beendet'; 7043 = 'nicht ordnungsgemäß heruntergefahren'
}

# "Was ist das?" - eine Erklaerung je Befundtyp
$Script:Wissen = @{
    'Absturz'           = 'Ein Neustart ohne ordentliches Herunterfahren. Windows trägt ihn beim nächsten Start als „Kernel-Power 41“ ein - egal ob der PC eingefroren war, der Strom weg war oder der Einschaltknopf gedrückt gehalten wurde.'
    'Bluescreen'        = 'Ein kontrollierter Notstopp von Windows, wenn ein Treiber oder die Hardware einen schweren Fehler verursacht. Dabei entsteht ein Absturzabbild, mit dem sich der Verursacher oft eingrenzen lässt.'
    'WHEA'              = 'Das Hardware-Fehlerprotokoll von Windows: Prozessor, Arbeitsspeicher und PCIe-Geräte melden hier Fehler. Korrigierte Fehler sind Vorwarnungen, schwere Fehler führen fast immer zum Absturz.'
    'WheaStart'         = 'Sucht ein bestimmtes Muster: PC ausgeschaltet, kurz darauf wieder eingeschaltet und in den ersten Minuten ein schwerer Hardwarefehler. Auf diesem PC ist das am 19.09. und 25.10. passiert - es deutet auf die Stromversorgung beim Kaltstart, nicht auf einen Defekt im laufenden Betrieb.'
    'Chkdsk'            = 'Prüft die Verwaltungsdaten des Dateisystems jedes Laufwerks (Ordnerstruktur, Belegung) - im Lesemodus, es wird nichts repariert.'
    'LiveKernel'        = 'Hängt die Grafikkarte oder ein Kerneltreiber kurz, setzt Windows ihn zurück, ohne dass der PC abstürzt (das Bild wird kurz schwarz), und legt einen Livedump an.'
    'Zuverlaessigkeit'  = 'Windows bewertet die Stabilität stündlich von 1 (schlecht) bis 10 (perfekt). Jeder Absturz oder Fehler senkt den Wert, danach steigt er nur langsam wieder an.'
    'AppCrash'          = 'Programme, die mit einem Fehler beendet wurden. Das „Modul“ ist die Datei, in der der Fehler passiert ist - oft ein Treiber oder eine Spieldatei.'
    'AppHang'           = 'Programme, die nicht mehr reagiert haben und beendet wurden.'
    'Dienstfehler'      = 'Windows-Dienste sind Hintergrundprogramme. Hier stehen Dienste, die nicht starten konnten oder unerwartet beendet wurden.'
    'BootTreiber'       = 'Treiber, die schon beim Start von Windows geladen werden sollen.'
    'Dumps'             = 'Speicherauszüge, die bei Abstürzen und Hängern entstehen. Wertvoll für die Fehlersuche, aber teils mehrere GB groß.'
    'WerQueue'          = 'Fehlerberichte, die Windows an Microsoft senden möchte. Bleiben sie hängen, versucht Windows es immer wieder.'
    'Ereignisse'        = 'Alle übrigen Fehlermeldungen aus dem Ereignisprotokoll, die keiner der speziellen Prüfungen zugeordnet sind.'
    'Rauschen'          = 'Meldungen, die auf diesem PC bekannt und nachweislich harmlos sind. Sie werden gezählt, aber nicht als Problem gewertet.'
    'Starts'            = 'Wie oft Windows im Zeitraum gestartet wurde.'
    'Geraetefehler'     = 'Geräte, die im Geräte-Manager mit einem Warnsymbol erscheinen würden - zum Beispiel ohne Treiber oder mit Startfehler.'
    'Geister'           = 'Geräte, die früher angeschlossen waren (Handy, Headset, USB-Sticks, Bluetooth). Windows merkt sich ihre Treiberzuordnung für den nächsten Anschluss.'
    'Treiberspeicher'   = 'Der Treiberspeicher enthält alle installierten Treiberpakete. Bei Treiber-Updates bleiben die alten Pakete oft liegen.'
    'Treiberversionen'  = 'Die Versionen der wichtigsten Treiber, gelesen aus der Registry - verlässlicher als die Windows-Geräteliste.'
    'Grafikkarte'       = 'Der Status, den der Grafiktreiber an Windows meldet.'
    'RAM'               = 'Die erkannten Arbeitsspeicher-Riegel und ihr Takt. Mit aktivem DOCP/XMP läuft der RAM mit voller Geschwindigkeit, sonst deutlich langsamer.'
    'Tiefentreiber'     = 'Kernel-Treiber, über die RGB-, Lüfter-, Sensor- und Übertaktungsprogramme direkt auf Bausteine des Mainboards (SMBus, Embedded Controller) zugreifen.'
    'Fremdtreiber'      = 'Alle laufenden Kernel-Treiber, die nicht von Microsoft stammen.'
    'VerwaisteTreiber'  = 'Treiber, die automatisch starten sollen, deren Datei aber nicht mehr existiert - Reste deinstallierter Software.'
    'Laufwerk'          = 'Freier Speicherplatz und Zustand des Laufwerks.'
    'Datentraeger'      = 'Die SSD bzw. Festplatte selbst: Zustand, Firmware und im Gründlich-Modus die SMART-Werte und Fehlerzähler.'
    'Dateisystem'       = 'Windows markiert ein Laufwerk als „fehlerhaft“, wenn das Dateisystem beschädigt sein könnte. Dann sollte chkdsk laufen.'
    'Aufraeumen'        = 'Ordner, in denen sich überflüssige Dateien sammeln: Temp-Dateien, Update-Reste, alte Protokolle, Browser- und Programm-Caches. Welche Ordner dazugehören, steht fest in Reinigung.psd1.'
    'Wiederherstellung' = 'Wiederherstellungspunkte sichern Systemdateien, Treiber und Registry, damit Windows auf einen früheren Stand zurückgesetzt werden kann.'
    'Schattenkopien'    = 'Der Speicherbereich, in dem Windows die Wiederherstellungspunkte ablegt.'
    'DISM'              = 'DISM prüft den Komponentenspeicher - die Vorlage, aus der Windows beschädigte Systemdateien repariert.'
    'SFC'               = 'SFC prüft alle geschützten Windows-Systemdateien auf Beschädigungen und Veränderungen.'
    'Signaturen'        = 'Jede Treiber- und Autostart-Datei sollte digital signiert sein. Die Signatur belegt, von wem die Datei stammt und dass sie seitdem nicht verändert wurde.'
    'Speicherdiagnose'  = 'Die Windows-Speicherdiagnose (mdsched) testet den Arbeitsspeicher beim Neustart und schreibt ihr Ergebnis ins Ereignisprotokoll.'
    'Startkonfiguration' = 'Die Startkonfiguration (BCD) legt fest, wie Windows startet - mit welchen Schutzmechanismen, ob im Test- oder Debugmodus.'
    'WMI'               = 'WMI ist die Datenbank, über die Windows und Programme Hardware- und Systemzustände abfragen (auch diese Diagnose).'
    'Trim'              = 'TRIM teilt einer SSD mit, welche Bereiche frei sind. Ohne TRIM wird sie langsamer und verschleißt schneller.'
    'Speicherbelegung'  = 'Wo der Platz auf dem Laufwerk steckt: die größten Ordner bis zur zweiten Ebene, aus einer kompletten Begehung des Laufwerks.'
    'Treiberpruefung'   = 'Die Treiberüberprüfung (Driver Verifier) ist ein Entwicklerwerkzeug, das Treiber absichtlich unter Extrembedingungen testet.'
    'Bedrohungen'       = 'Alles, was Microsoft Defender gefunden hat, und ob es entfernt werden konnte.'
    'Hosts'             = 'Die hosts-Datei leitet Internetadressen fest um - noch bevor der DNS-Server gefragt wird.'
    'Autostart'         = 'Programme, die bei der Anmeldung automatisch mitstarten.'
    'Aufgaben'          = 'Geplante Aufgaben von Drittanbietern, zum Beispiel Updater oder Hilfsprogramme von ASUS, AMD oder Mozilla.'
    'ProfilDienst'      = 'Ein Dienst, dessen Sollzustand für diesen PC festgelegt ist (PC-Profil).'
    'VerwaisteDienste'  = 'Dienste, die auf ein Programm zeigen, das nicht mehr existiert.'
    'Skripte'           = 'PowerShell- und VBS-Skripte, die im Hintergrund laufen - etwa der Blitz-Watcher oder das FPS-Overlay.'
    'Programme'         = 'Alle installierten Programme laut „Apps & Features“.'
    'Energieplan'       = 'Der aktive Energiesparplan bestimmt, wie schnell der Prozessor hoch- und heruntertaktet.'
    'Schnellstart'      = 'Schnellstart fährt Windows beim Herunterfahren nur halb herunter. Der nächste Start ist schneller, aber Treiber werden nicht frisch geladen.'
    'Ruhezustand'       = 'Der Ruhezustand speichert den Arbeitsspeicher auf die SSD (hiberfil.sys) und schaltet den PC dann aus.'
    'GameDVR'           = 'Die Xbox-Hintergrundaufnahme zeichnet laufend die letzten Spielminuten auf.'
    'DynamicLighting'   = 'Eine Windows-Funktion, die RGB-Geräte zentral steuert. Auf diesem PC streitet sie sich mit G HUB um das Logitech-RGB.'
    'Spielmodus'        = 'Der Spielmodus priorisiert Spiele und verschiebt Hintergrundarbeiten von Windows.'
    'HAGS'              = 'Bei der hardwarebeschleunigten GPU-Planung verwaltet die Grafikkarte ihren Arbeitsspeicher selbst.'
    'DirectX'           = 'Globale DirectX-Einstellungen: Optimierungen für Spiele im Fenstermodus und Auto HDR.'
    'DumpTyp'           = 'Legt fest, welche Art Absturzabbild Windows bei einem Bluescreen schreibt.'
    'Auslagerung'       = 'Die Auslagerungsdatei erweitert den Arbeitsspeicher auf die SSD.'
    'Monitor'           = 'Auflösung, Bildwiederholrate und HDR eines Monitors im Vergleich zum Sollwert im PC-Profil.'
    'ClipCursor'        = 'Spiele können den Mauszeiger auf einen Bildschirmbereich einsperren. Beenden sie sich unsauber, bleibt die Sperre manchmal bestehen.'
    'SecureBoot'        = 'Secure Boot lässt beim Start nur signierte Bootloader zu - Pflicht für Windows 11 und für manche Anti-Cheat-Systeme.'
    'UAC'               = 'Die Benutzerkontensteuerung fragt nach, bevor Programme Adminrechte bekommen.'
    'Sperrliste'        = 'Microsofts Liste bekannter anfälliger Treiber, die Windows nicht mehr laden lässt.'
    'TPM'               = 'Das TPM ist ein Sicherheitschip (bei AMD als fTPM in der CPU). Windows 11 setzt ihn voraus.'
    'VBS'               = 'Virtualisierungsbasierte Sicherheit isoliert Teile von Windows in einer geschützten Umgebung. Die Speicherintegrität prüft damit Treiber.'
    'Defender'          = 'Der eingebaute Virenschutz von Windows.'
    'Firewall'          = 'Die Windows-Firewall filtert eingehende Verbindungen.'
    'Updates'           = 'Zuletzt installierte Windows-, Treiber- und Office-Updates.'
    'UpdateFehler'      = 'Updates, deren Installation fehlgeschlagen ist.'
    'Aktivierung'       = 'Ob Windows mit einer gültigen Lizenz aktiviert ist.'
    'Windows'           = 'Installierte Windows-Version und -Build.'
    'Laufzeit'          = 'Zeit seit dem letzten Neustart.'
    'Neustart'          = 'Ob Updates oder Installationen auf einen Neustart warten.'
    'Insider'           = 'Ob Windows Vorabversionen aus dem Insider-Programm bezieht.'
    'BIOS'              = 'Die Firmware des Mainboards.'
    'Prozessor'         = 'Prozessor und nutzbarer Arbeitsspeicher.'
    'Zeit'              = 'Ob die Uhr regelmäßig mit einem Zeitserver abgeglichen wird.'
    'Netzwerk'          = 'Die physischen Netzwerkadapter und ihr Verbindungsstatus.'
    'WLAN'              = 'Signalstärke und Geschwindigkeit der WLAN-Verbindung.'
    'DNS'               = 'Ob Internet-Adressen in IP-Adressen übersetzt werden können.'
    'Router'            = 'Ob der Router (Standardgateway) antwortet.'
}

# ============================================================================
#  2. SICHERHEITS-FREIGABELISTEN (Mechanismus A und D)
# ============================================================================
$Script:ErlaubteBefehle = @(
    'Get-CimInstance', 'Get-WinEvent', 'Get-AuthenticodeSignature', 'Get-MpThreatDetection', 'Get-MpThreat', 'Get-ChildItem', 'Get-Item', 'Get-Content', 'Test-Path', 'Join-Path', 'Split-Path',
    'Get-ScheduledTask', 'Get-ScheduledTaskInfo', 'Get-PnpDevice', 'Get-PhysicalDisk', 'Get-StorageReliabilityCounter',
    'Get-Volume', 'Get-HotFix', 'Get-MpComputerStatus', 'Get-NetFirewallProfile', 'Get-NetAdapter', 'Get-NetIPConfiguration',
    'Resolve-DnsName', 'Test-Connection', 'Get-Tpm', 'Get-ComputerRestorePoint', 'Get-Date',
    'Import-PowerShellDataFile', 'Get-DeliveryOptimizationStatus', 'Get-DeliveryOptimizationPerfSnap',
    'Select-Object', 'Where-Object', 'ForEach-Object', 'Sort-Object', 'Group-Object', 'Measure-Object', 'Out-String',
    'ConvertTo-Json', 'ConvertFrom-Json', 'Write-Host', 'Write-Progress', 'Read-Host',
    'New-Object', 'Add-Type', 'Invoke-Item', 'Import-Module'
)
# Diese Befehle sind nur innerhalb genau einer Funktion zulaessig
$Script:BefehlNurIn = @{ 'Add-Type' = 'Initialize-NativeHelper'; 'Invoke-Item' = 'Open-Report' }
$Script:ErlaubteTypen = @('System.Collections.Generic.List[object]', 'System.Collections.Generic.List[string]', 'System.Text.StringBuilder', 'System.Text.UTF8Encoding')

# Methoden, die etwas veraendern koennen (exakte Namen und Namensanfaenge)
$Script:VerboteneMethoden = @('Start', 'Stop', 'Pause', 'Continue', 'ExecuteCommand', 'Invoke', 'InvokeMethod', 'InvokeMember',
    'InvokeReturnAsIs', 'Put', 'Open', 'OpenWrite', 'OpenRemoteBaseKey', 'Flush', 'Save', 'CloseMainWindow', 'Replace',
    'Exit', 'Shutdown', 'Reset', 'AppendAllText', 'AppendAllLines', 'AppendText', 'Refresh')
$Script:VerbotenePraefixe = '^(Set|Delete|Create|Write|Move|Copy|Kill|Encrypt|Decrypt|Install|Uninstall|Register|Unregister|Remove)'
# Einzige Ausnahmen: der Berichtsspeicher darf seinen Ordner anlegen und seine Datei schreiben
$Script:MethodenAusnahmen = @{ 'WriteAllText' = 'Save-ReportFile'; 'CreateDirectory' = 'Save-ReportFile' }

# Windows-Funktionen, die im C#-Hilfscode nicht vorkommen duerfen
$Script:VerboteneNativeApis = '\b(SetDisplayConfig|ChangeDisplaySettings\w*|DisplayConfigSetDeviceInfo|SetWindowPos|SetWindowLong\w*|ClipCursor|SetCursorPos|SystemParametersInfo\w*|Reg(Set|Delete|Create)\w*|WriteFile|CreateFile\w*|DeleteFile\w*|MoveFile\w*|CopyFile\w*|CreateProcess\w*|ShellExecute\w*|WinExec|SendMessage\w*|PostMessage\w*|SendInput|keybd_event|mouse_event)\b'

# Externe Programme: nur aus System32 und nur mit diesen (lesenden) Parametern
$Script:ProgrammFreigaben = @{
    'pnputil.exe'  = @('^/enum-drivers$')
    'powercfg.exe' = @('^/getactivescheme$')
    'w32tm.exe'    = @('^/query /status$')
    'netsh.exe'    = @('^wlan show interfaces$')
    'dism.exe'     = @('^/online /cleanup-image /checkhealth$', '^/online /cleanup-image /analyzecomponentstore$', '^/online /cleanup-image /scanhealth$')
    'sfc.exe'      = @('^/verifyonly$')
    # chkdsk ohne /f /r /x /scan /spotfix = reiner Lesemodus
    'chkdsk.exe'   = @('^[A-Z]:$')
    'vssadmin.exe' = @('^list shadowstorage$')
    'fsutil.exe'   = @('^dirty query [A-Z]:$', '^repair state [A-Z]:$', '^behavior query DisableDeleteNotify$')
    'bcdedit.exe'  = @('^/enum \{current\}$')
    # WMI-Datenbank nur pruefen (/verifyrepository), nie /salvagerepository oder /resetrepository
    'wbem\winmgmt.exe' = @('^/verifyrepository$')
}

$Script:NurAdmin = @('Smart', 'DirtyBits', 'Tpm', 'RestorePoints', 'ShadowStorage', 'DismHealth', 'DismAnalyse', 'SfcVerify', 'DismScan', 'Chkdsk', 'FsReparatur', 'Trim', 'Bootkonfiguration', 'WmiDatenbank')
# Gruendlich = alles, was Windows ohne Schreibzugriff hergibt, auch wenn es Minuten kostet
$Script:NurGruendlich = @('SfcVerify', 'DismScan', 'Chkdsk', 'Activation', 'Signaturen', 'Speicherdiagnose', 'NtfsEreignisse', 'FsReparatur', 'Trim', 'Bootkonfiguration', 'WmiDatenbank', 'Speicherbelegung')

$Script:BereichReihenfolge = @('Stabilitaet', 'Hardware & Treiber', 'Laufwerke', 'Software & Autostart', 'Einstellungen (PC-Profil)', 'Sicherheit & Updates', 'System', 'Netzwerk')
$Script:BereichTitel = @{
    'Stabilitaet' = 'Stabilität'; 'Hardware & Treiber' = 'Hardware & Treiber'; 'Laufwerke' = 'Laufwerke'; 'Software & Autostart' = 'Software & Autostart'
    'Einstellungen (PC-Profil)' = 'Einstellungen (PC-Profil)'; 'Sicherheit & Updates' = 'Sicherheit & Updates'; 'System' = 'System'; 'Netzwerk' = 'Netzwerk'
}

$Script:QuellenNamen = @{
    'OS' = 'Betriebssystem'; 'WinVersion' = 'Windows-Version (Registry)'; 'ComputerSystem' = 'Computer'; 'BIOS' = 'BIOS'
    'CPU' = 'Prozessor'; 'RAM' = 'Arbeitsspeicher'; 'VideoController' = 'Grafikkarte'; 'PnpDevices' = 'Geräte (inkl. Geistergeräte)'
    'SignedDrivers' = 'Gebundene Treiber'; 'DriverStore' = 'Treiberspeicher (pnputil)'; 'DriverClasses' = 'Treiberversionen (Registry)'
    'Services' = 'Dienste'; 'SystemDrivers' = 'Kernel-Treiber'; 'Processes' = 'Prozesse'
    'EventsSystem' = 'Ereignisprotokoll System'; 'EventsApplication' = 'Ereignisprotokoll Anwendung'; 'Reliability' = 'Zuverlässigkeitsverlauf'; 'ZuverlEreignisse' = 'Zuverlässigkeitsereignisse (Ausfälle)'
    'Volumes' = 'Laufwerke (Volumes)'; 'PhysicalDisks' = 'Datenträger'; 'Smart' = 'SMART-Werte'; 'DirtyBits' = 'Dateisystem-Status'
    'Uninstall' = 'Installierte Programme'; 'Autostart' = 'Autostart'; 'Tasks' = 'Geplante Aufgaben'; 'Display' = 'Monitore / HDR'
    'CursorClip' = 'Mauszeiger-Bereich'; 'NetAdapters' = 'Netzwerkadapter'; 'Wlan' = 'WLAN'; 'IpConfig' = 'IP-Konfiguration'
    'DnsTest' = 'DNS-Test'; 'GatewayTest' = 'Router-Erreichbarkeit'; 'Defender' = 'Microsoft Defender'; 'Firewall' = 'Firewall'
    'HotFix' = 'Installierte Updates'; 'PowerScheme' = 'Energieplan'; 'TimeStatus' = 'Zeitdienst'; 'Dumps' = 'Absturzabbilder'
    'WerFolders' = 'Fehlerberichte (WER)'; 'Reinigung' = 'Reinigungskategorien'; 'Settings' = 'Einstellungen (Registry)'
    'DeviceGuard' = 'Virtualisierungsschutz'; 'PageFile' = 'Auslagerungsdatei'; 'PendingReboot' = 'Ausstehender Neustart'
    'Insider' = 'Insider-Programm'; 'Activation' = 'Aktivierung'; 'Tpm' = 'TPM'; 'RestorePoints' = 'Wiederherstellungspunkte'
    'ShadowStorage' = 'Schattenkopien'; 'DismHealth' = 'Komponentenspeicher (DISM)'; 'DismAnalyse' = 'Komponentenspeicher-Größe (DISM)'; 'SfcVerify' = 'Systemdateien (SFC)'; 'DismScan' = 'Komponentenspeicher-Tiefenscan (DISM)'; 'Chkdsk' = 'Dateisysteme (chkdsk, nur lesen)'
    'Trim' = 'TRIM für SSDs'; 'Bootkonfiguration' = 'Startkonfiguration (bcdedit)'; 'WmiDatenbank' = 'WMI-Datenbank'; 'Speicherbelegung' = 'Speicherbelegung (größte Ordner)'
    'Treiberpruefung' = 'Treiberüberprüfung (Driver Verifier)'; 'Bedrohungen' = 'Defender-Funde'; 'Hosts' = 'hosts-Datei'
    'Signaturen' = 'Signaturen (Treiber, Dienste, Aufgaben, Autostart)'; 'Speicherdiagnose' = 'Windows-Speicherdiagnose'; 'NtfsEreignisse' = 'Dateisystem-Ereignisse (NTFS)'; 'FsReparatur' = 'Dateisystem-Reparaturstatus'
}

# C#-Hilfscode: liest Monitore, HDR-Status und den Mauszeiger-Bereich (nur lesende Windows-Funktionen)
$Script:NativeQuelle = @'
// SD-NATIVE-BEGIN
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class SdMonitor {
    public string Gdi; public string Edid; public string Name;
    public int X; public int Y; public int Width; public int Height;
    public double Hz; public bool Primary;
    public bool HdrSupported; public bool HdrActive; public string Encoding; public int Bits;
}

public static class SdNative {
    [StructLayout(LayoutKind.Sequential)] public struct LUID { public uint Low; public int High; }
    [StructLayout(LayoutKind.Sequential)] public struct RATIONAL { public uint Num; public uint Den; }
    [StructLayout(LayoutKind.Sequential)] public struct SRC { public LUID Adapter; public uint Id; public uint ModeIdx; public uint Status; }
    [StructLayout(LayoutKind.Sequential)] public struct TGT { public LUID Adapter; public uint Id; public uint ModeIdx; public uint Tech; public uint Rotation; public uint Scaling; public RATIONAL Refresh; public uint ScanLine; public int Available; public uint Status; }
    [StructLayout(LayoutKind.Sequential)] public struct PATH { public SRC Source; public TGT Target; public uint Flags; }
    [StructLayout(LayoutKind.Explicit, Size = 64)] public struct MODE {
        [FieldOffset(0)] public uint InfoType; [FieldOffset(4)] public uint Id; [FieldOffset(8)] public LUID Adapter;
        [FieldOffset(16)] public uint Width; [FieldOffset(20)] public uint Height; [FieldOffset(24)] public uint PixFmt;
        [FieldOffset(28)] public int PosX; [FieldOffset(32)] public int PosY;
    }
    [StructLayout(LayoutKind.Sequential)] public struct HEADER { public uint Type; public uint Size; public LUID Adapter; public uint Id; }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] public struct TNAME {
        public HEADER Header; public uint Flags; public uint Tech; public ushort Mfg; public ushort Product; public uint Connector;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)] public string Friendly;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] public string Path;
    }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] public struct SNAME {
        public HEADER Header; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string Gdi;
    }
    [StructLayout(LayoutKind.Sequential)] public struct COLOR { public HEADER Header; public uint Value; public uint Encoding; public uint Bits; }
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [DllImport("user32.dll")] static extern int GetDisplayConfigBufferSizes(uint flags, out uint numPaths, out uint numModes);
    [DllImport("user32.dll")] static extern int QueryDisplayConfig(uint flags, ref uint numPaths, [Out] PATH[] paths, ref uint numModes, [Out] MODE[] modes, IntPtr topology);
    [DllImport("user32.dll")] static extern int DisplayConfigGetDeviceInfo(ref TNAME info);
    [DllImport("user32.dll")] static extern int DisplayConfigGetDeviceInfo(ref SNAME info);
    [DllImport("user32.dll")] static extern int DisplayConfigGetDeviceInfo(ref COLOR info);
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] static extern bool GetClipCursor(out RECT rect);
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] public struct DEVMODE {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string DeviceName;
        public ushort SpecVersion; public ushort DriverVersion; public ushort Size; public ushort DriverExtra; public uint Fields;
        public int PositionX; public int PositionY; public uint DisplayOrientation; public uint DisplayFixedOutput;
        public short Color; public short Duplex; public short YResolution; public short TTOption; public short Collate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string FormName;
        public ushort LogPixels; public uint BitsPerPel; public uint PelsWidth; public uint PelsHeight; public uint DisplayFlags; public uint DisplayFrequency;
        public uint ICMMethod; public uint ICMIntent; public uint MediaType; public uint DitherType; public uint Reserved1; public uint Reserved2; public uint PanningWidth; public uint PanningHeight;
    }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern bool EnumDisplaySettings(string device, int modeNum, ref DEVMODE mode);

    // Alle Modi, die der Monitor gerade anbietet ("BxH@Hz") - nur lesen, nichts umstellen
    public static List<string> Modes(string gdi) {
        var list = new List<string>();
        var dm = new DEVMODE(); dm.Size = (ushort)Marshal.SizeOf(typeof(DEVMODE));
        for (int i = 0; EnumDisplaySettings(gdi, i, ref dm); i++) {
            string s = dm.PelsWidth + "x" + dm.PelsHeight + "@" + dm.DisplayFrequency;
            if (!list.Contains(s)) list.Add(s);
        }
        return list;
    }

    public static List<SdMonitor> Monitors() {
        var list = new List<SdMonitor>();
        IntPtr old = SetThreadDpiAwarenessContext(new IntPtr(-4));
        try {
            uint np, nm;
            if (GetDisplayConfigBufferSizes(2, out np, out nm) != 0) return list;
            var paths = new PATH[np];
            var modes = new MODE[nm];
            if (QueryDisplayConfig(2, ref np, paths, ref nm, modes, IntPtr.Zero) != 0) return list;
            for (int i = 0; i < np; i++) {
                var p = paths[i];
                var m = new SdMonitor();
                var tn = new TNAME();
                tn.Header.Type = 2; tn.Header.Size = (uint)Marshal.SizeOf(typeof(TNAME)); tn.Header.Adapter = p.Target.Adapter; tn.Header.Id = p.Target.Id;
                if (DisplayConfigGetDeviceInfo(ref tn) == 0) {
                    m.Name = tn.Friendly;
                    string[] parts = (tn.Path ?? "").Split('#');
                    m.Edid = parts.Length > 1 ? parts[1] : "";
                }
                var sn = new SNAME();
                sn.Header.Type = 1; sn.Header.Size = (uint)Marshal.SizeOf(typeof(SNAME)); sn.Header.Adapter = p.Source.Adapter; sn.Header.Id = p.Source.Id;
                if (DisplayConfigGetDeviceInfo(ref sn) == 0) m.Gdi = sn.Gdi;
                var c = new COLOR();
                c.Header.Type = 9; c.Header.Size = (uint)Marshal.SizeOf(typeof(COLOR)); c.Header.Adapter = p.Target.Adapter; c.Header.Id = p.Target.Id;
                if (DisplayConfigGetDeviceInfo(ref c) == 0) {
                    m.HdrSupported = (c.Value & 1) != 0;
                    m.HdrActive = (c.Value & 2) != 0;
                    m.Bits = (int)c.Bits;
                    m.Encoding = c.Encoding == 0 ? "RGB" : c.Encoding == 1 ? "YCbCr444" : c.Encoding == 2 ? "YCbCr422" : c.Encoding == 3 ? "YCbCr420" : "Sonstige";
                }
                if (p.Target.Refresh.Den != 0) m.Hz = Math.Round((double)p.Target.Refresh.Num / p.Target.Refresh.Den, 2);
                uint mi = p.Source.ModeIdx;
                if (mi < nm && modes[mi].InfoType == 1) {
                    m.X = modes[mi].PosX; m.Y = modes[mi].PosY;
                    m.Width = (int)modes[mi].Width; m.Height = (int)modes[mi].Height;
                    m.Primary = (m.X == 0 && m.Y == 0);
                }
                list.Add(m);
            }
        } finally {
            SetThreadDpiAwarenessContext(old);
        }
        return list;
    }

    public static int[] CursorClip() {
        IntPtr old = SetThreadDpiAwarenessContext(new IntPtr(-4));
        try {
            RECT r;
            GetClipCursor(out r);
            return new int[] { r.Left, r.Top, r.Right, r.Bottom, GetSystemMetrics(76), GetSystemMetrics(77), GetSystemMetrics(78), GetSystemMetrics(79) };
        } finally {
            SetThreadDpiAwarenessContext(old);
        }
    }
}

// Groesse jedes Ordners der Ebenen 1..Tiefe unter einer Wurzel - nur Lesen (Eintraege aufzaehlen, Dateilaengen).
// Abzweigungen (Junctions/Symlinks) werden nie verfolgt; nicht lesbare Ordner zaehlen mit dem, was lesbar war.
public class SdOrdner { public string Pfad; public int Ebene; public long Bytes; public long Dateien; }
public static class SdBelegung {
    public static List<SdOrdner> Messen(string wurzel, int tiefe) {
        var liste = new List<SdOrdner>();
        Gehe(new System.IO.DirectoryInfo(wurzel), 0, tiefe, liste);
        return liste;
    }
    static long[] Gehe(System.IO.DirectoryInfo d, int ebene, int tiefe, List<SdOrdner> liste) {
        long bytes = 0, anzahl = 0;
        try {
            foreach (var e in d.EnumerateFileSystemInfos()) {
                if ((e.Attributes & System.IO.FileAttributes.ReparsePoint) != 0) continue;
                var u = e as System.IO.DirectoryInfo;
                if (u != null) { var r = Gehe(u, ebene + 1, tiefe, liste); bytes += r[0]; anzahl += r[1]; }
                else { bytes += ((System.IO.FileInfo)e).Length; anzahl++; }
            }
        } catch (Exception) { }
        if (ebene >= 1 && ebene <= tiefe) liste.Add(new SdOrdner { Pfad = d.FullName, Ebene = ebene, Bytes = bytes, Dateien = anzahl });
        return new long[] { bytes, anzahl };
    }
}
// SD-NATIVE-END
'@

# ============================================================================
#  3. SICHERHEITSFUNKTIONEN
# ============================================================================

function Get-EnclosingFunctionName {
    param($Knoten)
    $p = $Knoten.Parent
    while ($null -ne $p) {
        if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $p.Name }
        $p = $p.Parent
    }
    return ''
}

# Mechanismus A: Das Skript prueft seinen eigenen Quelltext, bevor irgendetwas gelesen wird.
function Invoke-SelfAudit {
    $tokens = $null
    $syntaxFehler = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$tokens, [ref]$syntaxFehler)
    $verstoesse = New-Object System.Collections.Generic.List[string]
    foreach ($f in @($syntaxFehler)) {
        if ($f) { $verstoesse.Add("Syntaxfehler in Zeile $($f.Extent.StartLineNumber): $($f.Message)") }
    }

    $eigene = @{}
    foreach ($fd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $eigene[$fd.Name] = $true }

    $nBefehle = 0
    foreach ($c in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $nBefehle++
        $zeile = $c.Extent.StartLineNumber
        $fn = Get-EnclosingFunctionName $c
        $name = $c.GetCommandName()
        if ([string]::IsNullOrEmpty($name)) {
            if ($fn -ne 'Invoke-SafeTool') { $verstoesse.Add("Zeile ${zeile}: dynamischer Programmaufruf ausserhalb von Invoke-SafeTool") }
            continue
        }
        if ($eigene.ContainsKey($name)) { continue }
        if ($Script:ErlaubteBefehle -notcontains $name) {
            $verstoesse.Add("Zeile ${zeile}: '$name' steht nicht auf der Erlaubt-Liste")
            continue
        }
        $nurIn = $Script:BefehlNurIn[$name]
        if ($nurIn -and $fn -ne $nurIn) { $verstoesse.Add("Zeile ${zeile}: '$name' ist nur in $nurIn erlaubt") }
        if ($name -eq 'New-Object') {
            if ($c.Extent.Text -match '(?i)-Com') { $verstoesse.Add("Zeile ${zeile}: COM-Objekte sind verboten") }
            $typ = $c.CommandElements | Select-Object -Skip 1 |
                Where-Object { $_ -is [System.Management.Automation.Language.StringConstantExpressionAst] } | Select-Object -First 1
            if (-not $typ -or $Script:ErlaubteTypen -notcontains $typ.Value) { $verstoesse.Add("Zeile ${zeile}: New-Object mit nicht freigegebenem Typ") }
        }
    }

    $nMethoden = 0
    foreach ($m in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true)) {
        $nMethoden++
        $zeile = $m.Extent.StartLineNumber
        if ($m.Member -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) {
            $verstoesse.Add("Zeile ${zeile}: dynamischer Methodenaufruf")
            continue
        }
        $mn = $m.Member.Value
        if (($Script:VerboteneMethoden -contains $mn) -or ($mn -match $Script:VerbotenePraefixe)) {
            $ausnahme = $Script:MethodenAusnahmen[$mn]
            $fn = Get-EnclosingFunctionName $m
            if (-not ($ausnahme -and $fn -eq $ausnahme)) { $verstoesse.Add("Zeile ${zeile}: Methode '$mn' kann schreiben und ist hier verboten") }
        }
        if ($mn -eq 'OpenSubKey' -and $m.Extent.Text -notmatch 'RegistryRights\]::ReadKey') {
            $verstoesse.Add("Zeile ${zeile}: Registry-Schluessel muss schreibgeschuetzt (ReadKey) geoeffnet werden")
        }
    }

    foreach ($r in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FileRedirectionAst] }, $true)) {
        $verstoesse.Add("Zeile $($r.Extent.StartLineNumber): Umleitung in eine Datei ist verboten")
    }

    # PowerShell wertet „ “ ” als Anfuehrungszeichen - in doppelt gequotetem Text beenden sie ihn vorzeitig
    foreach ($tk in @($tokens)) {
        if ([string]$tk.Kind -in @('StringExpandable', 'HereStringExpandable') -and $tk.Text -match '[„“”]') {
            $verstoesse.Add("Zeile $($tk.Extent.StartLineNumber): typografische Anführungszeichen in doppelt gequotetem Text (beenden ihn vorzeitig)")
        }
    }

    # Sicherheitsnetz B muss genau einmal und unveraendert vorhanden sein
    $whatIf = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$WhatIfPreference' }, $true))
    if ($whatIf.Count -ne 1 -or $whatIf[0].Right.Extent.Text -ne '$true') { $verstoesse.Add('$WhatIfPreference = $true fehlt oder wurde veraendert') }

    # C#-Hilfscode auf verbotene Windows-Funktionen pruefen
    $quelltext = [System.IO.File]::ReadAllText($PSCommandPath)
    $marke = '// SD-NATIVE-'
    $anfang = $quelltext.IndexOf($marke + 'BEGIN')
    $ende = $quelltext.IndexOf($marke + 'END')
    if ($anfang -lt 0 -or $ende -lt $anfang) {
        $verstoesse.Add('C#-Hilfscode nicht gefunden')
    } else {
        $cs = $quelltext.Substring($anfang, $ende - $anfang)
        foreach ($t in [regex]::Matches($cs, $Script:VerboteneNativeApis)) { $verstoesse.Add("C#-Hilfscode: verbotene Windows-Funktion '$($t.Value)'") }
    }

    [pscustomobject]@{ Ok = ($verstoesse.Count -eq 0); Verstoesse = @($verstoesse); Befehle = $nBefehle; Methoden = $nMethoden }
}

# Mechanismus C: Registry nur schreibgeschuetzt oeffnen
function Open-RegKeyReadOnly {
    param([ValidateSet('HKLM', 'HKCU')] [string] $Hive, [string] $Pfad)
    $h = if ($Hive -eq 'HKLM') { [Microsoft.Win32.RegistryHive]::LocalMachine } else { [Microsoft.Win32.RegistryHive]::CurrentUser }
    try {
        $basis = [Microsoft.Win32.RegistryKey]::OpenBaseKey($h, [Microsoft.Win32.RegistryView]::Registry64)
        return $basis.OpenSubKey($Pfad, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadSubTree, [System.Security.AccessControl.RegistryRights]::ReadKey)
    } catch {
        return $null
    }
}

function Read-RegValue {
    param([string] $Hive, [string] $Pfad, [string] $Name, $Standard = $null)
    $k = Open-RegKeyReadOnly $Hive $Pfad
    if ($null -eq $k) { return $Standard }
    try {
        $v = $k.GetValue($Name, $null)
        if ($null -eq $v) { return $Standard }
        return , $v
    } catch {
        return $Standard
    } finally {
        $k.Close()
    }
}

function Read-RegValues {
    param([string] $Hive, [string] $Pfad)
    $k = Open-RegKeyReadOnly $Hive $Pfad
    if ($null -eq $k) { return $null }
    $werte = [ordered]@{}
    try {
        foreach ($n in $k.GetValueNames()) {
            if ([string]::IsNullOrEmpty($n)) { continue }
            $werte[$n] = $k.GetValue($n, $null)
        }
    } catch {
    } finally {
        $k.Close()
    }
    return $werte
}

function Read-RegSubKeyNames {
    param([string] $Hive, [string] $Pfad)
    $k = Open-RegKeyReadOnly $Hive $Pfad
    if ($null -eq $k) { return @() }
    try { return @($k.GetSubKeyNames()) } catch { return @() } finally { $k.Close() }
}

function Test-RegKey {
    param([string] $Hive, [string] $Pfad)
    $k = Open-RegKeyReadOnly $Hive $Pfad
    if ($null -eq $k) { return $false }
    $k.Close()
    return $true
}

# Mechanismus D: externe Programme nur aus System32 und nur mit freigegebenen Parametern
function Invoke-SafeTool {
    param([string] $Exe, [string[]] $Argumente)
    $schluessel = $Exe.ToLowerInvariant()
    if (-not $Script:ProgrammFreigaben.ContainsKey($schluessel)) { throw "Programm '$Exe' ist nicht freigegeben" }
    $zeile = ($Argumente -join ' ')
    $erlaubt = $false
    foreach ($muster in $Script:ProgrammFreigaben[$schluessel]) { if ($zeile -match $muster) { $erlaubt = $true } }
    if (-not $erlaubt) { throw "Parameter '$zeile' fuer '$Exe' sind nicht freigegeben" }
    $voll = Join-Path ([Environment]::GetFolderPath('System')) $schluessel
    if (-not (Test-Path -LiteralPath $voll -PathType Leaf)) { throw "'$voll' nicht gefunden" }
    $ausgabe = (& $voll @Argumente 2>&1 | Out-String) -replace "`0", ''
    $code = $LASTEXITCODE
    # Fehlende Rechte melden diese Programme nur als Text - das darf nie als Ergebnis durchgehen
    if ($ausgabe -match '(?i)Administrator|Zugriff verweigert|Access is denied|erh.hte Rechte|elevat') { throw "keine ausreichenden Rechte ($schluessel)" }
    if ($code -ne 0 -and $schluessel -notin @('sfc.exe', 'netsh.exe', 'chkdsk.exe', 'wbem\winmgmt.exe')) { throw "$schluessel meldet Fehlercode $code - $(Format-Kurz $ausgabe 120)" }
    return $ausgabe
}

# Mechanismus E: einziger Schreibzugriff - neue Datei im Berichtsordner, nie ueberschreiben
function Save-ReportFile {
    param([string] $DateiName, [string] $Inhalt)
    if ($DateiName -notmatch '^[\w\-\.]+\.(html|json)$') { throw "Unzulaessiger Dateiname '$DateiName'" }
    $tool = [System.IO.Path]::GetFullPath($Script:DatenOrdner)
    $ordner = [System.IO.Path]::GetFullPath($Script:BerichtOrdner)
    if (-not $ordner.StartsWith($tool + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw 'Berichtsordner liegt ausserhalb des Datenordners' }
    if (-not [System.IO.Directory]::Exists($ordner)) { [void][System.IO.Directory]::CreateDirectory($ordner) }
    $ziel = [System.IO.Path]::GetFullPath((Join-Path $ordner $DateiName))
    if ((Split-Path -Parent $ziel) -ne $ordner) { throw 'Zielpfad liegt ausserhalb des Berichtsordners' }
    if ([System.IO.File]::Exists($ziel)) { throw "Datei '$ziel' existiert bereits - es wird nichts ueberschrieben" }
    [System.IO.File]::WriteAllText($ziel, $Inhalt, (New-Object System.Text.UTF8Encoding($true)))
    return $ziel
}

function Open-Report {
    param([string] $Pfad)
    $voll = [System.IO.Path]::GetFullPath($Pfad)
    $ordner = [System.IO.Path]::GetFullPath($Script:BerichtOrdner)
    if (-not $voll.StartsWith($ordner + '\', [System.StringComparison]::OrdinalIgnoreCase) -or $voll -notmatch '\.html$') { return }
    Invoke-Item -LiteralPath $voll -WhatIf:$false
}

function Initialize-NativeHelper {
    if ('SdNative' -as [type]) { return $true }
    try {
        Add-Type -TypeDefinition $Script:NativeQuelle -Language CSharp -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Test-IstAdmin {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = [System.Security.Principal.WindowsPrincipal]$id
    return $pr.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ============================================================================
#  4. HILFSFUNKTIONEN
# ============================================================================

# Ein Befund: Was ist das (aus $Script:Wissen), Ergebnis, Ursache, Massnahme, Empfehlung und
# optional Loesungsvorschlaege. Die Loesungen sind reine DATEN - ausgefuehrt werden sie nie hier,
# sondern nur vom getrennten Bereinigungsteil, der jede Aktion selbst noch einmal prueft.
function Add-Finding {
    param(
        [Parameter(Mandatory, Position = 0)] [string] $Typ,
        [Parameter(Mandatory, Position = 1)] [string] $Bereich,
        [Parameter(Mandatory, Position = 2)] [string] $Pruefung,
        [Parameter(Mandatory, Position = 3)] [ValidateSet('OK', 'Info', 'Warnung', 'Kritisch')] [string] $Status,
        [Parameter(Position = 4)] [string] $Ergebnis = '',
        [string] $Ursache = '',
        [string] $Massnahme = '',
        [ValidateSet('', 'Nichts zu tun', 'Beheben', 'Entfernen', 'Beobachten', 'Behalten', 'Prüfen')] [string] $Empfehlung = '',
        [string] $Hinweis = '',
        [object] $Details = $null,
        [object[]] $Loesungen = @(),
        # Bei Befunden aus Ereignissen: Zeitpunkt des neuesten Ereignisses (die App erkennt daran neue Vorfaelle)
        [object] $Zuletzt = $null
    )
    if (-not $Empfehlung) { $Empfehlung = if ($Status -eq 'OK') { 'Nichts zu tun' } elseif ($Status -eq 'Info') { 'Behalten' } else { 'Prüfen' } }
    # Hinweise nur, wenn es etwas zu tun oder zu pruefen gibt - reine Uebersicht („Behalten“, „Beobachten“) zaehlt als
    # in Ordnung und steht nur unter „Alle Prüfungen“ (Wunsch 28.10.: keine Hinweis-Flut)
    if ($Status -eq 'Info' -and $Empfehlung -in 'Behalten', 'Beobachten' -and -not @($Loesungen | Where-Object { $_ -and $_.Id -ne 'SeiteOeffnen' }).Count) { $Status = 'OK' }
    if (-not $Massnahme -and $Status -eq 'OK') { $Massnahme = 'Nichts zu tun.' }
    $was =if ($Script:Wissen.ContainsKey($Typ)) { $Script:Wissen[$Typ] } else { '' }
    $Script:Befunde.Add([pscustomobject]@{
        Nr = $Script:Befunde.Count; Typ = $Typ; Bereich = $Bereich; Pruefung = $Pruefung; Status = $Status
        Ergebnis = $Ergebnis; Was = $was; Ursache = $Ursache; Massnahme = $Massnahme; Empfehlung = $Empfehlung
        Hinweis = $Hinweis; Details = $Details; Loesungen = @($Loesungen | Where-Object { $_ })
        Zuletzt = $(if ($Zuletzt) { ([datetime]$Zuletzt).ToString('yyyy-MM-dd HH:mm:ss') } else { '' })
    })
}

# Neuester Zeitpunkt aus einer Liste (leere Eintraege werden uebergangen)
function Get-Neuestes {
    param([object[]] $Zeiten)
    return @($Zeiten | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1)[0]
}

# Beschreibt eine Loesung (nur Daten). Art 'Bereinigung' = braucht Adminrechte, laeuft im
# Bereinigungsteil; Art 'Sofort' = harmlose Aktion der App (z. B. Einstellungsseite oeffnen).
function New-Loesung {
    param(
        [string] $Id,
        [string] $Titel,
        [hashtable] $Parameter = @{},
        [ValidateSet('Bereinigung', 'Sofort')] [string] $Art = 'Bereinigung',
        [bool] $Umkehrbar = $true,
        [string] $Dauer = '',
        [string] $Beschreibung = ''
    )
    return [ordered]@{ Id = $Id; Titel = $Titel; Art = $Art; Umkehrbar = $Umkehrbar; Dauer = $Dauer; Beschreibung = $Beschreibung; Parameter = $Parameter }
}

function New-SeitenLoesung {
    param([string] $Titel, [string] $Seite)
    return New-Loesung -Id 'SeiteOeffnen' -Titel $Titel -Art 'Sofort' -Parameter @{ Seite = $Seite }
}

# Fuer Schritte, die bewusst von Hand bestaetigt werden: die App kopiert einen Befehl aus ihrer eigenen festen Liste
# (hier steht nur die Kennung) und oeffnet eine Admin-PowerShell - ausgefuehrt wird erst nach Einfuegen + Enter.
function New-BefehlLoesung {
    param([string] $Titel, [string] $Befehl, [string] $Laufwerk = '')
    return New-Loesung -Id 'BefehlKopieren' -Titel $Titel -Art 'Sofort' -Parameter @{ Befehl = $Befehl; Laufwerk = ([string]$Laufwerk).TrimEnd(':') }
}

function Get-StatusRang {
    param([string] $Status)
    switch ($Status) { 'Kritisch' { 0 } 'Warnung' { 1 } 'Info' { 2 } default { 3 } }
}

# Kennung dieses PCs: Rechnername + Mainboard-UUID (Platzhalter-UUIDs mancher Boards zaehlen nicht)
function Get-MaschinenId {
    if ($Maschine) { return $Maschine }
    $uuid = ''
    try { $uuid = [string](Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop).UUID } catch { }
    if ($uuid -match '^(0{8}|F{8})-') { $uuid = '' }
    return "$env:COMPUTERNAME|$uuid"
}

function Get-ProfilDateiName { return ($Script:MaschinenId -replace '[^A-Za-z0-9\-]', '_') + '.json' }

# Welches Profil gilt? PC 1 = eingebautes Profil; jeder andere PC startet neutral (keine Sollwerte, keine
# Geschichte von PC 1). Eingelernte Sollwerte aus Profile\<Kennung>.json (schreibt die App) haben Vorrang.
function Initialize-Profil {
    $Script:MaschinenId = Get-MaschinenId
    $Script:IstEigenerPc = ($Script:MaschinenId -eq $Script:Pc1Id)
    $Script:ProfilQuelle = 'eingebaut'
    if (-not $Script:IstEigenerPc) {
        $Script:ProfilQuelle = 'neu'
        $Script:Profil.Name = "$env:COMPUTERNAME (automatisch erkannt)"
        $Script:Profil.EnergieplanMuster = ''
        $Script:Profil.RamMhz = 0
        $Script:Profil.RamModule = 0
        $Script:Profil.DiensteMuessenLaufen = [ordered]@{}
        $Script:Profil.DiensteMuessenAusSein = [ordered]@{}
        $Script:Profil.Monitore = @()
        $Script:Profil.LaufwerkAusnahmen = [ordered]@{}
        $Script:BekanntesRauschen = @($Script:BekanntesRauschen | Where-Object { $_.Quelle -ne 'e2fnexpress' })
        $Script:Wissen['WheaStart'] = 'Sucht ein bestimmtes Muster: PC ausgeschaltet, kurz darauf wieder eingeschaltet und in den ersten Minuten ein schwerer Hardwarefehler. Das deutet meist auf die Stromversorgung beim Kaltstart, nicht auf einen Defekt im laufenden Betrieb.'
        $Script:Wissen['DynamicLighting'] = 'Eine Windows-Funktion, die RGB-Geräte zentral steuert. Sie kann sich mit Hersteller-Software (z. B. Logitech G HUB, Corsair iCUE) um die Beleuchtung streiten.'
    }
    $datei = Join-Path $Script:DatenOrdner ('Profile\' + (Get-ProfilDateiName))
    if (Test-Path -LiteralPath $datei) {
        try {
            $p = Get-Content -LiteralPath $datei -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($p.Name) { $Script:Profil.Name = [string]$p.Name }
            if ($p.Monitore) { $Script:Profil.Monitore = @(@($p.Monitore) | ForEach-Object { [ordered]@{ Edid = [string]$_.Edid; Name = [string]$_.Name; Aufloesung = [string]$_.Aufloesung; Hz = [double]$_.Hz; Hdr = [bool]$_.Hdr } }) }
            if ($p.RamModule) { $Script:Profil.RamModule = [int]$p.RamModule }
            if ($p.RamMhz) { $Script:Profil.RamMhz = [int]$p.RamMhz }
            if ($p.Energieplan) {
                $Script:Profil.EnergieplanMuster = '^' + [regex]::Escape([string]$p.Energieplan) + '$'
                # Fuer die Behebung den exakten Namen weitergeben - der Bereiniger nimmt kein Regex mit Escapes an
                $Script:Profil['EnergieplanName'] = [string]$p.Energieplan
            }
            $Script:ProfilQuelle = 'eingelernt'
            $Script:ProfilStand = [string]$p.Eingelernt
        } catch { }
    }
}

# Kurzbeschreibung des PCs fuer die App (Profil einlernen) - nur bereits gelesene Daten
function Get-MaschinenDaten {
    $cs = Get-Data 'ComputerSystem'
    $ram = @(Get-Data 'RAM')
    $plan = Get-Data 'PowerScheme'
    [ordered]@{
        Id = $Script:MaschinenId; Name = $env:COMPUTERNAME; ProfilQuelle = $Script:ProfilQuelle; ProfilStand = $Script:ProfilStand
        Hersteller = $(if ($cs) { [string]$cs.Manufacturer } else { '' }); Modell = $(if ($cs) { [string]$cs.Model } else { '' })
        Cpu = (@(@(Get-Data 'CPU') | ForEach-Object { ([string]$_.Name).Trim() }) -join ', ')
        Gpu = (@(@(Get-Data 'VideoController') | ForEach-Object { [string]$_.Name }) -join ', ')
        RamModule = $ram.Count
        RamMhz = $(if ($ram.Count) { [int]($ram | Measure-Object -Property ConfiguredClockSpeed -Minimum).Minimum } else { 0 })
        RamGb = $(if ($ram.Count) { [Math]::Round((($ram | Measure-Object -Property Capacity -Sum).Sum / 1GB)) } else { 0 })
        Energieplan = $(if ($plan -and $plan -match '\((.+)\)') { $Matches[1] } else { '' })
        Monitore = @(@(Get-Data 'Display') | ForEach-Object { [ordered]@{ Edid = $_.Edid; Name = $_.Name; Aufloesung = $_.Aufloesung; Hz = [Math]::Round([double]$_.Hz); HdrAktiv = [bool]$_.HdrAktiv } })
    }
}

# Zerlegt eine Befehlszeile / einen Dienstpfad in den reinen Dateipfad
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

# $true = vorhanden, $false = fehlt, $null = nicht pruefbar
function Test-FileExists {
    param([string] $Pfad)
    if ([string]::IsNullOrWhiteSpace($Pfad)) { return $null }
    if ($Pfad -notmatch '\\') { return $true }
    if ($Pfad -match '(?i)\\WindowsApps\\') { return $null }
    return [bool](Test-Path -LiteralPath $Pfad -PathType Leaf -ErrorAction SilentlyContinue)
}

function Get-EventText {
    param($Ereignis)
    try {
        $m = $Ereignis.Message
        if ($m) { return [string]$m }
    } catch { }
    return (@($Ereignis.Properties | ForEach-Object { [string]$_.Value }) -join ' ')
}

function Get-EventDataMap {
    param($Ereignis)
    $karte = @{}
    try {
        $x = [xml]$Ereignis.ToXml()
        foreach ($d in @($x.Event.EventData.Data)) { if ($d.Name) { $karte[$d.Name] = $d.'#text' } }
    } catch { }
    return $karte
}

function Get-NoiseReason {
    param($Ereignis)
    foreach ($n in $Script:BekanntesRauschen) {
        if ($Ereignis.ProviderName -ne $n.Quelle -or $Ereignis.Id -ne $n.Id) { continue }
        if ($n.Muster -and ((Get-EventText $Ereignis) -notmatch $n.Muster)) { continue }
        return $n.Grund
    }
    return $null
}

function Format-Kurz {
    param([string] $Text, [int] $Laenge = 160)
    if (-not $Text) { return '' }
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -le $Laenge) { return $t }
    return $t.Substring(0, $Laenge) + ' ...'
}

# "09/02/2026 32.0.31041.3013" -> [version] 32.0.31041.3013 (ungueltig -> 0.0)
function ConvertTo-TreiberVersion {
    param([string] $Text)
    $teil = @(([string]$Text).Trim() -split '\s+')[-1]
    $v = $null
    if ([version]::TryParse($teil, [ref]$v)) { return $v }
    return [version]'0.0'
}

function Format-Datum {
    param($Wert)
    if ($null -eq $Wert) { return '' }
    try { return ([datetime]$Wert).ToString('dd.MM.yyyy HH:mm') } catch { return [string]$Wert }
}

# ============================================================================
#  5. DATENQUELLEN - jede wird genau einmal gelesen (Cache)
# ============================================================================

function Get-Data {
    param([Parameter(Mandatory)] [string] $Name)
    # Listen werden ungepackt zurueckgegeben (Pipeline-tauglich), fehlende Werte ganz ohne Ausgabe
    if ($Script:Cache.ContainsKey($Name)) {
        if ($null -eq $Script:Cache[$Name]) { return }
        return $Script:Cache[$Name]
    }

    $titel = if ($Script:QuellenNamen.ContainsKey($Name)) { $Script:QuellenNamen[$Name] } else { $Name }
    $uebersprungen = $null
    if ($Script:NurAdmin -contains $Name -and -not $Script:AdminModus) { $uebersprungen = 'nur im Gründlich-Modus (mit Adminrechten)' }
    elseif ($Script:NurGruendlich -contains $Name -and $Modus -ne 'Gruendlich') { $uebersprungen = 'nur im Gründlich-Modus' }
    if ($uebersprungen) {
        $Script:Cache[$Name] = $null
        $Script:Quellen.Add([pscustomobject]@{ Datenquelle = $titel; Sekunden = 0; Status = "übersprungen ($uebersprungen)" })
        return
    }

    Write-Progress -Activity 'Systemdiagnose' -Status "Lese: $titel"
    $uhr = [System.Diagnostics.Stopwatch]::StartNew()
    $wert = $null
    $fehler = $null
    try {
        switch ($Name) {
            'OS'                { $wert = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop }
            'WinVersion'        { $wert = Read-RegValues 'HKLM' 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' }
            'ComputerSystem'    { $wert = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop }
            'BIOS'              { $wert = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop }
            'CPU'               { $wert = @(Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop) }
            'RAM'               { $wert = @(Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction Stop) }
            'VideoController'   { $wert = @(Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop) }
            'PnpDevices'        { $wert = @(Get-PnpDevice -ErrorAction Stop) }
            'SignedDrivers'     { $wert = @(Get-CimInstance -ClassName Win32_PnPSignedDriver -ErrorAction Stop | Select-Object DeviceName, InfName, DriverVersion, DriverProviderName) }
            'DriverStore'       { $wert = @(Read-DriverStore) }
            'DriverClasses'     { $wert = @(Read-DriverClasses) }
            'Services'          { $wert = @(Read-Services) }
            'SystemDrivers'     { $wert = @(Read-SystemDrivers) }
            'Processes'         { $wert = @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop | Select-Object Name, ProcessId, CommandLine, CreationDate) }
            'EventsSystem'      { $wert = @(Read-EventWindow 'System' "Level=1 or Level=2 or Level=3 or EventID=1074 or EventID=6005 or EventID=6006 or (Provider[@Name='Microsoft-Windows-Time-Service'] and (EventID=35 or EventID=37)) or (Provider[@Name='Microsoft-Windows-WindowsUpdateClient'] and EventID=19)") }
            'EventsApplication' { $wert = @(Read-EventWindow 'Application' "Level=1 or Level=2 or (Provider[@Name='Windows Error Reporting'] and EventID=1001)") }
            'Reliability'       { $wert = @(Get-CimInstance -ClassName Win32_ReliabilityStabilityMetrics -ErrorAction Stop | Select-Object TimeGenerated, SystemStabilityIndex) }
            'ZuverlEreignisse'  { $wert = @(Read-ZuverlEreignisse) }
            'Volumes'           { $wert = @(Get-Volume -ErrorAction Stop) }
            'PhysicalDisks'     { $wert = @(Read-PhysicalDisks) }
            'Smart'             { $wert = @(Read-SmartData) }
            'DirtyBits'         { $wert = @(Read-DirtyBits) }
            'Uninstall'         { $wert = @(Read-UninstallEntries) }
            'Autostart'         { $wert = @(Read-AutostartEntries) }
            'Tasks'             { $wert = @(Read-ScheduledTasks) }
            'Display'           { $wert = @(Read-Display) }
            'CursorClip'        { if (-not (Initialize-NativeHelper)) { throw 'C#-Hilfscode nicht ladbar' }; $wert = [SdNative]::CursorClip() }
            'NetAdapters'       { $wert = @(Get-NetAdapter -Physical -ErrorAction Stop | Select-Object Name, InterfaceDescription, Status, LinkSpeed) }
            'Wlan'              { $wert = Read-Wlan }
            'IpConfig'          { $wert = @(Get-NetIPConfiguration -ErrorAction Stop) }
            'DnsTest'           { $wert = Read-DnsTest }
            'GatewayTest'       { $wert = Read-GatewayTest }
            'Defender'          { $wert = Get-MpComputerStatus -ErrorAction Stop }
            'Firewall'          { $wert = @(Get-NetFirewallProfile -ErrorAction Stop | Select-Object Name, Enabled) }
            'HotFix'            { $wert = @(Get-HotFix -ErrorAction Stop) }
            'PowerScheme'       { $wert = Invoke-SafeTool -Exe 'powercfg.exe' -Argumente @('/getactivescheme') }
            'TimeStatus'        { $wert = Invoke-SafeTool -Exe 'w32tm.exe' -Argumente @('/query', '/status') }
            'Dumps'             { $wert = Read-DumpFiles }
            'WerFolders'        { $wert = Read-WerFolders }
            'Reinigung'         { $wert = @(Read-Reinigung) }
            'Settings'          { $wert = Read-SettingsRegistry }
            'DeviceGuard'       { $wert = Get-CimInstance -Namespace 'root\Microsoft\Windows\DeviceGuard' -ClassName Win32_DeviceGuard -ErrorAction Stop }
            'PageFile'          { $wert = @(Get-CimInstance -ClassName Win32_PageFileUsage -ErrorAction Stop) }
            'PendingReboot'     { $wert = Read-PendingReboot }
            'Insider'           { $wert = Read-RegValues 'HKLM' 'SOFTWARE\Microsoft\WindowsSelfHost\Applicability' }
            'Activation'        { $wert = @(Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" -ErrorAction Stop | Select-Object Name, LicenseStatus) }
            'Tpm'               { $wert = Get-Tpm -ErrorAction Stop }
            'RestorePoints'     { $wert = @(Get-ComputerRestorePoint -ErrorAction Stop) }
            'ShadowStorage'     { $wert = Invoke-SafeTool -Exe 'vssadmin.exe' -Argumente @('list', 'shadowstorage') }
            'DismHealth'        { $wert = Invoke-SafeTool -Exe 'dism.exe' -Argumente @('/online', '/cleanup-image', '/checkhealth') }
            'DismAnalyse'       { $wert = Invoke-SafeTool -Exe 'dism.exe' -Argumente @('/online', '/cleanup-image', '/analyzecomponentstore') }
            'SfcVerify'         { $wert = Invoke-SafeTool -Exe 'sfc.exe' -Argumente @('/verifyonly') }
            'DismScan'          { $wert = Invoke-SafeTool -Exe 'dism.exe' -Argumente @('/online', '/cleanup-image', '/scanhealth') }
            'Chkdsk'            { $wert = @(Read-Chkdsk) }
            'FsReparatur'       { $wert = @(Read-FsReparatur) }
            'Trim'              { $wert = Invoke-SafeTool -Exe 'fsutil.exe' -Argumente @('behavior', 'query', 'DisableDeleteNotify') }
            'Bootkonfiguration' { $wert = Invoke-SafeTool -Exe 'bcdedit.exe' -Argumente @('/enum', '{current}') }
            'WmiDatenbank'      { $wert = Invoke-SafeTool -Exe 'wbem\winmgmt.exe' -Argumente @('/verifyrepository') }
            'Speicherbelegung'  { $wert = @(Read-Speicherbelegung) }
            'Treiberpruefung'   { $wert = Read-RegValues 'HKLM' 'SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management' }
            'Bedrohungen'       { $wert = @(Read-Bedrohungen) }
            'Hosts'             { $wert = @(Get-Content -LiteralPath (Join-Path ([Environment]::GetFolderPath('System')) 'drivers\etc\hosts') -ErrorAction Stop) }
            'NtfsEreignisse'    { $wert = @(Read-EventWindow 'System' "(Provider[@Name='Ntfs'] or Provider[@Name='Microsoft-Windows-Ntfs']) and (Level=1 or Level=2 or Level=3)") }
            'Speicherdiagnose'  { $wert = @(Read-Speicherdiagnose) }
            'Signaturen'        { $wert = @(Read-Signaturen) }
            default            { throw "Unbekannte Datenquelle '$Name'" }
        }
    } catch {
        $fehler = $_.Exception.Message
        $Script:QuellenFehler[$Name] = $fehler
    }
    $sek = $uhr.Elapsed.TotalSeconds
    $status = if ($fehler) { 'nicht lesbar: ' + (Format-Kurz $fehler 120) } else { 'gelesen' }
    $Script:Quellen.Add([pscustomobject]@{ Datenquelle = $titel; Sekunden = [Math]::Round($sek, 2); Status = $status })
    $farbe = if ($fehler) { 'DarkYellow' } else { 'DarkGray' }
    Write-Host ('    {0,-36} {1,6:N1} s{2}' -f $titel, $sek, $(if ($fehler) { '  (nicht lesbar)' } else { '' })) -ForegroundColor $farbe
    $Script:Cache[$Name] = $wert
    if ($null -eq $wert) { return }
    return $wert
}

# Meldet eine Quelle, die gelesen werden sollte, aber nicht lesbar war, als Info-Befund.
# Rueckgabe $true = Quelle war nicht lesbar (die Pruefung soll dann nichts weiter behaupten).
function Test-QuelleFehlt {
    param([string] $Name, [string] $Typ, [string] $Bereich, [string] $Pruefung)
    if (-not $Script:QuellenFehler.ContainsKey($Name)) { return $false }
    Add-Finding $Typ $Bereich $Pruefung 'Info' ('nicht lesbar: ' + (Format-Kurz $Script:QuellenFehler[$Name] 140)) -Ursache 'Windows hat die Abfrage abgelehnt - meist fehlen Adminrechte.' -Massnahme 'Im Modus „Gründlich“ (mit Adminrechten) erneut prüfen.' -Empfehlung 'Prüfen'
    return $true
}

function Read-EventWindow {
    param([string] $Protokoll, [string] $Bedingung)
    $ms = [int64]$Tage * 86400000
    $xpath = "*[System[TimeCreated[timediff(@SystemTime) <= $ms] and ($Bedingung)]]"
    # „Keine Treffer“ meldet Get-WinEvent als Fehler (NoMatchingEventsFound) - das ist ein gesundes leeres Ergebnis.
    # Jeder andere Fehler (Zugriff verweigert, Protokoll defekt) wird geworfen -> Get-Data meldet die Quelle als nicht lesbar,
    # statt dass daraus „Keine Abstürze“ wird.
    $lesefehler = $null
    $ereignisse = @(Get-WinEvent -LogName $Protokoll -FilterXPath $xpath -ErrorAction SilentlyContinue -ErrorVariable lesefehler)
    $echt = @($lesefehler | Where-Object { $_.FullyQualifiedErrorId -notmatch 'NoMatchingEventsFound' })
    if ($echt.Count) { throw $echt[0].Exception }
    return $ereignisse
}

function Read-DriverStore {
    $text = Invoke-SafeTool -Exe 'pnputil.exe' -Argumente @('/enum-drivers')
    foreach ($block in ($text -split '(?:\r?\n){2,}')) {
        if ($block -notmatch '(oem\d+\.inf)') { continue }
        $inf = $Matches[1]
        $orig = if ($block -match '(?m)^\s*(?:Originalname|Original Name)\s*:\s*(\S+)') { $Matches[1] } else { '' }
        $anb = if ($block -match '(?m)^\s*(?:Anbietername|Provider Name)\s*:\s*(.+?)\s*$') { $Matches[1] } else { '' }
        $kl = if ($block -match '(?m)^\s*(?:Klassenname|Class Name)\s*:\s*(.+?)\s*$') { $Matches[1] } else { '' }
        $ver = if ($block -match '(?m)^\s*(?:Treiberversion|Driver Version)\s*:\s*(.+?)\s*$') { $Matches[1] } else { '' }
        [pscustomobject]@{ Inf = $inf; Original = $orig; Anbieter = $anb; Klasse = $kl; Version = $ver }
    }
}

function Read-DriverClasses {
    $klassen = [ordered]@{
        'Grafik'          = '{4d36e968-e325-11ce-bfc1-08002be10318}'
        'Netzwerk'        = '{4d36e972-e325-11ce-bfc1-08002be10318}'
        'Audio'           = '{4d36e96c-e325-11ce-bfc1-08002be10318}'
        'Bluetooth'       = '{e0cbf06c-cd8b-4647-bb8a-263b43f0f974}'
        'System/Chipsatz' = '{4d36e97d-e325-11ce-bfc1-08002be10318}'
    }
    foreach ($k in $klassen.Keys) {
        $basis = 'SYSTEM\CurrentControlSet\Control\Class\' + $klassen[$k]
        foreach ($sub in @(Read-RegSubKeyNames 'HKLM' $basis)) {
            if ($sub -notmatch '^\d{4}$') { continue }
            $v = Read-RegValues 'HKLM' ($basis + '\' + $sub)
            if (-not $v -or -not $v['ProviderName'] -or [string]$v['ProviderName'] -eq 'Microsoft') { continue }
            if ([string]$v['DriverDesc'] -match 'WAN Miniport|Microsoft') { continue }
            [pscustomobject]@{
                Klasse = $k; Geraet = [string]$v['DriverDesc']; Anbieter = [string]$v['ProviderName']
                Version = [string]$v['DriverVersion']; Datum = [string]$v['DriverDate']
                Zusatz = $(if ($v['RadeonSoftwareVersion']) { 'Adrenalin ' + $v['RadeonSoftwareVersion'] } else { '' })
            }
        }
    }
}

function Read-Services {
    foreach ($s in @(Get-CimInstance -ClassName Win32_Service -ErrorAction Stop)) {
        $pfad = Get-ExePath $s.PathName
        [pscustomobject]@{ Name = $s.Name; Anzeige = $s.DisplayName; Status = $s.State; Start = $s.StartMode; Pfad = $pfad; Vorhanden = (Test-FileExists $pfad) }
    }
}

function Read-SystemDrivers {
    foreach ($d in @(Get-CimInstance -ClassName Win32_SystemDriver -ErrorAction Stop)) {
        $pfad = Get-ExePath $d.PathName
        $da = Test-FileExists $pfad
        $hersteller = ''
        if ($da -and $d.State -eq 'Running') {
            try { $hersteller = [string](Get-Item -LiteralPath $pfad -Force -ErrorAction Stop).VersionInfo.CompanyName } catch { $hersteller = '' }
        }
        [pscustomobject]@{ Name = $d.Name; Status = $d.State; Start = $d.StartMode; Pfad = $pfad; Vorhanden = $da; Hersteller = $hersteller }
    }
}

function Read-PhysicalDisks {
    foreach ($d in @(Get-PhysicalDisk -ErrorAction Stop)) {
        $r = $null
        if ($Script:AdminModus) { try { $r = $d | Get-StorageReliabilityCounter -ErrorAction Stop } catch { $r = $null } }
        [pscustomobject]@{
            Nr = $d.DeviceId; Modell = $d.FriendlyName; Seriennummer = ([string]$d.SerialNumber).Trim(); Firmware = $d.FirmwareVersion
            Typ = [string]$d.MediaType; Bus = [string]$d.BusType; Zustand = [string]$d.HealthStatus
            Betrieb = (@($d.OperationalStatus) -join ','); GB = [Math]::Round($d.Size / 1GB)
            TempC = $(if ($r) { $r.Temperature }); Verschleiss = $(if ($r) { $r.Wear })
            LeseFehler = $(if ($r) { $r.ReadErrorsUncorrected }); SchreibFehler = $(if ($r) { $r.WriteErrorsUncorrected })
            Stunden = $(if ($r) { $r.PowerOnHours })
        }
    }
}

function Read-SmartData {
    $namen = @{
        5 = 'Neu zugewiesene Sektoren'; 9 = 'Betriebsstunden'; 177 = 'Wear Leveling Count'; 179 = 'Reserveblöcke genutzt'
        181 = 'Program-Fehler'; 182 = 'Erase-Fehler'; 183 = 'Bad Blocks'; 187 = 'Unkorrigierbare Fehler'; 194 = 'Temperatur'
        195 = 'ECC-Fehler'; 196 = 'Neuzuweisungs-Ereignisse'; 197 = 'Ausstehende Sektoren'; 198 = 'Offline unkorrigierbar'
        199 = 'CRC-Fehler (Kabel)'; 241 = 'Geschrieben (LBA)'
    }
    $laufwerke = @(Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue)
    $prognosen = @(Get-CimInstance -Namespace 'root\wmi' -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue)
    foreach ($s in @(Get-CimInstance -Namespace 'root\wmi' -ClassName MSStorageDriver_FailurePredictData -ErrorAction Stop)) {
        $inst = $s.InstanceName -replace '_0$', ''
        $lw = $laufwerke | Where-Object { $_.PNPDeviceID -eq $inst } | Select-Object -First 1
        $b = $s.VendorSpecific
        $attr = [ordered]@{}
        for ($o = 2; $o -lt 362; $o += 12) {
            $id = [int]$b[$o]
            if ($id -eq 0 -or -not $namen.ContainsKey($id)) { continue }
            $roh = [uint64]0
            for ($k = 5; $k -ge 0; $k--) { $roh = $roh * 256 + $b[$o + 5 + $k] }
            if ($id -eq 194) { $roh = $roh -band 0xFF }
            $attr[[string]$id] = [pscustomobject]@{ Id = $id; Name = $namen[$id]; Zustand = [int]$b[$o + 3]; Rohwert = $roh }
        }
        $pf = $prognosen | Where-Object { $_.InstanceName -eq $s.InstanceName } | Select-Object -First 1
        [pscustomobject]@{
            Modell = $(if ($lw) { $lw.Model } else { $inst }); Seriennummer = $(if ($lw) { ([string]$lw.SerialNumber).Trim() } else { '' })
            Firmware = $(if ($lw) { $lw.FirmwareRevision } else { '' }); Ausfallprognose = $(if ($pf) { [bool]$pf.PredictFailure } else { $null })
            Attribute = $attr
        }
    }
}

# Dateisystem jedes festen NTFS-Laufwerks im Lesemodus pruefen (chkdsk ohne Schalter repariert nichts)
function Read-Chkdsk {
    foreach ($v in @(Get-Data 'Volumes' | Where-Object { $_.DriveLetter -and $_.FileSystem -eq 'NTFS' -and [string]$_.DriveType -eq 'Fixed' })) {
        $l = [string]$v.DriveLetter
        $start = Get-Date
        $t = Invoke-SafeTool -Exe 'chkdsk.exe' -Argumente @($l + ':')
        # Erfolg zuerst pruefen - „keine Probleme festgestellt“ enthaelt sonst auch „Probleme“
        $ok = if ($t -match '(?i)keine Probleme|found no problems|no further action is required') { $true } elseif ($t -match '(?i)Probleme (gefunden|festgestellt)|Fehler (gefunden|festgestellt)|found problems|errors found|beschädigt|corrupt') { $false } else { $null }
        $zeilen = @(($t -split "`r?`n") | Where-Object { $_ -match '(?i)problem|fehler|error|beschädigt|corrupt' } | Select-Object -Last 3)
        [pscustomobject]@{ Laufwerk = $l; InOrdnung = $ok; Sekunden = [Math]::Round(((Get-Date) - $start).TotalSeconds); Meldung = (Format-Kurz (@($zeilen) -join ' / ') 200) }
    }
}

# Reparaturstatus je NTFS-Laufwerk (fsutil repair state, nur Abfrage) - 0x0 = keine bekannte Beschaedigung
function Read-FsReparatur {
    foreach ($v in @(Get-Data 'Volumes' | Where-Object { $_.DriveLetter -and $_.FileSystem -eq 'NTFS' -and [string]$_.DriveType -eq 'Fixed' })) {
        $l = [string]$v.DriveLetter
        $t = Invoke-SafeTool -Exe 'fsutil.exe' -Argumente @('repair', 'state', ($l + ':'))
        $wert = if ($t -match '0x([0-9a-f]+)') { [Convert]::ToInt32($Matches[1], 16) } else { $null }
        [pscustomobject]@{ Laufwerk = $l; InOrdnung = $(if ($null -eq $wert) { $null } else { $wert -eq 0 }); Text = (Format-Kurz $t 120) }
    }
}

# Groesste Ordner (Ebene 1 und 2) jedes festen NTFS-Laufwerks - komplette Begehung, nur im Gruendlich-Modus
function Read-Speicherbelegung {
    if (-not (Initialize-NativeHelper)) { throw 'C#-Hilfscode nicht ladbar' }
    foreach ($v in @(Get-Data 'Volumes' | Where-Object { $_.DriveLetter -and $_.FileSystem -eq 'NTFS' -and [string]$_.DriveType -eq 'Fixed' -and $_.Size -gt 0 })) {
        $l = [string]$v.DriveLetter
        $uhr = [System.Diagnostics.Stopwatch]::StartNew()
        $ordner = @([SdBelegung]::Messen($l + ':\', 2))
        [pscustomobject]@{
            Laufwerk = $l; Groesse = [int64]$v.Size; Frei = [int64]$v.SizeRemaining; Sekunden = [Math]::Round($uhr.Elapsed.TotalSeconds, 1)
            Ordner = @($ordner | Sort-Object Bytes -Descending | Select-Object -First 60 | ForEach-Object { [pscustomobject]@{ Pfad = $_.Pfad; Ebene = $_.Ebene; Bytes = $_.Bytes; Dateien = $_.Dateien } })
        }
    }
}

# Die Ausfaelle, die Windows in den Zuverlaessigkeitsindex einrechnet (Abschaltung, Absturz, Haenger, Bluescreen)
function Read-ZuverlEreignisse {
    $filter = "EventIdentifier=6008 OR EventIdentifier=41 OR EventIdentifier=1000 OR EventIdentifier=1002 OR EventIdentifier=1001"
    foreach ($r in @(Get-CimInstance -ClassName Win32_ReliabilityRecords -Filter $filter -ErrorAction Stop)) {
        $quelle = [string]$r.SourceName; $id = [int]$r.EventIdentifier
        $text = if ($id -eq 6008 -or $id -eq 41) { 'Unerwartete Abschaltung' }
            elseif ($id -eq 1000 -and $quelle -eq 'Application Error') { "Programmabsturz: $($r.ProductName)" }
            elseif ($id -eq 1002 -and $quelle -eq 'Application Hang') { "Programm hing: $($r.ProductName)" }
            elseif ($id -eq 1001 -and $quelle -match 'SystemErrorReporting|BugCheck') { 'Bluescreen' }
            else { $null }
        if ($text) { [pscustomobject]@{ Zeit = [datetime]$r.TimeGenerated; Text = $text } }
    }
}

# Funde von Microsoft Defender mit Namen der Bedrohung (nur lesen)
function Read-Bedrohungen {
    $namen = @{}
    foreach ($t in @(Get-MpThreat -ErrorAction SilentlyContinue)) { $namen[[string]$t.ThreatID] = [string]$t.ThreatName }
    foreach ($d in @(Get-MpThreatDetection -ErrorAction Stop)) {
        [pscustomobject]@{
            Zeit = $d.InitialDetectionTime; Name = $(if ($namen.ContainsKey([string]$d.ThreatID)) { $namen[[string]$d.ThreatID] } else { "Id $($d.ThreatID)" })
            Erfolgreich = [bool]$d.ActionSuccess; Dateien = (Format-Kurz (@($d.Resources) -join ', ') 160)
        }
    }
}

# Letzte Ergebnisse der Windows-Speicherdiagnose (mdsched) - unabhaengig vom Zeitraum, sie laeuft selten
function Read-Speicherdiagnose {
    $lesefehler = $null
    $e = @(Get-WinEvent -LogName 'System' -FilterXPath "*[System[Provider[@Name='Microsoft-Windows-MemoryDiagnostics-Results']]]" -MaxEvents 5 -ErrorAction SilentlyContinue -ErrorVariable lesefehler)
    $echt = @($lesefehler | Where-Object { $_.FullyQualifiedErrorId -notmatch 'NoMatchingEventsFound' })
    if ($echt.Count) { throw $echt[0].Exception }
    foreach ($x in $e) { [pscustomobject]@{ Zeit = $x.TimeCreated; Id = $x.Id; Meldung = (Format-Kurz $x.Message 200) } }
}

# Digitale Signatur jeder Kernel-Treiber- und Autostart-Datei (Get-AuthenticodeSignature liest nur)
function Read-Signaturen {
    $gesehen = @{}
    $ziele = @(@(Get-Data 'SystemDrivers' | Where-Object { $_.Vorhanden } | ForEach-Object { [pscustomobject]@{ Art = 'Treiber'; Name = $_.Name; Datei = $_.Pfad } }) +
               @(Get-Data 'Autostart' | Where-Object { $_.ZielVorhanden } | ForEach-Object { [pscustomobject]@{ Art = 'Autostart'; Name = $_.Name; Datei = $_.Ziel } }) +
               @(Get-Data 'Services' | Where-Object { $_.Vorhanden -and $_.Start -ne 'Disabled' } | ForEach-Object { [pscustomobject]@{ Art = 'Dienst'; Name = $_.Name; Datei = $_.Pfad } }) +
               @(Get-Data 'Tasks' | Where-Object { $_.ProgrammVorhanden } | ForEach-Object { [pscustomobject]@{ Art = 'Aufgabe'; Name = $_.Aufgabe; Datei = $_.Programm } }))
    foreach ($z in $ziele) {
        $schluessel = ([string]$z.Datei).ToLowerInvariant()
        # Nur Programmdateien - Verknuepfungen und Skripte koennen keine Authenticode-Signatur tragen
        if (-not $schluessel -or $gesehen.ContainsKey($schluessel) -or $schluessel -notmatch '\.(exe|dll|sys|com|scr)$') { continue }
        $gesehen[$schluessel] = $true
        # Aufgaben mit nacktem „powershell.exe“ loest Get-ExePath nach System32 auf - dort liegt die Datei nicht
        if (-not (Test-FileExists $z.Datei)) { continue }
        $status = 'nicht prüfbar'; $von = ''
        try {
            $s = Get-AuthenticodeSignature -LiteralPath $z.Datei -ErrorAction Stop
            $status = [string]$s.Status
            if ($s.SignerCertificate) { $von = ($s.SignerCertificate.Subject -replace '^CN=("?)([^,"]+).*$', '$2') }
        } catch { $status = 'nicht prüfbar' }
        [pscustomobject]@{ Art = $z.Art; Name = $z.Name; Datei = $z.Datei; Status = $status; Signiert = $von }
    }
}

function Read-DirtyBits {
    foreach ($v in @(Get-Data 'Volumes' | Where-Object { $_.DriveLetter -and $_.FileSystem -eq 'NTFS' -and [string]$_.DriveType -eq 'Fixed' })) {
        $l = [string]$v.DriveLetter
        $t = Invoke-SafeTool -Exe 'fsutil.exe' -Argumente @('dirty', 'query', ($l + ':'))
        $kaputt = if ($t -match '(?i)nicht fehlerhaft|not dirty') { $false } elseif ($t -match '(?i)fehlerhaft|dirty') { $true } else { $null }
        [pscustomobject]@{ Laufwerk = $l; Fehlerhaft = $kaputt; Text = (Format-Kurz $t 120) }
    }
}

function Read-UninstallEntries {
    $wurzeln = @(
        @('HKLM', 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'),
        @('HKLM', 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'),
        @('HKCU', 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
    )
    foreach ($w in $wurzeln) {
        foreach ($sub in @(Read-RegSubKeyNames $w[0] $w[1])) {
            $v = Read-RegValues $w[0] ($w[1] + '\' + $sub)
            if (-not $v -or -not $v['DisplayName']) { continue }
            if ([string]$v['SystemComponent'] -eq '1') { continue }
            [pscustomobject]@{
                Name = ([string]$v['DisplayName']).Trim(); Version = [string]$v['DisplayVersion']; Hersteller = [string]$v['Publisher']
                Installiert = [string]$v['InstallDate']; Deinstallation = [string]$v['UninstallString']
            }
        }
    }
}

function Get-StartupApprovedState {
    param($Freigaben, [string] $Name)
    if (-not $Freigaben -or -not $Freigaben.Contains($Name)) { return 'an' }
    $b = $Freigaben[$Name]
    if ($b -is [byte[]] -and $b.Length -gt 0 -and ($b[0] -band 1)) { return 'aus' }
    return 'an'
}

function Read-AutostartEntries {
    $orte = @(
        [ordered]@{ Ort = 'Run (Benutzer)';         Hive = 'HKCU'; Pfad = 'Software\Microsoft\Windows\CurrentVersion\Run';             Freigabe = 'Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' }
        [ordered]@{ Ort = 'Run (Computer)';         Hive = 'HKLM'; Pfad = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Run';             Freigabe = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' }
        [ordered]@{ Ort = 'Run (Computer, 32 Bit)'; Hive = 'HKLM'; Pfad = 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; Freigabe = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32' }
    )
    foreach ($o in $orte) {
        $werte = Read-RegValues $o.Hive $o.Pfad
        if (-not $werte) { continue }
        $frei = Read-RegValues $o.Hive $o.Freigabe
        foreach ($name in @($werte.Keys)) {
            $befehl = [string]$werte[$name]
            $ziel = Get-ExePath $befehl
            [pscustomobject]@{ Ort = $o.Ort; Name = $name; Status = (Get-StartupApprovedState $frei $name); Ziel = $ziel; ZielVorhanden = (Test-FileExists $ziel) }
        }
    }
    $ordner = @(
        [ordered]@{ Ort = 'Autostart-Ordner (Benutzer)'; Hive = 'HKCU'; Pfad = (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup') }
        [ordered]@{ Ort = 'Autostart-Ordner (alle)';     Hive = 'HKLM'; Pfad = (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup') }
    )
    foreach ($o in $ordner) {
        $frei = Read-RegValues $o.Hive 'Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder'
        foreach ($f in @(Get-ChildItem -LiteralPath $o.Pfad -File -Force -ErrorAction SilentlyContinue)) {
            if ($f.Name -eq 'desktop.ini') { continue }
            [pscustomobject]@{ Ort = $o.Ort; Name = $f.Name; Status = (Get-StartupApprovedState $frei $f.Name); Ziel = $f.FullName; ZielVorhanden = $true }
        }
    }
    # Store-Apps (StartupTask): State 2/4 = an, 0/1/3 = aus (1 = vom Benutzer abgeschaltet)
    $uwp = 'Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData'
    foreach ($paket in @(Read-RegSubKeyNames 'HKCU' $uwp)) {
        foreach ($aufgabe in @(Read-RegSubKeyNames 'HKCU' "$uwp\$paket")) {
            $w = Read-RegValues 'HKCU' "$uwp\$paket\$aufgabe"
            if (-not $w -or -not $w.Contains('State')) { continue }
            $name = ($paket -split '_')[0]
            [pscustomobject]@{ Ort = 'Store-App'; Name = $name; Status = $(if ([int]$w['State'] -in 2, 4) { 'an' } else { 'aus' }); Ziel = $paket; ZielVorhanden = $null }
        }
    }
}

function Read-ScheduledTasks {
    foreach ($t in @(Get-ScheduledTask -ErrorAction Stop | Where-Object { $_.TaskPath -notmatch '^\\Microsoft\\' })) {
        $info = $null
        try { $info = Get-ScheduledTaskInfo -InputObject $t -ErrorAction Stop } catch { $info = $null }
        $a = @($t.Actions)[0]
        $exe = if ($a -and $a.Execute) { Get-ExePath ([string]$a.Execute) } else { $null }
        $lauf = ''
        if ($info -and $info.LastRunTime -and $info.LastRunTime.Year -gt 2000) { $lauf = $info.LastRunTime.ToString('dd.MM.yyyy HH:mm') }
        [pscustomobject]@{
            Aufgabe = ($t.TaskPath + $t.TaskName); Status = [string]$t.State; Programm = $exe
            ProgrammVorhanden = (Test-FileExists $exe); ZuletztGelaufen = $lauf
            Ergebnis = $(if ($info) { '0x{0:X}' -f $info.LastTaskResult } else { '' })
        }
    }
}

function Read-Display {
    if (-not (Initialize-NativeHelper)) { throw 'C#-Hilfscode nicht ladbar' }
    foreach ($m in [SdNative]::Monitors()) {
        [pscustomobject]@{
            Edid = $m.Edid; Name = $m.Name; Gdi = $m.Gdi; Aufloesung = "$($m.Width)x$($m.Height)"; Position = "$($m.X),$($m.Y)"
            Hz = $m.Hz; Primaer = $m.Primary; HdrMoeglich = $m.HdrSupported; HdrAktiv = $m.HdrActive; Farbformat = $m.Encoding; BitProKanal = $m.Bits
            Modi = $(if ($m.Gdi) { @([SdNative]::Modes($m.Gdi)) } else { @() })
        }
    }
}

function Read-Wlan {
    $t = Invoke-SafeTool -Exe 'netsh.exe' -Argumente @('wlan', 'show', 'interfaces')
    $w = [ordered]@{}
    foreach ($z in ($t -split '\r?\n')) {
        if ($z -match '^\s{2,}(.+?)\s*:\s(.*)$') {
            $k = $Matches[1].Trim()
            if (-not $w.Contains($k)) { $w[$k] = $Matches[2].Trim() }
        }
    }
    return $w
}

function Read-DnsTest {
    try {
        $r = @(Resolve-DnsName -Name 'www.microsoft.com' -DnsOnly -QuickTimeout -ErrorAction Stop)
        return [pscustomobject]@{ Ok = ($r.Count -gt 0); Text = "$($r.Count) Antworten" }
    } catch {
        return [pscustomobject]@{ Ok = $false; Text = $_.Exception.Message }
    }
}

function Read-GatewayTest {
    $gw = $null
    foreach ($c in @(Get-Data 'IpConfig')) {
        if ($c.IPv4DefaultGateway) { $gw = @($c.IPv4DefaultGateway)[0].NextHop; break }
    }
    if (-not $gw) { return [pscustomobject]@{ Router = ''; Ok = $false } }
    $ok = [bool](Test-Connection -ComputerName $gw -Count 2 -Quiet -ErrorAction SilentlyContinue)
    return [pscustomobject]@{ Router = $gw; Ok = $ok }
}

function Read-DumpFiles {
    $win = $env:SystemRoot
    [pscustomobject]@{
        Speicherabbild = Get-Item -LiteralPath (Join-Path $win 'MEMORY.DMP') -Force -ErrorAction SilentlyContinue
        Minidumps      = @(Get-ChildItem -LiteralPath (Join-Path $win 'Minidump') -Filter '*.dmp' -Force -ErrorAction SilentlyContinue)
        Programme      = @(Get-ChildItem -LiteralPath (Join-Path $env:LOCALAPPDATA 'CrashDumps') -Filter '*.dmp' -Force -ErrorAction SilentlyContinue)
        LiveKernel     = @(Get-ChildItem -LiteralPath (Join-Path $win 'LiveKernelReports') -Filter '*.dmp' -Recurse -Force -ErrorAction SilentlyContinue)
    }
}

function Read-WerFolders {
    $basis = Join-Path $env:ProgramData 'Microsoft\Windows\WER'
    $q = @(Get-ChildItem -LiteralPath (Join-Path $basis 'ReportQueue') -Directory -Force -ErrorAction SilentlyContinue | Select-Object Name, @{ N = 'Erstellt'; E = { $_.CreationTime } })
    $a = @(Get-ChildItem -LiteralPath (Join-Path $basis 'ReportArchive') -Directory -Force -ErrorAction SilentlyContinue | Select-Object Name, @{ N = 'Erstellt'; E = { $_.CreationTime } })
    [pscustomobject]@{ Warteschlange = $q; Archiv = $a }
}

# --- Reinigungskategorien (Definition in Reinigung.psd1, geloescht wird nur in Bereinigung.ps1) ---
$Script:ReinigungsBasen = @('LocalAppData', 'AppData', 'Windows', 'ProgramData', 'Systemlaufwerk')

function Get-ReinigungsBasis {
    param([string] $Basis)
    switch ($Basis) {
        'LocalAppData'   { return [Environment]::GetFolderPath('LocalApplicationData') }
        'AppData'        { return [Environment]::GetFolderPath('ApplicationData') }
        'Windows'        { return [Environment]::GetFolderPath('Windows') }
        'ProgramData'    { return [Environment]::GetFolderPath('CommonApplicationData') }
        'Systemlaufwerk' { return [System.IO.Path]::GetPathRoot([Environment]::GetFolderPath('Windows')) }
    }
    return $null
}

# Gleiche Formatregeln wie in Bereinigung.ps1 - '' = in Ordnung, sonst der Grund
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

# Loest Basis + Pfad (mit '*' fuer ganze Ordnernamen) in vorhandene Ordner auf - ohne Verknuepfungen zu folgen
function Resolve-ZielOrdner {
    param($Ziel)
    $basis = Get-ReinigungsBasis ([string]$Ziel.Basis)
    if (-not $basis) { return }
    # "C:\" muss so bleiben - "C:" waere der aktuelle Ordner auf Laufwerk C
    $aktuell = @($(if ($basis.Length -gt 3) { $basis.TrimEnd('\') } else { $basis }))
    $pfad = [string]$Ziel.Pfad
    if (-not $pfad) { return $aktuell }
    foreach ($teil in $pfad.Split('\')) {
        $naechste = New-Object System.Collections.Generic.List[string]
        foreach ($o in $aktuell) {
            try {
                foreach ($d in ([System.IO.DirectoryInfo]::new($o)).GetDirectories($teil)) {
                    if (-not ($d.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) { $naechste.Add($d.FullName) }
                }
            } catch { }
        }
        $aktuell = $naechste.ToArray()
        if ($aktuell.Count -eq 0) { return }
    }
    return $aktuell
}

# Zaehlt passende Dateien (Groesse, Anzahl). Verknuepfungen (Junctions/Symlinks) werden nie betreten.
function Measure-Ordner {
    param([string] $Wurzel, [string] $Muster, [bool] $Rekursiv, [int] $MinAlterStunden)
    $grenze = [DateTime]::UtcNow.AddHours(-$MinAlterStunden)
    # Teilweise = ein Unterordner war nicht lesbar (ohne Adminrechte) - die Groesse ist dann ein Mindestwert
    $m = [pscustomobject]@{ Bytes = [int64]0; Dateien = 0; Lesbar = $true; Teilweise = $false }
    $ordner = New-Object System.Collections.Generic.List[string]
    $ordner.Add($Wurzel)
    for ($i = 0; $i -lt $ordner.Count; $i++) {
        try { $eintraege = ([System.IO.DirectoryInfo]::new($ordner[$i])).GetFileSystemInfos() } catch { if ($i -eq 0) { $m.Lesbar = $false } else { $m.Teilweise = $true }; continue }
        foreach ($e in $eintraege) {
            try {
                if ($e.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                if ($e -is [System.IO.DirectoryInfo]) { if ($Rekursiv) { $ordner.Add($e.FullName) }; continue }
                if ($e.Name -like $Muster -and $e.LastWriteTimeUtc -le $grenze) { $m.Bytes += $e.Length; $m.Dateien++ }
            } catch { }   # z. B. ueberlange Pfade
        }
    }
    return $m
}

function Format-Groesse {
    param([int64] $Bytes, [bool] $Messbar = $true)
    if (-not $Messbar) { return 'nur mit Adminrechten messbar' }
    if ($Bytes -ge 1GB) { return '{0:N1} GB' -f ($Bytes / 1GB) }
    return '{0:N0} MB' -f ($Bytes / 1MB)
}

function Read-Reinigung {
    $def = Import-PowerShellDataFile -LiteralPath $Script:ReinigungsDatei -ErrorAction Stop
    if ($def.Format -ne 'Systemdiagnose-Reinigung/1') { throw 'Reinigung.psd1 hat ein unbekanntes Format' }
    $laufend = @{}
    foreach ($p in @(Get-Data 'Processes')) { $laufend[([string]$p.Name -replace '(?i)\.exe$', '').ToLowerInvariant()] = $true }
    $neustart = Get-Data 'PendingReboot'
    foreach ($k in @($def.Kategorien)) {
        $art = if ($k.Art) { [string]$k.Art } else { 'Dateien' }
        $bytes = [int64]0; $dateien = 0; $messbar = $true; $vorhanden = $false; $teilweise = $false
        $jeLaufwerk = [ordered]@{}
        $fehler = New-Object System.Collections.Generic.List[string]
        if ($art -eq 'Papierkorb') {
            $vorhanden = $true
            # Direkt im Dateisystem gemessen - so zaehlen auch Reste mit, die Windows selbst nicht mehr anzeigt
            $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
            foreach ($lw in @([System.IO.DriveInfo]::GetDrives() | Where-Object { [string]$_.DriveType -eq 'Fixed' -and $_.IsReady })) {
                $korb = $lw.RootDirectory.FullName + '$Recycle.Bin\' + $sid
                if (-not [System.IO.Directory]::Exists($korb)) { continue }
                # Nur Eintraege des Papierkorbs ($I = Info, $R = Inhalt) - desktop.ini gehoert zum Ordner selbst
                try { $eintraege = ([System.IO.DirectoryInfo]::new($korb)).GetFileSystemInfos('$*') } catch { $messbar = $false; continue }
                $vorher = $bytes
                foreach ($e in $eintraege) {
                    if ($e.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                    if ($e -is [System.IO.DirectoryInfo]) { $r = Measure-Ordner $e.FullName '*' $true 0; $bytes += $r.Bytes; $dateien += $r.Dateien; if ($r.Teilweise) { $teilweise = $true } }
                    else { $bytes += $e.Length; $dateien++ }
                }
                $jeLaufwerk[$lw.Name.Substring(0, 1)] = $bytes - $vorher
            }
        } elseif ($art -eq 'Lieferoptimierung') {
            $vorhanden = $true
            # Echte Cache-Groesse wie in der Datentraegerbereinigung (PerfSnap). Die Jobliste (Status) zaehlt auch
            # laufende/angeheftete Downloads mit - Laptop 02.11.: Status 6,5 GB, Cache wirklich ~80 MB
            try {
                $st = @(Get-DeliveryOptimizationStatus -ErrorAction Stop -WarningAction SilentlyContinue)
                $dateien = $st.Count
                $bytes = [int64](($st | Measure-Object -Property FileSizeInCache -Sum).Sum)
                $snap = Get-DeliveryOptimizationPerfSnap -ErrorAction Stop -WarningAction SilentlyContinue
                if ($null -ne $snap.CacheSizeBytes) { $bytes = [int64]$snap.CacheSizeBytes; $dateien = [int]$snap.Files }
            } catch { if ($bytes -le 0) { $messbar = $false } }
        } else {
            foreach ($z in @($k.Ziele)) {
                $grund = Test-ZielFormat $z
                if ($grund) { $fehler.Add($grund); continue }
                foreach ($o in @(Resolve-ZielOrdner $z)) {
                    $vorhanden = $true
                    $r = Measure-Ordner $o ([string]$z.Muster) ([bool]$z.Rekursiv) ([int]$z.MinAlterStunden)
                    if (-not $r.Lesbar) { $messbar = $false }
                    if ($r.Teilweise) { $teilweise = $true }
                    $bytes += $r.Bytes; $dateien += $r.Dateien
                }
            }
        }
        if (-not $vorhanden) { continue }
        $laeuft = @(@($k.Prozesse) | Where-Object { $_ -and $laufend.ContainsKey(([string]$_).ToLowerInvariant()) })
        $sperre = if ($k.Bedingung -eq 'KeinNeustartAusstehend' -and $neustart -and @($neustart.Gruende).Count) { 'Updates warten auf einen Neustart' } else { '' }
        Write-Host ('      > ' + $k.Name)
        [pscustomobject]@{
            Id = [string]$k.Id; Gruppe = [string]$k.Gruppe; Name = [string]$k.Name; Beschreibung = [string]$k.Beschreibung
            Hinweis = [string]$k.Hinweis; Vorauswahl = [bool]$k.Vorauswahl; Art = $art; Bytes = $bytes; Dateien = $dateien
            Messbar = $messbar; Laeuft = ($laeuft -join ', '); Sperre = $sperre; Fehler = (@($fehler) -join '; ')
            Mindestens = ($messbar -and $teilweise); JeLaufwerk = $jeLaufwerk
        }
    }
}

function Read-SettingsRegistry {
    $f = '<fehlt>'
    $s = [ordered]@{}
    $s['Schnellstart (HiberbootEnabled)']       = Read-RegValue 'HKLM' 'SYSTEM\CurrentControlSet\Control\Session Manager\Power' 'HiberbootEnabled' $f
    $s['Ruhezustand (HibernateEnabled)']        = Read-RegValue 'HKLM' 'SYSTEM\CurrentControlSet\Control\Power' 'HibernateEnabled' $f
    $s['GameDVR Aufnahme (AppCaptureEnabled)']  = Read-RegValue 'HKCU' 'Software\Microsoft\Windows\CurrentVersion\GameDVR' 'AppCaptureEnabled' $f
    $s['GameDVR (GameDVR_Enabled)']             = Read-RegValue 'HKCU' 'System\GameConfigStore' 'GameDVR_Enabled' $f
    $s['Dynamic Lighting (AmbientLighting)']    = Read-RegValue 'HKCU' 'Software\Microsoft\Windows\CurrentVersion\Lighting' 'AmbientLightingEnabled' $f
    $s['Dynamic Lighting (Vordergrund-App)']    = Read-RegValue 'HKCU' 'Software\Microsoft\Windows\CurrentVersion\Lighting' 'ControlledByForegroundApp' $f
    $s['Spielmodus (AutoGameModeEnabled)']      = Read-RegValue 'HKCU' 'Software\Microsoft\GameBar' 'AutoGameModeEnabled' $f
    $s['GPU-Planung HAGS (HwSchMode)']          = Read-RegValue 'HKLM' 'SYSTEM\CurrentControlSet\Control\GraphicsDrivers' 'HwSchMode' $f
    $s['Secure Boot (UEFISecureBootEnabled)']   = Read-RegValue 'HKLM' 'SYSTEM\CurrentControlSet\Control\SecureBoot\State' 'UEFISecureBootEnabled' $f
    $s['Treiber-Sperrliste (VulnerableDriver)'] = Read-RegValue 'HKLM' 'SYSTEM\CurrentControlSet\Control\CI\Config' 'VulnerableDriverBlocklistEnable' $f
    $s['Benutzerkontensteuerung (EnableLUA)']   = Read-RegValue 'HKLM' 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' 'EnableLUA' $f
    $s['UAC-Nachfrage Admins (ConsentPrompt)']  = Read-RegValue 'HKLM' 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' 'ConsentPromptBehaviorAdmin' $f
    $s['UAC sicherer Desktop (PromptOnSecure)'] = Read-RegValue 'HKLM' 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' 'PromptOnSecureDesktop' $f
    $s['Absturzabbild-Typ (CrashDumpEnabled)']  = Read-RegValue 'HKLM' 'SYSTEM\CurrentControlSet\Control\CrashControl' 'CrashDumpEnabled' $f
    $s['DirectX global']                        = Read-RegValue 'HKCU' 'Software\Microsoft\DirectX\UserGpuPreferences' 'DirectXUserGlobalSettings' $f
    $s['Mausspur (MouseTrails)']                = Read-RegValue 'HKCU' 'Control Panel\Mouse' 'MouseTrails' $f
    return $s
}

function Read-PendingReboot {
    $gruende = New-Object System.Collections.Generic.List[string]
    if (Test-RegKey 'HKLM' 'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $gruende.Add('Windows-Komponenten') }
    if (Test-RegKey 'HKLM' 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $gruende.Add('Windows Update') }
    $umbenennen = Read-RegValue 'HKLM' 'SYSTEM\CurrentControlSet\Control\Session Manager' 'PendingFileRenameOperations' $null
    [pscustomobject]@{ Gruende = @($gruende); DateiOperationen = ($null -ne $umbenennen) }
}

# ============================================================================
#  6. PRUEFUNGEN - werten nur noch zwischengespeicherte Daten aus
# ============================================================================

function Test-Stabilitaet {
    $B = 'Stabilitaet'
    $sys = @(Get-Data 'EventsSystem')
    $app = @(Get-Data 'EventsApplication')
    # Nicht lesbares Protokoll: ehrlich melden und nichts behaupten (sonst hiesse leer = „Keine Abstürze“)
    $sysFehlt = Test-QuelleFehlt 'EventsSystem' 'Ereignisse' $B 'Ereignisprotokoll System'
    $appFehlt = Test-QuelleFehlt 'EventsApplication' 'Ereignisse' $B 'Ereignisprotokoll Anwendung'
    if ($sysFehlt -or $appFehlt) { return }
    $build = [string](Get-Data 'WinVersion')['CurrentBuild']
    $grenze = (Get-Date).AddDays(-$Tage)

    # --- Harte Abstuerze / Resets ---
    $kp = @($sys | Where-Object { $_.ProviderName -eq 'Microsoft-Windows-Kernel-Power' -and $_.Id -eq 41 })
    if ($kp.Count -eq 0) {
        Add-Finding 'Absturz' $B 'Harte Abstürze / Resets' 'OK' "Keine in den letzten $Tage Tagen"
    } else {
        $zeilen = foreach ($e in $kp) {
            $d = Get-EventDataMap $e
            $bc = [int64]0
            [void][int64]::TryParse([string]$d['BugcheckCode'], [ref]$bc)
            $knopf = if (@('', '0') -notcontains [string]$d['PowerButtonTimestamp']) { 'gedrückt gehalten' } else { '' }
            [pscustomobject]@{ Zeit = Format-Datum $e.TimeCreated; Bugcheck = ('0x{0:X}' -f $bc); Einschaltknopf = $knopf }
        }
        $mitKnopf = @($zeilen | Where-Object { $_.Einschaltknopf }).Count
        $st = if ($kp.Count -ge 3) { 'Kritisch' } else { 'Warnung' }
        $ursache = 'Bugcheck 0x0 heißt: kein Bluescreen - der PC ist schlagartig ausgegangen oder eingefroren. Häufige Ursachen: Stromversorgung (Netzteil, Kabel), instabile Übertaktung von Grafikkarte oder CPU, ein hängender Treiber.'
        if ($mitKnopf) { $ursache += " $mitKnopf-mal wurde der Einschaltknopf gedrückt gehalten - das war ein manueller Reset nach einem Einfrieren." }
        $massnahme = '1. Prüfen, was zum Zeitpunkt lief (Spiel? Leerlauf?). 2. Übertaktungen testweise zurücknehmen. 3. HWiNFO-Protokoll mitlaufen lassen, damit beim nächsten Mal Spannungen und Temperaturen vorliegen. 4. Bei Häufung: Netzteil und getrennte PCIe-Stromkabel prüfen.'
        Add-Finding 'Absturz' $B 'Harte Abstürze / Resets' $st ('{0}x in {1} Tagen, zuletzt {2}' -f $kp.Count, $Tage, (Format-Datum $kp[0].TimeCreated)) -Ursache $ursache -Massnahme $massnahme -Empfehlung $(if ($kp.Count -ge 3) { 'Beheben' } else { 'Beobachten' }) -Hinweis 'Die Zeitstempel können bis ~30 Minuten hinterherhinken.' -Details @($zeilen) -Zuletzt (Get-Neuestes @($kp.TimeCreated))
    }

    # --- Bluescreens ---
    $bs = @($sys | Where-Object { $_.ProviderName -eq 'Microsoft-Windows-WER-SystemErrorReporting' -and $_.Id -eq 1001 })
    if ($bs.Count -eq 0) {
        Add-Finding 'Bluescreen' $B 'Bluescreens' 'OK' 'Keine'
    } else {
        $zeilen = foreach ($e in $bs) {
            $code = if ((Get-EventText $e) -match '(0x[0-9a-fA-F]{8})') { $Matches[1] } else { '?' }
            [pscustomobject]@{ Zeit = Format-Datum $e.TimeCreated; Stopcode = $code }
        }
        $codes = @($zeilen.Stopcode | Select-Object -Unique)
        $ursache = if ($codes.Count -gt 1) { 'Wechselnde Stopcodes deuten auf instabilen Arbeitsspeicher oder eine zu scharfe Übertaktung (CPU, RAM, Grafikkarte).' } else { "Immer derselbe Stopcode ($($codes[0])) deutet auf einen bestimmten Treiber oder ein bestimmtes Bauteil." }
        Add-Finding 'Bluescreen' $B 'Bluescreens' 'Kritisch' "$($bs.Count) Bluescreen(s) mit Speicherabbild" -Ursache $ursache -Massnahme '1. Übertaktungen zurücknehmen (DOCP bleibt an). 2. Bei wechselnden Codes Arbeitsspeicher testen (Windows-Speicherdiagnose). 3. Das Absturzabbild (siehe „Absturzabbilder“) auswerten lassen.' -Empfehlung 'Beheben' -Loesungen @(New-SeitenLoesung -Titel 'Windows-Speicherdiagnose öffnen' -Seite 'mdsched') -Details @($zeilen) -Zuletzt (Get-Neuestes @($bs.TimeCreated))
    }

    # --- Hardwarefehler (WHEA) ---
    $whea = @($sys | Where-Object { $_.ProviderName -eq 'Microsoft-Windows-WHEA-Logger' })
    if ($whea.Count -eq 0) {
        Add-Finding 'WHEA' $B 'Hardwarefehler (WHEA)' 'OK' 'Keine'
    } else {
        $schwer = @($whea | Where-Object { $_.Level -le 2 })
        $zeilen = $whea | Group-Object Id | ForEach-Object { [pscustomobject]@{ Ereignis = "WHEA $($_.Name)"; Anzahl = $_.Count; Zuletzt = Format-Datum $_.Group[0].TimeCreated; Art = $(if ($_.Group[0].Level -le 2) { 'schwer' } else { 'korrigiert' }) } }
        $st = if ($schwer.Count -gt 0) { 'Kritisch' } elseif ($whea.Count -ge 5) { 'Warnung' } else { 'Info' }
        $ursache = 'Die Hardware selbst meldet Fehler.' + $(if ($Script:IstEigenerPc) { ' WHEA 18 (schwer) war im September die Absturz-Signatur dieses PCs: ein Cache-Fehler im Prozessor.' } else { ' WHEA 18 (schwer) ist meist ein Cache-Fehler im Prozessor.' }) + ' Typische Auslöser: CPU-Übertaktung/PBO, zu wenig Spannung, Netzteil, seltener ein defektes Bauteil.'
        Add-Finding 'WHEA' $B 'Hardwarefehler (WHEA)' $st "$($whea.Count) Einträge, davon $($schwer.Count) schwer, zuletzt $(Format-Datum $whea[0].TimeCreated)" -Ursache $ursache -Massnahme '1. CPU-Übertaktung/PBO testweise abschalten (im BIOS „PBO Disabled“, DOCP an lassen). 2. Bei weiteren schweren Fehlern Netzteil prüfen. 3. Korrigierte Fehler beobachten - nur Häufungen sind kritisch.' -Empfehlung $(if ($schwer.Count) { 'Beheben' } else { 'Beobachten' }) -Details @($zeilen) -Zuletzt (Get-Neuestes @($whea.TimeCreated))
    }

    # --- Muster: Hardwarefehler direkt nach dem Einschalten (19.09. und 25.10. so passiert) ---
    # Ausschalten (1074 „Ausschalten“ bzw. 6006), Start (6005) kurz darauf, schwerer WHEA-Eintrag in den ersten Minuten danach
    $starts = @($sys | Where-Object { $_.ProviderName -eq 'EventLog' -and $_.Id -eq 6005 } | Sort-Object TimeCreated)
    $ausAus = @($sys | Where-Object { ($_.Id -eq 1074 -and (Get-EventText $_) -match '(?i)ausschalt|power ?off') -or ($_.ProviderName -eq 'EventLog' -and $_.Id -eq 6006) } | Sort-Object TimeCreated)
    $treffer = @(foreach ($w in @($whea | Where-Object { $_.Level -le 2 })) {
        $start = @($starts | Where-Object { $_.TimeCreated -le $w.TimeCreated -and ($w.TimeCreated - $_.TimeCreated).TotalMinutes -le 5 }) | Select-Object -Last 1
        if (-not $start) { continue }
        $aus = @($ausAus | Where-Object { $_.TimeCreated -le $start.TimeCreated }) | Select-Object -Last 1
        [pscustomobject]@{
            Ausgeschaltet = $(if ($aus) { Format-Datum $aus.TimeCreated } else { '-' })
            Pause = $(if ($aus) { '{0:N0} Min.' -f ($start.TimeCreated - $aus.TimeCreated).TotalMinutes } else { '-' })
            Gestartet = Format-Datum $start.TimeCreated
            Hardwarefehler = "WHEA $($w.Id), $('{0:N0}' -f ($w.TimeCreated - $start.TimeCreated).TotalSeconds) s nach dem Start"
            Zeit = $w.TimeCreated
        }
    })
    # Get-WinEvent liefert die neuesten zuerst - aufsteigend sortieren, damit $treffer[-1] wirklich der letzte ist (Archivgrenze!)
    $treffer = @($treffer | Sort-Object Zeit)
    if ($treffer.Count) {
        # PC 1 kennt die Vorgeschichte (Ryzen 3000 auf B550, BIOS bleibt) - andere PCs bekommen die allgemeine Fassung
        $ursache = if ($Script:IstEigenerPc) { 'Ein schwerer WHEA-Fehler kurz nach dem Kaltstart deutet auf die Stromversorgung im Leerlauf: Ryzen 3000 auf B550 kann beim Hochfahren in sehr sparsamen Zuständen kurz zu wenig Spannung bekommen. Übertaktung/PBO und ein älteres Netzteil verstärken das. Mit dem Rest des Betriebs hat das meist nichts zu tun.' } else { 'Ein schwerer WHEA-Fehler kurz nach dem Kaltstart deutet meist auf die Stromversorgung beim Hochfahren: Der Prozessor bekommt in sehr sparsamen Leerlaufzuständen kurz zu wenig Spannung. Übertaktung und ein älteres Netzteil verstärken das.' }
        $massnahme = if ($Script:IstEigenerPc) { '1. Im BIOS „Power Supply Idle Control“ auf „Typical Current Idle“ stellen (Advanced > AMD CBS). 2. Tritt es danach wieder auf: PBO testweise abschalten (DOCP bleibt an). 3. Kein BIOS-Update dafür - das würde Secure Boot zurücksetzen.' } else { '1. Übertaktung (CPU, Arbeitsspeicher-Profil) testweise zurücknehmen. 2. Bei AMD-Prozessoren im BIOS „Power Supply Idle Control“ auf „Typical Current Idle“ stellen. 3. Hilft beides nicht: Netzteil prüfen lassen.' }
        Add-Finding 'WheaStart' $B 'Hardwarefehler beim Einschalten' 'Warnung' "$($treffer.Count)x schwerer Hardwarefehler in den ersten Minuten nach dem Start, zuletzt $($treffer[-1].Gestartet)" -Ursache $ursache -Massnahme $massnahme -Empfehlung 'Beheben' -Details @($treffer | Select-Object Ausgeschaltet, Pause, Gestartet, Hardwarefehler) -Zuletzt $treffer[-1].Zeit
    } elseif ($whea.Count) {
        Add-Finding 'WheaStart' $B 'Hardwarefehler beim Einschalten' 'OK' 'Kein Hardwarefehler direkt nach einem Start'
    }

    # --- GPU- und Kernel-Livedumps ---
    $lke = @(foreach ($e in $app) {
        if ($e.ProviderName -ne 'Windows Error Reporting' -or $e.Id -ne 1001) { continue }
        $werte = @($e.Properties | ForEach-Object { [string]$_.Value })
        $i = [Array]::IndexOf($werte, 'LiveKernelEvent')
        if ($i -lt 0) { continue }
        $p1 = if ($werte.Count -gt $i + 3) { $werte[$i + 3] } else { '' }
        $p6 = if ($werte.Count -gt $i + 8) { $werte[$i + 8] } else { '' }
        if ($p6 -notmatch '^\d+_\d+_\d+') {
            $txt = Get-EventText $e
            if ($txt -match 'P1:\s*(\S+)') { $p1 = $Matches[1] }
            if ($txt -match 'P6:\s*(\S+)') { $p6 = $Matches[1] }
        }
        [pscustomobject]@{ Zeit = $e.TimeCreated; Code = $p1; Build = $p6 }
    })
    $wer = Get-Data 'WerFolders'
    $kernelNeu = @(@($wer.Warteschlange) + @($wer.Archiv) | Where-Object { $_ -and $_.Name -like 'Kernel_*' -and $_.Erstellt -gt $grenze })
    $aktuell = @($lke | Where-Object { $_.Build -like "*_$build" })
    $alt = @($lke | Where-Object { $_.Build -and $_.Build -notlike "*_$build" })
    if ($kernelNeu.Count -eq 0 -and $aktuell.Count -eq 0) {
        $txt = 'Keine neuen'
        if ($alt.Count -gt 0) { $txt += " ($($alt.Count) Meldungen alter Berichte aus früheren Windows-Versionen - nur Log-Rauschen)" }
        Add-Finding 'LiveKernel' $B 'Grafik-/Kernel-Hänger' 'OK' $txt
    } else {
        $zeilen = $kernelNeu | ForEach-Object {
            $code = ($_.Name -split '_')[1]
            [pscustomobject]@{ Erstellt = Format-Datum $_.Erstellt; Code = $code; Bedeutung = $(if ($Script:LiveKernelCodes.ContainsKey($code)) { $Script:LiveKernelCodes[$code] } else { 'Diagnose-Livedump' }) }
        }
        $gpu = @($kernelNeu | Where-Object { $_.Name -match '^Kernel_(141|117|124)_' })
        $st = if ($gpu.Count -ge 3) { 'Warnung' } else { 'Info' }
        $ursache = 'Code 141/117 heißt: die Grafikkarte hat nicht rechtzeitig geantwortet und wurde zurückgesetzt. Häufige Ursachen: GPU-Übertaktung/Undervolting, Treiberfehler, Lastspitzen bei der Stromversorgung. Die Codes a1000001/a2000002 sind Begleit-Dumps desselben Ereignisses.'
        $massnahme = if ($gpu.Count -ge 3) { '1. GPU-Übertaktung in Adrenalin auf Standard zurücksetzen und beobachten. 2. Hilft das, das Profil schrittweise wieder anheben. 3. Treiber-Version prüfen (bei frischem Update ggf. Vorversion).' } else { 'Einzelne Hänger beim Spielen oder beim Starten/Beenden eines Spiels sind kein Grund zum Handeln. Nur bei Häufung (3 oder mehr im Zeitraum) die Übertaktung zurücknehmen.' }
        Add-Finding 'LiveKernel' $B 'Grafik-/Kernel-Hänger' $st "$($kernelNeu.Count) neue Berichte im Zeitraum, davon $($gpu.Count) Grafik-/Hardware-Hänger" -Ursache $ursache -Massnahme $massnahme -Empfehlung $(if ($gpu.Count -ge 3) { 'Beheben' } else { 'Beobachten' }) -Details @($zeilen) -Zuletzt (Get-Neuestes (@($kernelNeu.Erstellt) + @($aktuell.Zeit)))
    }

    # --- Zuverlaessigkeitsindex ---
    $rel = @(Get-Data 'Reliability' | Where-Object { $_.TimeGenerated } | Sort-Object TimeGenerated)
    if ($rel.Count -gt 0) {
        $letzter = [Math]::Round([double]$rel[-1].SystemStabilityIndex, 2)
        $fenster = @($rel | Where-Object { $_.TimeGenerated -gt $grenze })
        $tief = if ($fenster.Count) { [Math]::Round([double]($fenster | Measure-Object -Property SystemStabilityIndex -Minimum).Minimum, 2) } else { $letzter }
        # Der Index bestimmt bewusst KEINEN Status (geprueft 28.10. an den Daten von PC 1): Windows zieht fuer einen
        # Programmabsturz oder -haenger (-1,6 bei ApplicationFrameHost, -1,3 bei einem Spiel) fast so viel ab wie fuer einen
        # harten Reset (-1,0 bis -1,7), fehlgeschlagene Store-Updates zaehlen kaum, und nach dem Ereignis bleibt der Wert
        # ~10 Tage gedrueckt. Die Ursachen selbst sind eigene Befunde mit passender Schwere - der Index waere doppelt gezaehlt
        # und bliebe auch nach dem Beheben/Archivieren noch rot. Er dient nur als Verlauf, jeder Einbruch mit seiner Ursache.
        $ausfaelle = @(Get-Data 'ZuverlEreignisse')
        $einbrueche = for ($i = 1; $i -lt $rel.Count; $i++) {
            $delta = [double]$rel[$i].SystemStabilityIndex - [double]$rel[$i - 1].SystemStabilityIndex
            if ($delta -gt -0.3) { continue }
            $bis = [datetime]$rel[$i].TimeGenerated
            # Windows stempelt den Messpunkt mit dem Beginn der Stunde, in der der Ausfall lag
            $grund = @($ausfaelle | Where-Object { $_.Zeit -lt $bis.AddHours(1) -and $_.Zeit -gt $bis.AddHours(-23) } | Sort-Object { [Math]::Abs(($_.Zeit - $bis.AddMinutes(30)).TotalMinutes) } | Select-Object -First 1 | ForEach-Object { $_.Text })
            [pscustomobject]@{ Zeit = (Format-Datum $bis); Einbruch = ('{0:N2}' -f $delta); Danach = ('{0:N2}' -f [double]$rel[$i].SystemStabilityIndex); Ursache = $(if ($grund.Count) { $grund -join ' + ' } else { 'kein Ausfall zugeordnet' }) }
        }
        $txt = "Aktuell $letzter von 10, Tiefstwert im Zeitraum $tief"
        if (@($einbrueche).Count) { $txt += " - $(@($einbrueche).Count) Einbruch/Einbrüche, zuletzt $(@($einbrueche)[-1].Zeit)" }
        Add-Finding 'Zuverlaessigkeit' $B 'Zuverlässigkeitsindex (1-10)' 'OK' $txt -Massnahme 'Nur Verlauf. Windows senkt den Wert bei Programmabstürzen, Hängern, Bluescreens und unerwarteten Abschaltungen und lässt ihn danach jeden Tag um etwa 0,3 steigen. Die Ursachen stehen als eigene Befunde mit eigener Schwere in dieser Gruppe - der Index selbst ist deshalb kein Grund zum Handeln.' -Details @($einbrueche)
    } else {
        Add-Finding 'Zuverlaessigkeit' $B 'Zuverlässigkeitsindex (1-10)' 'Info' 'nicht verfügbar' -Ursache 'Windows hat keinen Zuverlässigkeitsverlauf geliefert.'
    }

    # --- Programmabstuerze ---
    $abst = @(foreach ($e in $app) {
        if ($e.ProviderName -ne 'Application Error' -or $e.Id -ne 1000) { continue }
        $p = $e.Properties
        [pscustomobject]@{ Zeit = $e.TimeCreated; Programm = [string]$p[0].Value; Modul = [string]$p[3].Value; Code = [string]$p[6].Value }
    })
    if ($abst.Count -eq 0) {
        Add-Finding 'AppCrash' $B 'Programmabstürze' 'OK' 'Keine'
    } else {
        $gruppen = @($abst | Group-Object Programm | Sort-Object Count -Descending)
        $zeilen = $gruppen | ForEach-Object {
            [pscustomobject]@{ Programm = $_.Name; Anzahl = $_.Count; Module = (@($_.Group.Modul | Select-Object -Unique) -join ', '); Codes = (@($_.Group.Code | Select-Object -Unique) -join ', '); Zuletzt = Format-Datum $_.Group[0].Zeit }
        }
        $wiederholt = @($gruppen | Where-Object { $_.Count -ge 3 })
        $treiberModul = @($abst | Where-Object { $_.Modul -match '(?i)^(amd|ati|nv|ig|atk|logi)' })
        $ursache = 'c0000005 bedeutet: das Programm hat auf Speicher zugegriffen, der ihm nicht gehört - typisch für Programmfehler oder instabile Übertaktung.'
        if ($treiberModul.Count) { $ursache += ' Mindestens ein Absturz passierte in einer Treiberdatei (z. B. amdxx64.dll = AMD-Grafiktreiber) - das deutet eher auf den Treiber als auf das Programm.' }
        $massnahme = if ($wiederholt.Count) { "Das Programm $(@($wiederholt.Name) -join ', ') stürzt wiederholt ab: auf aktuelle Version prüfen, ggf. neu installieren; bei Spielen die Spieldateien prüfen lassen (Steam: Eigenschaften > Installierte Dateien > Überprüfen)." } else { 'Einzelne Abstürze sind normal. Erst wenn dasselbe Programm wiederholt abstürzt, lohnt sich ein genauerer Blick.' }
        Add-Finding 'AppCrash' $B 'Programmabstürze' $(if ($wiederholt.Count) { 'Warnung' } else { 'Info' }) "$($abst.Count) Abstürze in $($gruppen.Count) Programm(en)$(if ($wiederholt.Count) { ', wiederholt: ' + (@($wiederholt.Name) -join ', ') })" -Ursache $ursache -Massnahme $massnahme -Empfehlung $(if ($wiederholt.Count) { 'Beheben' } else { 'Beobachten' }) -Details @($zeilen) -Zuletzt (Get-Neuestes @($abst.Zeit))
    }

    # --- Programme, die nicht mehr reagierten ---
    $haenger = @($app | Where-Object { $_.ProviderName -eq 'Application Hang' -and $_.Id -eq 1002 })
    if ($haenger.Count -gt 0) {
        $zeilen = $haenger | Group-Object { [string]$_.Properties[0].Value } | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ Programm = $_.Name; Anzahl = $_.Count; Zuletzt = Format-Datum $_.Group[0].TimeCreated } }
        $oft = @($zeilen | Where-Object { $_.Anzahl -ge 3 })
        Add-Finding 'AppHang' $B 'Programme hingen' $(if ($oft.Count) { 'Warnung' } else { 'Info' }) "$($haenger.Count) Hänger" -Ursache 'Das Programm hat einige Sekunden lang nicht auf Windows reagiert, zum Beispiel beim Laden oder beim Beenden eines Spiels.' -Massnahme $(if ($oft.Count) { 'Wiederholt hängende Programme aktualisieren bzw. neu installieren.' } else { 'Einzelne Hänger sind unkritisch.' }) -Empfehlung $(if ($oft.Count) { 'Beheben' } else { 'Beobachten' }) -Details @($zeilen) -Zuletzt (Get-Neuestes @($haenger.TimeCreated))
    } else {
        Add-Finding 'AppHang' $B 'Programme hingen' 'OK' 'Keine'
    }

    # --- Dienste ---
    $ids = @(7000, 7001, 7009, 7011, 7022, 7023, 7024, 7031, 7034, 7043)
    $scm = @($sys | Where-Object { $_.ProviderName -eq 'Service Control Manager' -and $ids -contains $_.Id -and -not (Get-NoiseReason $_) })
    if ($scm.Count -eq 0) {
        Add-Finding 'Dienstfehler' $B 'Dienst-Probleme' 'OK' 'Keine'
    } else {
        # Bei 7009/7011 steht im ersten Feld die Wartezeit, der Dienstname erst im zweiten
        $zeilen = $scm | Group-Object { $feld = if ($_.Id -in @(7009, 7011)) { 1 } else { 0 }; "$($_.Id)|$([string]$_.Properties[$feld].Value)" } | ForEach-Object {
            $id = [int]($_.Name -split '\|')[0]
            [pscustomobject]@{ Dienst = ($_.Name -split '\|', 2)[1]; Ereignis = "$id - $($Script:DienstEreignisse[$id])"; Anzahl = $_.Count; Zuletzt = Format-Datum $_.Group[0].TimeCreated }
        }
        $absturz = @($zeilen | Where-Object { $_.Ereignis -match '^70(31|34)' -and $_.Anzahl -ge 2 })
        $dienste = @($zeilen.Dienst | Select-Object -Unique) -join ', '
        $ursache = 'Start-Zeitüberschreitungen entstehen oft direkt nach dem Hochfahren, wenn viele Dienste gleichzeitig starten (z. B. der Steam Client Service, der bei Bedarf erneut startet). Wiederholte Abstürze desselben Dienstes sind dagegen meist das Symptom eines Konflikts' + $(if ($Script:IstEigenerPc) { ' - so kam es im September beim LampArray-Dienst durch Windows Dynamic Lighting.' } else { ', etwa zweier Programme, die dasselbe Gerät steuern wollen.' })
        $massnahme = if ($absturz.Count) { "Bei $(@($absturz.Dienst) -join ', '): die Ursache des Konflikts suchen, nicht den Dienst abschalten. Zugehörige Software aktualisieren." } else { 'Einzelne Start-Probleme sind unkritisch, solange die zugehörige Funktion (z. B. Steam) normal läuft.' }
        Add-Finding 'Dienstfehler' $B 'Dienst-Probleme' $(if ($absturz.Count) { 'Warnung' } else { 'Info' }) "$($scm.Count) Ereignisse ($dienste)" -Ursache $ursache -Massnahme $massnahme -Empfehlung $(if ($absturz.Count) { 'Beheben' } else { 'Beobachten' }) -Details @($zeilen) -Zuletzt (Get-Neuestes @($scm.TimeCreated))
    }

    # --- Boot-Treiber, die nicht laden ---
    $boot = @($sys | Where-Object { $_.ProviderName -eq 'Service Control Manager' -and $_.Id -eq 7026 })
    $treiber = @{}
    foreach ($e in $boot) {
        foreach ($n in @(([string]$e.Properties[0].Value) -split '\s+' | Where-Object { $_ })) { $treiber[$n] = 1 + [int]$treiber[$n] }
    }
    $unbekannt = @($treiber.Keys | Where-Object { $_ -notmatch '^(dam|cpuz\d+)$' })
    if ($treiber.Count -eq 0) {
        Add-Finding 'BootTreiber' $B 'Boot-Treiber' 'OK' 'Alle Boot-Treiber laden'
    } elseif ($unbekannt.Count -eq 0) {
        Add-Finding 'BootTreiber' $B 'Boot-Treiber' 'OK' ('Nur bekannte harmlose Meldungen: ' + (@($treiber.Keys) -join ', ')) -Hinweis '„dam“ ist der Modern-Standby-Treiber, den Desktop-PCs nicht unterstützen - die Meldung erscheint bei jedem Start.'
    } else {
        Add-Finding 'BootTreiber' $B 'Boot-Treiber' 'Warnung' ('Laden fehlgeschlagen: ' + ($unbekannt -join ', ')) -Ursache 'Meist gehört der Treiber zu Software, die deinstalliert wurde oder deren Datei fehlt.' -Massnahme 'Unter „Verwaiste Treiber“ nachsehen - dort lassen sich Treiber ohne Datei abschalten. Gehört der Treiber zu installierter Software, diese neu installieren.' -Empfehlung 'Prüfen'
    }

    # --- Absturzabbilder ---
    $d = Get-Data 'Dumps'
    if ($d) {
        $alle = @()
        if ($d.Speicherabbild) { $alle += [pscustomobject]@{ Art = 'Kernel (komplett)'; Datei = $d.Speicherabbild.Name; Pfad = $d.Speicherabbild.FullName; Zeit = $d.Speicherabbild.LastWriteTime; MB = [Math]::Round($d.Speicherabbild.Length / 1MB) } }
        foreach ($f in $d.Minidumps) { $alle += [pscustomobject]@{ Art = 'Kernel (Minidump)'; Datei = $f.Name; Pfad = $f.FullName; Zeit = $f.LastWriteTime; MB = [Math]::Round($f.Length / 1MB, 1) } }
        foreach ($f in $d.LiveKernel) { $alle += [pscustomobject]@{ Art = 'Livedump'; Datei = $f.Name; Pfad = $f.FullName; Zeit = $f.LastWriteTime; MB = [Math]::Round($f.Length / 1MB, 1) } }
        foreach ($f in $d.Programme) { $alle += [pscustomobject]@{ Art = 'Programm'; Datei = $f.Name; Pfad = $f.FullName; Zeit = $f.LastWriteTime; MB = [Math]::Round($f.Length / 1MB, 1) } }
        if ($alle.Count -eq 0) {
            Add-Finding 'Dumps' $B 'Absturzabbilder' 'OK' 'Keine vorhanden'
        } else {
            $neu = @($alle | Where-Object { $_.Zeit -gt $grenze -and $_.Art -ne 'Programm' })
            $loeschbar = @($alle | Where-Object { -not ($_.Zeit -gt $grenze -and $_.Art -ne 'Programm') })
            $mb = ($alle | Measure-Object -Property MB -Sum).Sum
            $zeilen = $alle | Sort-Object Zeit -Descending | ForEach-Object { [pscustomobject]@{ Art = $_.Art; Datei = $_.Datei; Zeit = Format-Datum $_.Zeit; MB = $_.MB } }
            $loesungen = @()
            $kat = @(Get-Data 'Reinigung') | Where-Object { $_.Id -eq 'Absturzabbilder' } | Select-Object -First 1
            if ($loeschbar.Count -and $kat -and $kat.Dateien -gt 0) { $loesungen += New-Loesung -Id 'Reinigen' -Titel ('{0} Absturzabbilder älter als 14 Tage löschen ({1})' -f $kat.Dateien, (Format-Groesse $kat.Bytes)) -Umkehrbar $false -Parameter @{ Kategorie = 'Absturzabbilder' } -Beschreibung 'Neuere Abbilder bleiben für die Analyse erhalten.' }
            $ursache = if ($neu.Count) { "Es gab im Zeitraum $($neu.Count) neue Kernel-Abbilder - sie gehören zu Bluescreens oder Grafik-Hängern." } else { 'Die Abbilder stammen von älteren Abstürzen (Spiele, Programme oder Bluescreens).' }
            $massnahme = if ($neu.Count) { 'Neue Kernel-Abbilder aufheben, bis die Ursache geklärt ist. Alte Abbilder können gelöscht werden.' } else { 'Alle Abbilder sind älter als der Prüfzeitraum und können gelöscht werden, sofern die zugehörigen Abstürze geklärt sind.' }
            Add-Finding 'Dumps' $B 'Absturzabbilder' $(if ($neu.Count) { 'Warnung' } else { 'Info' }) ("{0} Dateien, zusammen {1:N0} MB{2}" -f $alle.Count, $mb, $(if ($neu.Count) { ", davon $($neu.Count) neue Kernel-Abbilder" } else { ' (alle älter als der Zeitraum)' })) -Ursache $ursache -Massnahme $massnahme -Empfehlung $(if ($neu.Count) { 'Prüfen' } else { 'Entfernen' }) -Details @($zeilen) -Loesungen $loesungen
        }
    }

    # --- Fehlerberichts-Warteschlange ---
    if ($wer) {
        $q = @($wer.Warteschlange)
        if ($q.Count -gt 20) {
            Add-Finding 'WerQueue' $B 'Fehlerberichts-Warteschlange' 'Info' "$($q.Count) Berichte warten auf das Senden" -Ursache 'Die zugehörigen Absturzabbilder sind meist schon gelöscht, deshalb scheitert das Senden - und Windows versucht es täglich erneut. Das erzeugt Hunderte Einträge im Ereignisprotokoll.' -Massnahme 'Fehlerberichte löschen (wartende und gesendete). Die Berichte haben ohne ihre Abbilder keinen Wert mehr.' -Empfehlung 'Entfernen' -Loesungen @(New-Loesung -Id 'Reinigen' -Titel "Fehlerberichte löschen ($($q.Count) wartende)" -Umkehrbar $false -Parameter @{ Kategorie = 'Fehlerberichte' })
        } else {
            Add-Finding 'WerQueue' $B 'Fehlerberichts-Warteschlange' 'OK' "$($q.Count) Berichte"
        }
    }

    # --- Alle uebrigen Fehler im Ereignisprotokoll ---
    $behandelt = @('Microsoft-Windows-Kernel-Power', 'EventLog', 'Microsoft-Windows-WER-SystemErrorReporting', 'Microsoft-Windows-WHEA-Logger',
        'Service Control Manager', 'Microsoft-Windows-WindowsUpdateClient', 'Microsoft-Windows-Time-Service', 'Application Error',
        'Application Hang', 'Windows Error Reporting', 'User32')
    $rest = @(@($sys) + @($app) | Where-Object { $_.Level -le 2 -and $behandelt -notcontains $_.ProviderName })
    $rauschen = New-Object System.Collections.Generic.List[object]
    $echt = New-Object System.Collections.Generic.List[object]
    $echtZeiten = New-Object System.Collections.Generic.List[object]
    foreach ($g in @($rest | Group-Object { "$($_.LogName)|$($_.ProviderName)|$($_.Id)" })) {
        $erstes = $g.Group[0]
        $grund = Get-NoiseReason $erstes
        $zeile = [pscustomobject]@{ Protokoll = $erstes.LogName; Quelle = $erstes.ProviderName; Id = $erstes.Id; Anzahl = $g.Count; Zuletzt = Format-Datum $erstes.TimeCreated; Text = $(if ($grund) { $grund } else { Format-Kurz (Get-EventText $erstes) 180 }) }
        if ($grund) { $rauschen.Add($zeile) } else { $echt.Add($zeile); $echtZeiten.Add($erstes.TimeCreated) }
    }
    $warnRauschen = @($sys | Where-Object { $_.Level -eq 3 -and (Get-NoiseReason $_) })
    if ($echt.Count -eq 0) {
        Add-Finding 'Ereignisse' $B 'Weitere Fehler im Ereignisprotokoll' 'OK' 'Keine unbekannten Fehler'
    } else {
        $haeufig = @($echt | Where-Object { $_.Anzahl -ge 5 })
        $volsnap = @($echt | Where-Object { $_.Quelle -eq 'Volsnap' -and $_.Id -eq 36 })
        $ursache = 'Fehler, die keiner bekannten Ursache zugeordnet sind. Einmalige Einträge entstehen oft beim ersten Start nach einem Update oder beim Herunterfahren.'
        if ($volsnap.Count) { $ursache += ' Volsnap 36 heißt: Windows hat Wiederherstellungspunkte verworfen, weil auf dem Laufwerk zu wenig Platz war.' }
        $massnahme = if ($haeufig.Count) { 'Die häufigen Einträge (Details) ansehen - wiederkehrende Fehler derselben Quelle deuten auf ein echtes Problem.' } else { 'Nur einmalige Einträge - nichts zu tun, solange keine Funktion gestört ist.' }
        if ($volsnap.Count) { $massnahme += ' Für die Wiederherstellungspunkte: Platz auf dem Systemlaufwerk schaffen (siehe „Laufwerke“).' }
        Add-Finding 'Ereignisse' $B 'Weitere Fehler im Ereignisprotokoll' $(if ($haeufig.Count) { 'Warnung' } else { 'Info' }) "$($echt.Count) unbekannte Fehlerarten, davon $($haeufig.Count) häufig (5x oder öfter)" -Ursache $ursache -Massnahme $massnahme -Empfehlung $(if ($haeufig.Count -or $volsnap.Count) { 'Prüfen' } else { 'Beobachten' }) -Details @($echt | Sort-Object Anzahl -Descending) -Zuletzt (Get-Neuestes $echtZeiten.ToArray())
    }
    if ($rauschen.Count -gt 0 -or $warnRauschen.Count -gt 0) {
        $rauschWarn = $warnRauschen | Group-Object { "$($_.ProviderName)|$($_.Id)" } | ForEach-Object { [pscustomobject]@{ Protokoll = 'System'; Quelle = $_.Group[0].ProviderName; Id = $_.Group[0].Id; Anzahl = $_.Count; Zuletzt = Format-Datum $_.Group[0].TimeCreated; Text = (Get-NoiseReason $_.Group[0]) } }
        $alleR = @($rauschen.ToArray()) + @($rauschWarn)
        Add-Finding 'Rauschen' $B 'Bekanntes Log-Rauschen (harmlos)' 'OK' "$(($alleR | Measure-Object -Property Anzahl -Sum).Sum) Einträge, die auf diesem PC bekannt und harmlos sind" -Details @($alleR | Sort-Object Anzahl -Descending)
    }

    # --- Starts und Herunterfahren ---
    $starts = @($sys | Where-Object { $_.Id -eq 6005 -and $_.ProviderName -eq 'EventLog' }).Count
    $unerwartet = @($sys | Where-Object { $_.Id -eq 6008 -and $_.ProviderName -eq 'EventLog' }).Count
    Add-Finding 'Starts' $B 'Starts im Zeitraum' $(if ($unerwartet) { 'Info' } else { 'OK' }) "$starts Systemstarts, davon $unerwartet nach unerwartetem Herunterfahren"
}

function Test-HardwareTreiber {
    $B = 'Hardware & Treiber'

    # --- Geraete mit Problemen ---
    $pnp = @(Get-Data 'PnpDevices')
    $da = @($pnp | Where-Object { $_.Present })
    $fehler = @($da | Where-Object { $null -ne $_.ConfigManagerErrorCode -and [int]$_.ConfigManagerErrorCode -ne 0 })
    if ($fehler.Count -eq 0) {
        Add-Finding 'Geraetefehler' $B 'Geräte mit Fehler' 'OK' "Keine ($($da.Count) angeschlossene Geräte geprüft)" -Hinweis 'Geräte melden auch dann „OK“, wenn ihre Funktion tot ist - bei Problemen immer die Funktion selbst testen.'
    } else {
        $zeilen = $fehler | ForEach-Object { $c = [int]$_.ConfigManagerErrorCode; [pscustomobject]@{ Geraet = $_.FriendlyName; Klasse = $_.Class; Code = $c; Bedeutung = $(if ($Script:GeraeteFehlerCodes.ContainsKey($c)) { $Script:GeraeteFehlerCodes[$c] } else { '' }) } }
        $nurAus = @($fehler | Where-Object { [int]$_.ConfigManagerErrorCode -eq 22 })
        $st = if ($nurAus.Count -eq $fehler.Count) { 'Info' } else { 'Warnung' }
        Add-Finding 'Geraetefehler' $B 'Geräte mit Fehler' $st "$($fehler.Count) Gerät(e)$(if ($nurAus.Count) { ", davon $($nurAus.Count) bewusst deaktiviert" })" -Ursache 'Die Bedeutung des Codes steht in den Details. Code 28 = kein Treiber, 10/43 = Gerät startet nicht bzw. meldet einen Fehler, 22 = vom Benutzer deaktiviert.' -Massnahme 'Bei Code 28: Treiber des Herstellers installieren. Bei Code 10/43: Gerät neu anschließen, Treiber neu installieren. Code 22 ist gewollt und braucht nichts.' -Empfehlung $(if ($st -eq 'Info') { 'Behalten' } else { 'Beheben' }) -Details @($zeilen) -Loesungen @(New-SeitenLoesung -Titel 'Geräte-Manager öffnen' -Seite 'devmgmt')
    }

    $geister = @($pnp | Where-Object { -not $_.Present -and $_.Class -notin @('SoftwareDevice', 'Volume', 'VolumeSnapshot', 'DiskDrive', 'WPD', 'PrintQueue', 'AudioEndpoint', 'Printer') })
    if ($geister.Count -gt 0) {
        $zeilen = $geister | Group-Object Class | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ Klasse = $(if ($_.Name) { $_.Name } else { '(ohne)' }); Anzahl = $_.Count; Beispiele = (@($_.Group.FriendlyName | Select-Object -Unique -First 4) -join ' | ') } }
        Add-Finding 'Geister' $B 'Geistergeräte' 'Info' "$($geister.Count) früher angeschlossene Geräte gespeichert" -Ursache 'Normal: Windows merkt sich jedes einmal angeschlossene Gerät, damit es beim nächsten Anschluss sofort funktioniert.' -Massnahme 'Nichts zu tun. Die Einträge kosten weder Leistung noch Platz.' -Empfehlung 'Behalten' -Details @($zeilen)
    }

    # --- Treiberspeicher: tote Altpakete ---
    $speicher = @(Get-Data 'DriverStore')
    $gebunden = @{}
    foreach ($s in @(Get-Data 'SignedDrivers')) { if ($s.InfName) { $gebunden[([string]$s.InfName).ToLowerInvariant()] = "$($s.DriverVersion)" } }
    if ($speicher.Count -gt 0) {
        $neuerAlsAktiv = New-Object System.Collections.Generic.List[string]
        $tot = @(foreach ($g in @($speicher | Group-Object Original | Where-Object { $_.Count -gt 1 })) {
            $aktiv = @($g.Group | Where-Object { $gebunden.ContainsKey($_.Inf.ToLowerInvariant()) })
            if ($aktiv.Count -eq 0) { continue }
            $aktivMax = @($aktiv | ForEach-Object { ConvertTo-TreiberVersion $_.Version } | Sort-Object -Descending)[0]
            foreach ($p in $g.Group) {
                if ($gebunden.ContainsKey($p.Inf.ToLowerInvariant())) { continue }
                # Ein ungenutztes Paket, das NEUER ist als das aktive, wird nicht zum Entfernen vorgeschlagen
                if ((ConvertTo-TreiberVersion $p.Version) -gt $aktivMax) { $neuerAlsAktiv.Add("$($p.Inf) ($($p.Original) $($p.Version))"); continue }
                [pscustomobject]@{ Paket = $p.Inf; Original = $p.Original; Version = $p.Version; Anbieter = $p.Anbieter; AktivStattdessen = (@($aktiv | ForEach-Object { "$($_.Inf) ($($_.Version))" }) -join ', ') }
            }
        })
        # Grafiktreiber-Hauptpakete heissen bei jeder Version anders (z. B. u0203731.inf -> u0205306.inf),
        # deshalb zusaetzlich: ungebundene Grafikpakete desselben Herstellers wie das aktive Paket
        $bildAktiv = @($speicher | Where-Object { $_.Klasse -match '(?i)^(Display|Grafik)' -and $gebunden.ContainsKey($_.Inf.ToLowerInvariant()) })
        foreach ($p in @($speicher | Where-Object { $_.Klasse -match '(?i)^(Display|Grafik)' -and -not $gebunden.ContainsKey($_.Inf.ToLowerInvariant()) })) {
            $gleich = @($bildAktiv | Where-Object { $_.Anbieter -eq $p.Anbieter })
            if ($gleich.Count -and -not @($tot | Where-Object { $_.Paket -eq $p.Inf }).Count) {
                $tot += [pscustomobject]@{ Paket = $p.Inf; Original = $p.Original; Version = $p.Version; Anbieter = $p.Anbieter; AktivStattdessen = (@($gleich | ForEach-Object { "$($_.Inf) ($($_.Version))" }) -join ', ') }
            }
        }
        $neuerHinweis = if ($neuerAlsAktiv.Count) { 'Nicht vorgeschlagen, weil neuer als das aktive Paket (Windows bevorzugt hier bewusst die ältere Version): ' + (@($neuerAlsAktiv) -join ', ') } else { '' }
        if ($tot.Count -eq 0) {
            Add-Finding 'Treiberspeicher' $B 'Treiberspeicher' 'OK' "$($speicher.Count) Treiberpakete, keine toten Altpakete" -Hinweis $neuerHinweis
        } else {
            $loesungen = @($tot | ForEach-Object { New-Loesung -Id 'TreiberpaketEntfernen' -Titel "$($_.Paket) entfernen - $($_.Original) $($_.Version)" -Umkehrbar $false -Parameter @{ Inf = $_.Paket } -Beschreibung "Aktiv ist stattdessen: $($_.AktivStattdessen)" })
            Add-Finding 'Treiberspeicher' $B 'Treiberspeicher' 'Info' "$($speicher.Count) Treiberpakete, davon $($tot.Count) tote Altpakete (durch Updates überholt)" -Ursache 'Bei Treiber-Updates (hier v. a. Grafiktreiber und Chipsatz) installiert Windows die neue Version neben der alten. Die alte wird von keinem Gerät mehr benutzt, bleibt aber im Speicher liegen.' -Massnahme 'Die alten Pakete können entfernt werden. Windows verweigert das Entfernen von selbst, falls ein Paket doch noch gebraucht wird - dann wird nichts erzwungen.' -Empfehlung 'Entfernen' -Hinweis $neuerHinweis -Details @($tot) -Loesungen $loesungen
        }
    }

    # --- Treiberversionen ---
    $klassen = @(Get-Data 'DriverClasses' | Sort-Object Klasse, Geraet, Version -Unique)
    if ($klassen.Count -gt 0) {
        $gpu = $klassen | Where-Object { $_.Klasse -eq 'Grafik' } | Select-Object -First 1
        $txt = if ($gpu) { "$($gpu.Geraet): $($gpu.Version) vom $($gpu.Datum) $($gpu.Zusatz)" } else { "$($klassen.Count) Treiber" }
        Add-Finding 'Treiberversionen' $B 'Treiberversionen' 'Info' $txt -Massnahme 'Nur zur Übersicht. Treiber-Updates nur einspielen, wenn sie ein Problem lösen oder für ein Spiel nötig sind.' -Empfehlung 'Behalten' -Details @($klassen)
    }

    $vc = @(Get-Data 'VideoController')
    foreach ($g in $vc) {
        if ($g.Status -eq 'OK') { Add-Finding 'Grafikkarte' $B "Grafikkarte $($g.Name)" 'OK' 'Status OK' }
        else { Add-Finding 'Grafikkarte' $B "Grafikkarte $($g.Name)" 'Warnung' "Status $($g.Status)" -Ursache 'Der Grafiktreiber meldet ein Problem an Windows.' -Massnahme 'PC neu starten. Bleibt es, Grafiktreiber über AMD Adrenalin neu installieren.' -Empfehlung 'Beheben' }
    }

    # --- Arbeitsspeicher ---
    $ram = @(Get-Data 'RAM')
    if ($ram.Count -gt 0) {
        $gb = [Math]::Round((($ram | Measure-Object -Property Capacity -Sum).Sum / 1GB))
        $mhz = ($ram | Measure-Object -Property ConfiguredClockSpeed -Minimum).Minimum
        $zeilen = $ram | ForEach-Object { [pscustomobject]@{ Steckplatz = $_.DeviceLocator; Hersteller = $_.Manufacturer; Modell = ([string]$_.PartNumber).Trim(); GB = [Math]::Round($_.Capacity / 1GB); MHz = $_.ConfiguredClockSpeed } }
        $probleme = @()
        # Sollwerte 0 = noch keine bekannt (neuer PC) - dann nur anzeigen
        if ($Script:Profil.RamModule -and $ram.Count -ne $Script:Profil.RamModule) { $probleme += "$($ram.Count) statt $($Script:Profil.RamModule) Riegel erkannt" }
        if ($Script:Profil.RamMhz -and $mhz -lt $Script:Profil.RamMhz) { $probleme += "läuft mit $mhz statt $($Script:Profil.RamMhz) MHz" }
        if ($probleme.Count) {
            $ursache = ''
            if ($ram.Count -ne $Script:Profil.RamModule) { $ursache += 'Ein Riegel wird nicht erkannt - fast immer sitzt er nach Arbeiten im Gehäuse nicht mehr richtig. ' }
            if ($mhz -lt $Script:Profil.RamMhz) { $ursache += 'DOCP/XMP ist im BIOS aus - das passiert nach jedem BIOS-Reset oder -Update.' }
            Add-Finding 'RAM' $B 'Arbeitsspeicher' 'Warnung' (($probleme -join ', ') + " ($gb GB)") -Ursache $ursache.Trim() -Massnahme 'Fehlender Riegel: PC ausschalten, Riegel herausnehmen und fest einrasten lassen. Zu niedriger Takt: im BIOS (Entf beim Start) unter AI Tweaker „DOCP“ aktivieren.' -Empfehlung 'Beheben' -Details @($zeilen)
        } else {
            Add-Finding 'RAM' $B 'Arbeitsspeicher' 'OK' ("$gb GB, $($ram.Count) Riegel, $mhz MHz" + $(if ($Script:Profil.RamMhz) { ' (DOCP/XMP aktiv)' } else { '' })) -Details @($zeilen)
        }
    }

    # --- Kernel-Treiber von Drittanbietern ---
    $drv = @(Get-Data 'SystemDrivers')
    $fremd = @($drv | Where-Object { $_.Status -eq 'Running' -and $_.Hersteller -and $_.Hersteller -notmatch '^Microsoft' })
    $tiefe = @($fremd | Where-Object { $_.Name -match $Script:TiefenTreiberMuster -or [string]$_.Pfad -match $Script:TiefenTreiberMuster })
    $neueTiefe = @($tiefe | Where-Object { $n = $_.Name; -not @($Script:Profil.BekannteTiefenTreiber | Where-Object { $n -like $_ }).Count })
    if ($tiefe.Count -gt 0) {
        $zeilen = $tiefe | ForEach-Object { [pscustomobject]@{ Treiber = $_.Name; Start = $_.Start; Hersteller = $_.Hersteller; Datei = $_.Pfad } }
        $txt = "$($tiefe.Count) aktiv: " + (@($tiefe.Name) -join ', ')
        if ($neueTiefe.Count) { $txt += ' - NEU: ' + (@($neueTiefe.Name) -join ', ') }
        $ursache = 'Sie gehören zu Armoury Crate (MSIO, AsIO3, IOMap), iCUE (CorsairLLAccess, cpuz), Ryzen Master/Adrenalin, Creative (CTIAIO) und Riot Vanguard (vgk). Mehrere davon greifen unabhängig voneinander auf dieselben Mainboard-Bausteine zu.'
        $massnahme = if ($neueTiefe.Count) { "Neu hinzugekommen: $(@($neueTiefe.Name) -join ', '). Prüfen, welches Programm ihn installiert hat und ob es gebraucht wird." } else { 'Nichts zu tun, solange der PC stabil läuft. Kommen Abstürze zurück, ist dieser Treiberstapel der erste Verdacht (testweise RGB-/Sensor-Software beenden).' }
        Add-Finding 'Tiefentreiber' $B 'Treiber mit direktem Hardwarezugriff' $(if ($neueTiefe.Count) { 'Warnung' } else { 'Info' }) $txt -Ursache $ursache -Massnahme $massnahme -Empfehlung $(if ($neueTiefe.Count) { 'Prüfen' } else { 'Beobachten' }) -Details @($zeilen)
    }
    if ($fremd.Count -gt 0) {
        $zeilen = $fremd | Sort-Object Hersteller, Name | ForEach-Object { [pscustomobject]@{ Treiber = $_.Name; Start = $_.Start; Hersteller = $_.Hersteller } }
        Add-Finding 'Fremdtreiber' $B 'Kernel-Treiber von Drittanbietern' 'Info' "$($fremd.Count) laufende Treiber" -Massnahme 'Nur zur Übersicht.' -Empfehlung 'Behalten' -Details @($zeilen)
    }

    $verwaist = @($drv | Where-Object { $_.Start -in @('Boot', 'System', 'Auto') -and $_.Vorhanden -eq $false })
    $verwaistEcht = @($verwaist | Where-Object { [string]$_.Pfad -notmatch '(?i)\\CPUID' -and $_.Name -notmatch '^cpuz' })
    if ($verwaistEcht.Count -gt 0) {
        $loesungen = @($verwaistEcht | ForEach-Object { New-Loesung -Id 'TreiberdienstDeaktivieren' -Titel "Treiber $($_.Name) abschalten" -Parameter @{ Name = $_.Name } -Beschreibung "Datei fehlt: $($_.Pfad)" })
        Add-Finding 'VerwaisteTreiber' $B 'Verwaiste Treiber' 'Warnung' "$($verwaistEcht.Count) Treiber starten automatisch, ihre Datei fehlt aber: $(@($verwaistEcht.Name) -join ', ')" -Ursache 'Die zugehörige Software wurde deinstalliert, ihr Treibereintrag ist geblieben. Windows versucht ihn bei jedem Start zu laden.' -Massnahme 'Den Treibereintrag abschalten (umkehrbar). Ganz entfernen lässt er sich später, wenn sicher ist, dass nichts fehlt.' -Empfehlung 'Entfernen' -Details @($verwaistEcht | Select-Object Name, Start, Pfad) -Loesungen $loesungen
    } else {
        $zusatz = if ($verwaist.Count) { " ($($verwaist.Count) CPU-Z-Temp-Treiber, harmlos)" } else { '' }
        Add-Finding 'VerwaisteTreiber' $B 'Verwaiste Treiber' 'OK' "Keine$zusatz"
    }
}

function Test-Laufwerke {
    $B = 'Laufwerke'
    # Papierkorb je Laufwerk aus der Reinigungsmessung (eigener Benutzer, auch Reste, die Windows nicht mehr anzeigt)
    $korb = @(Get-Data 'Reinigung') | Where-Object { $_.Art -eq 'Papierkorb' } | Select-Object -First 1

    foreach ($v in @(Get-Data 'Volumes' | Where-Object { $_.DriveLetter } | Sort-Object DriveLetter)) {
        $l = [string]$v.DriveLetter
        $frei = [double]$v.SizeRemaining / 1GB
        $gesamt = [double]$v.Size / 1GB
        $proz = if ($gesamt -gt 0) { 100 * $frei / $gesamt } else { 0 }
        $txt = '{0:N1} GB von {1:N0} GB frei ({2:N0} %), {3}{4}' -f $frei, $gesamt, $proz, $v.FileSystem, $(if ($v.FileSystemLabel) { ", Name '$($v.FileSystemLabel)'" } else { '' })
        $korbBytes = if ($korb -and $korb.JeLaufwerk -and $korb.JeLaufwerk.Contains($l)) { [int64]$korb.JeLaufwerk[$l] } else { 0 }
        if ($korbBytes -ge 100MB) { $txt += ', davon Papierkorb ' + (Format-Groesse $korbBytes) }
        $ausnahme = $Script:Profil.LaufwerkAusnahmen[$l]
        $istSystem = ($l -eq ([string]$env:SystemDrive).TrimEnd(':'))
        if ([string]$v.HealthStatus -ne 'Healthy') {
            Add-Finding 'Laufwerk' $B "Laufwerk ${l}:" 'Kritisch' "$txt - Zustand: $($v.HealthStatus)" -Ursache 'Windows meldet das Laufwerk als nicht gesund - das kann ein Dateisystemfehler oder ein Defekt sein.' -Massnahme 'Sofort wichtige Daten sichern. Danach im Gründlich-Modus die SMART-Werte des Datenträgers prüfen.' -Empfehlung 'Beheben'
        } elseif ($proz -lt $Script:Profil.MinFreiProzent -and -not $ausnahme) {
            $st = if ($proz -lt 5) { 'Kritisch' } else { 'Warnung' }
            if ($istSystem) {
                $ursache = 'Auf dem Systemlaufwerk ist zu wenig Platz. Windows braucht Puffer für Updates, die Auslagerungsdatei und Wiederherstellungspunkte - wird es eng, verwirft Windows die Wiederherstellungspunkte und Updates können scheitern.'
                $massnahme = '1. Unter Einstellungen > System > Speicher nachsehen, welche Kategorie wächst (Apps, Temporäre Dateien, Andere). 2. Spiele auf F: verschieben bzw. dort installieren. 3. Die Aufräum-Kandidaten unten bereinigen.'
            } else {
                $ursache = 'Das Laufwerk ist fast voll - meist durch Spiele und Spiele-Updates.'
                $massnahme = 'Nicht mehr gespielte Spiele deinstallieren. Spiele-Updates brauchen beim Installieren oft doppelt so viel Platz wie das Update selbst.'
            }
            $loes = @(New-SeitenLoesung -Titel 'Speicher-Einstellungen öffnen' -Seite 'ms-settings:storagesense')
            if ($korbBytes -ge 1GB) {
                $massnahme = "Zuerst den Papierkorb leeren: auf ${l}: liegen darin $(Format-Groesse $korbBytes)" + ' - oft verwaiste Reste, die der Papierkorb selbst nicht mehr anzeigt (Reiter „Reinigung“ > Papierkorb). Danach: ' + $massnahme
                $loes = @((New-Loesung -Id 'ReinigungOeffnen' -Titel 'Reinigung öffnen (Papierkorb)' -Art 'Sofort')) + $loes
            }
            Add-Finding 'Laufwerk' $B "Laufwerk ${l}:" $st $txt -Ursache $ursache -Massnahme $massnahme -Empfehlung 'Beheben' -Loesungen $loes
        } else {
            Add-Finding 'Laufwerk' $B "Laufwerk ${l}:" 'OK' $txt -Hinweis $(if ($ausnahme) { "Ausnahme laut PC-Profil: $ausnahme" } else { '' })
        }
    }

    $smart = @(Get-Data 'Smart')
    foreach ($d in @(Get-Data 'PhysicalDisks')) {
        $name = "Datenträger: $($d.Modell)"
        $probleme = New-Object System.Collections.Generic.List[string]
        $kritisch = $false
        if ($d.Zustand -ne 'Healthy') { $probleme.Add("Zustand $($d.Zustand)"); $kritisch = $true }
        if ($d.LeseFehler -gt 0) { $probleme.Add("$($d.LeseFehler) unkorrigierbare Lesefehler") }
        if ($d.SchreibFehler -gt 0) { $probleme.Add("$($d.SchreibFehler) unkorrigierbare Schreibfehler") }
        if ($d.Verschleiss -ge 80) { $probleme.Add("Verschleiß $($d.Verschleiss) %") }
        $fwHinweis = if ($d.Firmware -and $Script:Profil.FirmwareHinweise.Contains([string]$d.Firmware)) { $Script:Profil.FirmwareHinweise[[string]$d.Firmware] } else { '' }
        $details = $null
        $s = $smart | Where-Object { $_ -and $_.Seriennummer -and $_.Seriennummer -eq $d.Seriennummer } | Select-Object -First 1
        if ($s) {
            if ($s.Ausfallprognose) { $probleme.Add('SMART sagt einen Ausfall voraus'); $kritisch = $true }
            foreach ($id in @('5', '187', '197', '198', '199', '183')) {
                if ($s.Attribute.Contains($id) -and $s.Attribute[$id].Rohwert -gt 0) { $probleme.Add("$($s.Attribute[$id].Name): $($s.Attribute[$id].Rohwert)") }
            }
            $details = @($s.Attribute.Values | ForEach-Object { [pscustomobject]@{ Attribut = "$($_.Id) - $($_.Name)"; Zustand = $_.Zustand; Rohwert = $_.Rohwert } })
        }
        $info = '{0} {1}, {2} GB, Firmware {3}' -f $d.Bus, $d.Typ, $d.GB, $d.Firmware
        if ($null -ne $d.TempC -and $d.TempC -gt 0) { $info += ", $($d.TempC) °C" }
        if ($null -ne $d.Stunden -and $d.Stunden -gt 0) { $info += ", $($d.Stunden) Betriebsstunden" }
        $hinweis = if (-not $s -and -not $Script:AdminModus) { 'Fehlerzähler und SMART-Werte nur im Gründlich-Modus (mit Adminrechten).' } elseif (-not $s -and $Script:QuellenFehler.ContainsKey('Smart')) { 'SMART-Werte nicht lesbar: ' + (Format-Kurz $Script:QuellenFehler['Smart'] 100) } else { '' }
        if ($probleme.Count) {
            $st = if ($kritisch) { 'Kritisch' } else { 'Warnung' }
            Add-Finding 'Datentraeger' $B $name $st ((@($probleme) -join '; ') + " - $info") -Ursache 'Die SSD/Festplatte meldet selbst Fehler oder Verschleiß. Neu zugewiesene oder ausstehende Sektoren und unkorrigierbare Fehler sind Zeichen eines alternden oder defekten Datenträgers. CRC-Fehler deuten dagegen auf ein schlechtes SATA-Kabel.' -Massnahme '1. Sofort die Daten auf diesem Datenträger sichern. 2. Bei CRC-Fehlern das SATA-Kabel tauschen. 3. Steigen die Werte weiter, den Datenträger ersetzen.' -Empfehlung 'Beheben' -Hinweis $hinweis -Details $details
        } elseif ($fwHinweis) {
            Add-Finding 'Datentraeger' $B $name 'Info' $info -Ursache $fwHinweis -Massnahme 'Wichtige Daten auf dieser SSD regelmäßig sichern. Ein Firmware-Update ist möglich (Samsung Magician bzw. bootfähiges Firmware-Image), vorher aber unbedingt ein Backup.' -Empfehlung 'Beobachten' -Hinweis $hinweis -Details $details
        } else {
            Add-Finding 'Datentraeger' $B $name 'OK' $info -Hinweis $hinweis -Details $details
        }
    }

    if ($Script:AdminModus) {
        $dirty = @(Get-Data 'DirtyBits')
        if (-not (Test-QuelleFehlt 'DirtyBits' 'Dateisystem' $B 'Dateisystem') -and $dirty.Count -gt 0) {
            $kaputt = @($dirty | Where-Object { $_.Fehlerhaft -eq $true })
            $unklar = @($dirty | Where-Object { $null -eq $_.Fehlerhaft })
            if ($kaputt.Count) { Add-Finding 'Dateisystem' $B 'Dateisystem' 'Warnung' ('Als fehlerhaft markiert: ' + (@($kaputt.Laufwerk) -join ', ')) -Ursache 'Meist nach einem harten Absturz oder Stromausfall, als gerade auf das Laufwerk geschrieben wurde.' -Massnahme 'chkdsk laufen lassen: in einer Admin-Eingabeaufforderung „chkdsk X: /f“ (beim Systemlaufwerk beim nächsten Neustart).' -Empfehlung 'Beheben' -Loesungen @($kaputt | Select-Object -First 4 | ForEach-Object { New-BefehlLoesung -Titel "chkdsk $(([string]$_.Laufwerk).TrimEnd(':')): /f vorbereiten" -Befehl 'ChkdskReparatur' -Laufwerk $_.Laufwerk }) }
            elseif ($unklar.Count) { Add-Finding 'Dateisystem' $B 'Dateisystem' 'Info' ('Antwort nicht eindeutig für: ' + (@($unklar.Laufwerk) -join ', ')) -Details @($dirty) }
            else { Add-Finding 'Dateisystem' $B 'Dateisystem' 'OK' ('Alle NTFS-Laufwerke sauber: ' + (@($dirty.Laufwerk) -join ', ')) }
        }
    }

    $rein = @(Get-Data 'Reinigung')
    if (-not (Test-QuelleFehlt 'Reinigung' 'Aufraeumen' $B 'Reinigung (Speicherplatz)') -and $rein.Count) {
        $summe = [int64](($rein | Where-Object { $_.Messbar } | Measure-Object -Property Bytes -Sum).Sum)
        $zeilen = $rein | Sort-Object Bytes -Descending | ForEach-Object {
            $hinweis = @($(if ($_.Laeuft) { "läuft gerade: $($_.Laeuft)" }), $_.Sperre, $_.Fehler) | Where-Object { $_ }
            [pscustomobject]@{ Kategorie = $_.Name; Gruppe = $_.Gruppe; Groesse = (Format-Groesse $_.Bytes $_.Messbar); Dateien = $(if ($_.Messbar) { $_.Dateien } else { '' }); Hinweis = ($hinweis -join '; ') }
        }
        $loesungen = @(New-Loesung -Id 'ReinigungOeffnen' -Titel 'Reinigung öffnen und Kategorien auswählen' -Art 'Sofort')
        $hinweisAlt = ''
        if ([System.IO.Directory]::Exists((Join-Path $env:SystemDrive 'Windows.old'))) {
            $hinweisAlt = 'Zusätzlich liegt noch „Windows.old“ (das alte Windows, oft 20 GB und mehr) auf C: - das entfernt die Windows-Speicherbereinigung unter „Temporäre Dateien“ am sichersten.'
            $loesungen += New-SeitenLoesung -Titel 'Windows-Speicherbereinigung öffnen' -Seite 'ms-settings:storagesense'
        }
        $text = '{0} können freigegeben werden ({1} Kategorien)' -f (Format-Groesse $summe), $rein.Count
        $ursache = 'Bei Installationen, Updates, Abstürzen und beim Surfen sammeln sich Dateien an, die nicht mehr gebraucht werden. Browser und Programme legen außerdem große Zwischenspeicher an.'
        if ($summe -ge 5GB) {
            Add-Finding 'Aufraeumen' $B 'Reinigung (Speicherplatz)' 'Info' $text -Ursache $ursache -Massnahme 'Im Reiter „Reinigung“ die gewünschten Kategorien ankreuzen und reinigen. Vorausgewählt sind nur Kategorien ohne Nachteile.' -Empfehlung 'Entfernen' -Hinweis $hinweisAlt -Details @($zeilen) -Loesungen $loesungen
        } else {
            Add-Finding 'Aufraeumen' $B 'Reinigung (Speicherplatz)' 'OK' $text -Massnahme 'Kaum Ballast vorhanden - Reinigen ist bei Bedarf jederzeit im Reiter „Reinigung“ möglich.' -Hinweis $hinweisAlt -Details @($zeilen) -Loesungen $loesungen
        }
    }

    if ($Script:AdminModus) {
        $rp = @(Get-Data 'RestorePoints')
        if (Test-QuelleFehlt 'RestorePoints' 'Wiederherstellung' $B 'Wiederherstellungspunkte') {
        } elseif ($rp.Count -gt 0) {
            $zeilen = $rp | ForEach-Object { [pscustomobject]@{ Nr = $_.SequenceNumber; Erstellt = $(try { Format-Datum ([System.Management.ManagementDateTimeConverter]::ToDateTime($_.CreationTime)) } catch { $_.CreationTime }); Beschreibung = $_.Description } }
            Add-Finding 'Wiederherstellung' $B 'Wiederherstellungspunkte' 'OK' "$($rp.Count) vorhanden, neuester: $(@($zeilen)[-1].Erstellt)" -Details @($zeilen) -Loesungen @(New-Loesung -Id 'WiederherstellungspunktAnlegen' -Umkehrbar $false -Titel 'Jetzt einen Wiederherstellungspunkt anlegen')
        } else {
            Add-Finding 'Wiederherstellung' $B 'Wiederherstellungspunkte' 'Warnung' 'Keine vorhanden' -Ursache 'Entweder ist der Computerschutz aus, oder Windows hat die Punkte wegen Platzmangels verworfen.' -Massnahme 'Platz auf C: schaffen und einen Wiederherstellungspunkt anlegen - vor allem vor Treiber- oder Systemänderungen.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'WiederherstellungspunktAnlegen' -Umkehrbar $false -Titel 'Jetzt einen Wiederherstellungspunkt anlegen')
        }
        $vss = Get-Data 'ShadowStorage'
        if (-not (Test-QuelleFehlt 'ShadowStorage' 'Schattenkopien' $B 'Schattenkopien') -and $vss) {
            $belegt = if ($vss -match '(?i)(?:Verwendeter|Used).*?:\s*([\d\.,]+\s*[KMGT]?B)') { $Matches[1] } else { '?' }
            $max = if ($vss -match '(?i)(?:Max)\S*.*?:\s*([\d\.,]+\s*[KMGT]?B)') { $Matches[1] } else { '?' }
            Add-Finding 'Schattenkopien' $B 'Schattenkopien' 'Info' "belegt $belegt, Obergrenze $max" -Massnahme 'Nur zur Übersicht - erklärt manchmal plötzlich fehlenden Speicher auf C:.' -Empfehlung 'Behalten'
        }
        $dism = Get-Data 'DismHealth'
        if (-not (Test-QuelleFehlt 'DismHealth' 'DISM' $B 'Komponentenspeicher (DISM)') -and $dism) {
            if ($dism -match '(?i)keine Komponentenspeicherbesch|No component store corruption') { Add-Finding 'DISM' $B 'Komponentenspeicher (DISM)' 'OK' 'Keine Beschädigung erkannt' }
            elseif ($dism -match '(?i)reparierbar|repairable') { Add-Finding 'DISM' $B 'Komponentenspeicher (DISM)' 'Warnung' 'Beschädigung erkannt, reparierbar' -Ursache 'Teile der Windows-Reparaturvorlage sind beschädigt - meist nach Abstürzen während eines Updates.' -Massnahme 'Mit DISM reparieren (lädt fehlende Teile von Windows Update), danach die Systemdateien mit SFC prüfen.' -Empfehlung 'Beheben' -Loesungen @((New-Loesung -Id 'DismReparieren' -Titel 'Komponentenspeicher mit DISM reparieren' -Umkehrbar $false -Dauer '10-20 Min.'), (New-Loesung -Id 'SfcReparieren' -Titel 'Danach Systemdateien mit SFC reparieren' -Umkehrbar $false -Dauer '10-15 Min.')) }
            else { Add-Finding 'DISM' $B 'Komponentenspeicher (DISM)' 'Info' (Format-Kurz $dism 200) }
        }
        $ana = Get-Data 'DismAnalyse'
        if (-not (Test-QuelleFehlt 'DismAnalyse' 'DISM' $B 'Komponentenspeicher-Größe') -and $ana) {
            $groesse = if ($ana -match '(?i)(?:Tats\S{1,2}chliche Gr\S{1,3}e des Komponentenspeichers|Actual Size of Component Store)\s*:\s*([\d\.,]+\s*[KMGT]?B)') { $Matches[1] } else { '?' }
            $frei = if ($ana -match '(?i)(?:Freigebbare Pakete|Reclaimable Packages)\s*:\s*(\d+)') { [int]$Matches[1] } else { 0 }
            $empf = $ana -match '(?i)(?:Bereinigung des Komponentenspeichers empfohlen|Component Store Cleanup Recommended)\s*:\s*(Ja|Yes)'
            if ($empf) { Add-Finding 'DISM' $B 'Komponentenspeicher-Größe' 'Info' "$groesse, $frei freigebbare Pakete - Bereinigung empfohlen" -Massnahme 'In der Admin-PowerShell: Dism /online /Cleanup-Image /StartComponentCleanup (ohne /ResetBase - sonst lassen sich Updates nicht mehr deinstallieren).' -Empfehlung 'Prüfen' -Loesungen @(New-BefehlLoesung -Titel 'DISM-Bereinigung vorbereiten' -Befehl 'DismAufraeumen') }
            else { Add-Finding 'DISM' $B 'Komponentenspeicher-Größe' 'OK' "$groesse, keine Bereinigung nötig" }
        }
        $scan = Get-Data 'DismScan'
        if (-not (Test-QuelleFehlt 'DismScan' 'DISM' $B 'Komponentenspeicher-Tiefenscan') -and $scan) {
            if ($scan -match '(?i)keine Komponentenspeicherbesch|No component store corruption') { Add-Finding 'DISM' $B 'Komponentenspeicher-Tiefenscan' 'OK' 'Vollständig geprüft, keine Beschädigung' }
            elseif ($scan -match '(?i)reparierbar|repairable') { Add-Finding 'DISM' $B 'Komponentenspeicher-Tiefenscan' 'Warnung' 'Beschädigung gefunden, reparierbar' -Ursache 'Der Tiefenscan hat beschädigte Teile der Windows-Reparaturvorlage gefunden.' -Massnahme 'Mit DISM reparieren, danach SFC.' -Empfehlung 'Beheben' -Loesungen @((New-Loesung -Id 'DismReparieren' -Titel 'Komponentenspeicher mit DISM reparieren' -Umkehrbar $false -Dauer '10-20 Min.'), (New-Loesung -Id 'SfcReparieren' -Titel 'Danach Systemdateien mit SFC reparieren' -Umkehrbar $false -Dauer '10-15 Min.')) }
            else { Add-Finding 'DISM' $B 'Komponentenspeicher-Tiefenscan' 'Info' (Format-Kurz $scan 200) }
        }
        $cd = @(Get-Data 'Chkdsk')
        # Startet chkdsk.exe gar nicht (auf PC 1 seit 10/2026 so - auch mit Adminrechten), ist das ein eigener Befund statt „nicht lesbar“
        $chkdskGesperrt = [string]$Script:QuellenFehler['Chkdsk'] -match '(?i)Ausführen des Programms|failed to run|Win32Exception'
        # Riot Vanguard (vgk.sys) verweigert jedem Prozess namens chkdsk.exe den Start - auf PC 1 am 28.10.2026 belegt (vgk-Log waechst je Versuch)
        $vanguard = $chkdskGesperrt -and [string](Get-CimInstance Win32_SystemDriver -Filter "Name='vgk'" -ErrorAction SilentlyContinue).State -eq 'Running'
        if ($vanguard) {
            Add-Finding 'Chkdsk' $B 'Dateisysteme (chkdsk)' 'Info' 'chkdsk wird von Riot Vanguard blockiert' -Ursache 'Der Anti-Cheat-Treiber von Riot (Vanguard, für League of Legends/Valorant) verweigert jedem Programm namens chkdsk.exe den Start - auch mit Adminrechten. Kein Defekt von Windows oder des Laufwerks.' -Massnahme 'Ersatzweise prüft dieser Modus die NTFS-Ereignisse und den Reparaturstatus (siehe unten). Für einen echten chkdsk-Lauf: Rechtsklick auf das Vanguard-Symbol im Infobereich → „Vanguard beenden“, dann „Gründlich“ prüfen. Für League of Legends danach den PC neu starten.' -Hinweis ('Meldung: ' + (Format-Kurz $Script:QuellenFehler['Chkdsk'] 140))
        }
        elseif ($chkdskGesperrt) {
            Add-Finding 'Chkdsk' $B 'Dateisysteme (chkdsk)' 'Warnung' 'chkdsk lässt sich nicht starten - Windows verweigert den Programmstart' -Ursache 'Windows blockiert den Start von chkdsk.exe, bevor das Programm überhaupt läuft - auch mit Adminrechten. Die Datei selbst ist normalerweise intakt; meist steckt eine Richtlinie oder Sicherheitssoftware dahinter. Ohne chkdsk lässt sich ein beschädigtes Dateisystem im Notfall nicht reparieren.' -Massnahme 'Ersatzweise prüft dieser Modus die NTFS-Ereignisse und den Reparaturstatus (siehe unten). Zum Aufklären in einer Admin-Eingabeaufforderung „chkdsk C:“ eingeben: Kommt sofort „Zugriff verweigert“, ist der Start blockiert - dann Sicherheitssoftware und Richtlinien prüfen.' -Empfehlung 'Prüfen' -Loesungen @(New-BefehlLoesung -Titel 'chkdsk C: zum Testen vorbereiten' -Befehl 'ChkdskLesen' -Laufwerk 'C') -Hinweis ('Meldung: ' + (Format-Kurz $Script:QuellenFehler['Chkdsk'] 140))
        }
        elseif (-not (Test-QuelleFehlt 'Chkdsk' 'chkdsk' $B 'Dateisysteme (chkdsk)') -and $cd.Count) {
            $schlecht = @($cd | Where-Object { $_.InOrdnung -eq $false })
            $zeilen = @($cd | ForEach-Object { [pscustomobject]@{ Laufwerk = $_.Laufwerk + ':'; Ergebnis = $(if ($_.InOrdnung) { 'in Ordnung' } elseif ($_.InOrdnung -eq $false) { 'FEHLER' } else { 'unklar' }); Sekunden = $_.Sekunden; Meldung = $_.Meldung } })
            if ($schlecht.Count) { Add-Finding 'Chkdsk' $B 'Dateisysteme (chkdsk)' 'Warnung' ('Fehler auf ' + (@($schlecht | ForEach-Object { $_.Laufwerk + ':' }) -join ', ')) -Ursache 'Die Verwaltungsdaten des Dateisystems sind beschädigt - meist nach hartem Ausschalten oder Abstürzen, selten ein Laufwerksdefekt.' -Massnahme 'In der Admin-PowerShell „chkdsk <Laufwerk>: /scan“ und danach „chkdsk <Laufwerk>: /spotfix“ (kurzer Neustart beim Systemlaufwerk). Vorher wichtige Daten sichern und SMART-Werte des Laufwerks ansehen.' -Empfehlung 'Beheben' -Loesungen @($schlecht | Select-Object -First 4 | ForEach-Object { New-BefehlLoesung -Titel "chkdsk $($_.Laufwerk): /scan vorbereiten" -Befehl 'ChkdskScan' -Laufwerk $_.Laufwerk }) -Details $zeilen }
            else { Add-Finding 'Chkdsk' $B 'Dateisysteme (chkdsk)' 'OK' ("$($cd.Count) Laufwerk(e) im Lesemodus geprüft") -Details $zeilen }
        }
        $sfc = Get-Data 'SfcVerify'
        if (-not (Test-QuelleFehlt 'SfcVerify' 'SFC' $B 'Systemdateien (SFC)') -and $sfc) {
            if ($sfc -match '(?i)keine Integrit|did not find any integrity violations') { Add-Finding 'SFC' $B 'Systemdateien (SFC)' 'OK' 'Keine Integritätsverletzungen' }
            elseif ($sfc -match '(?i)Integrit\S*verletzungen gefunden|found integrity violations') { Add-Finding 'SFC' $B 'Systemdateien (SFC)' 'Warnung' 'Integritätsverletzungen gefunden' -Ursache 'Geschützte Windows-Dateien sind beschädigt oder wurden verändert.' -Massnahme 'Mit SFC reparieren lassen; meldet SFC danach weiter Fehler, vorher DISM reparieren.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'SfcReparieren' -Titel 'Systemdateien mit SFC reparieren' -Umkehrbar $false -Dauer '10-15 Min.') }
            else { Add-Finding 'SFC' $B 'Systemdateien (SFC)' 'Info' (Format-Kurz $sfc 200) }
        }
    }

    # Driver Verifier: absichtlich scharf gestellte Treiberpruefung loest selbst Bluescreens aus - gehoert nach jedem Test abgeschaltet
    $mm = Get-Data 'Treiberpruefung'
    if (-not (Test-QuelleFehlt 'Treiberpruefung' 'Treiberpruefung' 'Stabilitaet' 'Treiberüberprüfung (Driver Verifier)') -and $mm) {
        $vd = [string]$mm['VerifyDrivers']; $vl = $mm['VerifyDriverLevel']
        if ($vd.Trim() -or ($vl -and [int64]$vl -ne 0)) { Add-Finding 'Treiberpruefung' 'Stabilitaet' 'Treiberüberprüfung (Driver Verifier)' 'Kritisch' "aktiv: $(if ($vd.Trim()) { $vd } else { 'Stufe ' + $vl })" -Ursache 'Die Treiberüberprüfung ist eingeschaltet. Sie testet Treiber absichtlich unter Extrembedingungen und löst bei jedem Fehler sofort einen Bluescreen aus - im Alltag macht sie den PC instabil und langsamer.' -Massnahme 'In der Admin-Eingabeaufforderung „verifier /reset“ ausführen und neu starten.' -Empfehlung 'Beheben' }
        else { Add-Finding 'Treiberpruefung' 'Stabilitaet' 'Treiberüberprüfung (Driver Verifier)' 'OK' 'aus' }
    }

    # Defender-Funde: was gefunden wurde und ob es entfernt werden konnte
    $funde = @(Get-Data 'Bedrohungen')
    if (-not (Test-QuelleFehlt 'Bedrohungen' 'Bedrohungen' 'Sicherheit & Updates' 'Defender-Funde')) {
        $offen = @($funde | Where-Object { -not $_.Erfolgreich })
        $neu = @($funde | Where-Object { $_.Zeit -and ([datetime]$_.Zeit) -gt (Get-Date).AddDays(-$Tage) })
        $zeilen = @($funde | Sort-Object Zeit -Descending | Select-Object -First 15 | ForEach-Object { [pscustomobject]@{ Zeit = (Format-Datum $_.Zeit); Bedrohung = $_.Name; Entfernt = $(if ($_.Erfolgreich) { 'ja' } else { 'NEIN' }); Dateien = $_.Dateien } })
        $zuletzt = Get-Neuestes @($funde | ForEach-Object { $_.Zeit })
        if ($offen.Count) { Add-Finding 'Bedrohungen' 'Sicherheit & Updates' 'Defender-Funde' 'Kritisch' "$($offen.Count) Bedrohung(en) nicht entfernt" -Ursache 'Microsoft Defender hat Schadsoftware gefunden, konnte sie aber nicht vollständig entfernen.' -Massnahme 'Windows-Sicherheit → Viren- & Bedrohungsschutz → Schutzverlauf öffnen und die Empfehlung ausführen; danach eine vollständige Überprüfung oder die Offline-Überprüfung starten.' -Empfehlung 'Beheben' -Details $zeilen -Zuletzt $zuletzt -Loesungen @(New-SeitenLoesung -Titel 'Windows-Sicherheit öffnen' -Seite 'windowsdefender://threat') }
        elseif ($neu.Count) { Add-Finding 'Bedrohungen' 'Sicherheit & Updates' 'Defender-Funde' 'Warnung' "$($neu.Count) Fund(e) in den letzten $Tage Tagen - entfernt" -Ursache 'Defender hat in letzter Zeit etwas abgefangen und entfernt. Wichtig ist, woher es kam (Download, USB-Stick, Crack, Mod).' -Massnahme 'In den Details die betroffenen Dateien ansehen. Kam es aus einer bestimmten Quelle, diese meiden; bei Unsicherheit eine vollständige Überprüfung starten. Danach archivieren.' -Empfehlung 'Prüfen' -Details $zeilen -Zuletzt $zuletzt }
        else { Add-Finding 'Bedrohungen' 'Sicherheit & Updates' 'Defender-Funde' 'OK' $(if ($funde.Count) { "Keine neuen Funde ($($funde.Count) ältere, alle entfernt)" } else { 'Keine Funde' }) -Details $zeilen }
    }

    # hosts-Datei: Umleitungen hier ueberstimmen jeden DNS-Server - Schadsoftware sperrt damit gern Updates und Virenschutz
    $hosts = @(Get-Data 'Hosts')
    if (-not (Test-QuelleFehlt 'Hosts' 'Hosts' 'Netzwerk' 'hosts-Datei')) {
        $eintraege = @($hosts | ForEach-Object { ([string]$_ -replace '#.*$', '').Trim() } | Where-Object { $_ -and $_ -notmatch '(?i)\slocalhost$|^::1\s|^127\.0\.0\.1\s+localhost' })
        $heikel = @($eintraege | Where-Object { $_ -match '(?i)microsoft|windowsupdate|windows\.com|defender|msftncsi|live\.com|office' })
        if ($heikel.Count) { Add-Finding 'Hosts' 'Netzwerk' 'hosts-Datei' 'Warnung' "$($heikel.Count) Umleitung(en) für Microsoft-Dienste" -Ursache 'Die hosts-Datei leitet Adressen von Microsoft (Updates, Defender, Aktivierung) um oder sperrt sie. Das machen Anti-Telemetrie-Tools - aber auch Schadsoftware, um Updates und Virenschutz auszuschalten.' -Massnahme 'Prüfen, ob du das selbst eingerichtet hast. Wenn nicht: die Zeilen in C:\Windows\System32\drivers\etc\hosts (als Admin) entfernen und eine vollständige Defender-Überprüfung starten.' -Empfehlung 'Prüfen' -Details @($heikel | ForEach-Object { [pscustomobject]@{ Eintrag = $_ } }) }
        elseif ($eintraege.Count) { Add-Finding 'Hosts' 'Netzwerk' 'hosts-Datei' 'Info' "$($eintraege.Count) eigene Einträge" -Massnahme 'Eigene Umleitungen, keine für Microsoft-Dienste. Unbekannte Einträge prüfen, sonst archivieren.' -Empfehlung 'Prüfen' -Details @($eintraege | Select-Object -First 20 | ForEach-Object { [pscustomobject]@{ Eintrag = $_ } }) }
        else { Add-Finding 'Hosts' 'Netzwerk' 'hosts-Datei' 'OK' 'Unverändert (keine Umleitungen)' }
    }

    if ($Modus -ne 'Gruendlich') { return }

    # Ersatz bzw. Ergaenzung zu chkdsk: was NTFS selbst gemeldet hat, und ob eine Reparatur ansteht
    $ntfs = @(Get-Data 'NtfsEreignisse')
    if (-not (Test-QuelleFehlt 'NtfsEreignisse' 'Dateisystem' $B 'Dateisystem-Ereignisse (NTFS)')) {
        $zeilen = @($ntfs | Select-Object -First 15 | ForEach-Object { [pscustomobject]@{ Zeit = (Format-Datum $_.TimeCreated); Id = $_.Id; Stufe = $_.LevelDisplayName; Meldung = (Format-Kurz $_.Message 160) } })
        $fehler = @($ntfs | Where-Object { $_.Level -le 2 })
        if ($fehler.Count) { Add-Finding 'Dateisystem' $B 'Dateisystem-Ereignisse (NTFS)' 'Warnung' "$($fehler.Count) Fehler in den letzten $Tage Tagen" -Ursache 'NTFS hat Beschädigungen oder Schreibfehler im Dateisystem gemeldet - meist nach harten Abschaltungen, selten wegen eines alternden Datenträgers.' -Massnahme 'Betroffenes Laufwerk in den Details ablesen, wichtige Daten sichern, dann SMART-Werte prüfen. Windows repariert erkannte Fehler oft selbst beim nächsten Neustart.' -Empfehlung 'Prüfen' -Details $zeilen }
        elseif ($ntfs.Count) { Add-Finding 'Dateisystem' $B 'Dateisystem-Ereignisse (NTFS)' 'Info' "$($ntfs.Count) Warnung(en), keine Fehler in den letzten $Tage Tagen" -Massnahme 'Einzelne Warnungen entstehen z. B. beim Abziehen von USB-Laufwerken ohne Auswerfen. Nur bei Häufung genauer hinsehen.' -Empfehlung 'Beobachten' -Details $zeilen }
        else { Add-Finding 'Dateisystem' $B 'Dateisystem-Ereignisse (NTFS)' 'OK' "Keine Fehler oder Warnungen in den letzten $Tage Tagen" }
    }
    $rep = @(Get-Data 'FsReparatur')
    if (-not (Test-QuelleFehlt 'FsReparatur' 'Dateisystem' $B 'Dateisystem-Reparaturstatus') -and $rep.Count) {
        $offen = @($rep | Where-Object { $_.InOrdnung -eq $false })
        $zeilen = @($rep | ForEach-Object { [pscustomobject]@{ Laufwerk = $_.Laufwerk + ':'; Ergebnis = $(if ($_.InOrdnung) { 'keine Reparatur nötig' } elseif ($_.InOrdnung -eq $false) { 'REPARATUR ANSTEHEND' } else { 'unklar' }); Meldung = $_.Text } })
        if ($offen.Count) { Add-Finding 'Dateisystem' $B 'Dateisystem-Reparaturstatus' 'Warnung' ('Reparatur ansteht auf ' + (@($offen | ForEach-Object { $_.Laufwerk + ':' }) -join ', ')) -Ursache 'NTFS hat eine Beschädigung erkannt und für die Reparatur vorgemerkt.' -Massnahme 'PC neu starten - Windows repariert vorgemerkte Fehler beim Hochfahren. Danach erneut gründlich prüfen.' -Empfehlung 'Beheben' -Details $zeilen }
        else { Add-Finding 'Dateisystem' $B 'Dateisystem-Reparaturstatus' 'OK' "$($rep.Count) Laufwerk(e): keine Reparatur vorgemerkt" -Details $zeilen }
    }

    $mem = @(Get-Data 'Speicherdiagnose')
    if (-not (Test-QuelleFehlt 'Speicherdiagnose' 'Speicherdiagnose' 'Hardware & Treiber' 'Windows-Speicherdiagnose')) {
        if (-not $mem.Count) { Add-Finding 'Speicherdiagnose' 'Hardware & Treiber' 'Windows-Speicherdiagnose' 'Info' 'Noch nie gelaufen' -Massnahme 'Bei unerklärlichen Abstürzen oder WHEA-Fehlern lohnt ein RAM-Test: Win+R → mdsched → „Jetzt neu starten“ (dauert ca. 15-30 Min.). Das Ergebnis erscheint hier bei der nächsten gründlichen Prüfung.' -Empfehlung 'Beobachten' }
        else {
            $neu = $mem[0]
            $zeilen = @($mem | ForEach-Object { [pscustomobject]@{ Zeit = (Format-Datum $_.Zeit); Meldung = $_.Meldung } })
            if ($neu.Meldung -match '(?i)Hardwarefehler|hardware errors|Fehler erkannt|errors were detected') { Add-Finding 'Speicherdiagnose' 'Hardware & Treiber' 'Windows-Speicherdiagnose' 'Kritisch' "Fehler gefunden ($(Format-Datum $neu.Zeit))" -Ursache 'Der Arbeitsspeichertest hat Fehler gefunden - defekter Riegel, zu scharfes RAM-Profil (EXPO/XMP) oder zu wenig Spannung.' -Massnahme 'EXPO/XMP testweise aus und erneut testen; bleiben Fehler, Riegel einzeln testen.' -Empfehlung 'Prüfen' -Details $zeilen }
            else { Add-Finding 'Speicherdiagnose' 'Hardware & Treiber' 'Windows-Speicherdiagnose' 'OK' "Letzter Test $(Format-Datum $neu.Zeit): keine Fehler" -Details $zeilen }
        }
    }

    $sig = @(Get-Data 'Signaturen')
    if (-not (Test-QuelleFehlt 'Signaturen' 'Signaturen' 'Sicherheit & Updates' 'Signaturen (Treiber, Autostart)') -and $sig.Count) {
        $schlecht = @($sig | Where-Object { $_.Status -ne 'Valid' })
        $zeilen = @($schlecht | ForEach-Object { [pscustomobject]@{ Art = $_.Art; Name = $_.Name; Status = $(switch ($_.Status) { 'NotSigned' { 'nicht signiert' } 'HashMismatch' { 'VERÄNDERT' } 'NotTrusted' { 'nicht vertrauenswürdig' } default { $_ } }); Datei = $_.Datei } })
        $veraendert = @($schlecht | Where-Object { $_.Status -in 'HashMismatch', 'NotTrusted' })
        $text = "$($sig.Count) Dateien geprüft"
        if ($veraendert.Count) { Add-Finding 'Signaturen' 'Sicherheit & Updates' 'Signaturen (Treiber, Autostart)' 'Kritisch' "$text, $($veraendert.Count) verändert oder nicht vertrauenswürdig" -Ursache 'Die Datei passt nicht mehr zu ihrer Signatur oder das Zertifikat ist nicht vertrauenswürdig - nach einem Update selten, bei Schadsoftware typisch.' -Massnahme 'Mit Microsoft Defender (vollständige Überprüfung) scannen und das zugehörige Programm neu installieren.' -Empfehlung 'Prüfen' -Details $zeilen }
        elseif ($schlecht.Count) { Add-Finding 'Signaturen' 'Sicherheit & Updates' 'Signaturen (Treiber, Autostart)' 'Info' "$text, $($schlecht.Count) ohne gültige Signatur" -Massnahme 'Unsignierte Autostart-Programme sind bei kleinen Tools und Skripten normal. Unbekannte Einträge prüfen: Herkunft klar? Sonst im Reiter „Befunde“ unter Autostart deaktivieren.' -Empfehlung 'Prüfen' -Details $zeilen }
        else { Add-Finding 'Signaturen' 'Sicherheit & Updates' 'Signaturen (Treiber, Autostart)' 'OK' "$text, alle gültig signiert" }
    }

    # Startkonfiguration: Testmodus, abgeschaltete Signaturpruefung, Kernel-Debugger oder haengengebliebener abgesicherter Modus
    $bcd = [string](Get-Data 'Bootkonfiguration')
    if (-not (Test-QuelleFehlt 'Bootkonfiguration' 'Startkonfiguration' 'Sicherheit & Updates' 'Startkonfiguration') -and $bcd) {
        $flags = @(foreach ($n in 'testsigning', 'nointegritychecks', 'debug', 'bootdebug') { if ($bcd -match "(?im)^$n\s+(Yes|Ja)\s*$") { $n } })
        if ($bcd -match '(?im)^safeboot\s+\S') { $flags += 'safeboot' }
        if ($flags.Count) { Add-Finding 'Startkonfiguration' 'Sicherheit & Updates' 'Startkonfiguration' 'Warnung' ('Gesetzt: ' + ($flags -join ', ')) -Ursache 'Diese Startoptionen schalten Schutzmechanismen ab (Treibersignatur, Integritätsprüfung) oder lassen Windows im Debug-/abgesicherten Modus starten - meist Überbleibsel von Treiber-Experimenten oder Tuning-Tools, bei Anti-Cheat-Spielen ein häufiger Startgrund für Fehler.' -Massnahme 'In der Admin-Eingabeaufforderung zurücksetzen, z. B. „bcdedit /set testsigning off“ bzw. „bcdedit /deletevalue safeboot“, danach neu starten.' -Empfehlung 'Prüfen' }
        else { Add-Finding 'Startkonfiguration' 'Sicherheit & Updates' 'Startkonfiguration' 'OK' 'Normaler Start - kein Testmodus, Signaturprüfung an, kein Debugger' }
    }

    # WMI-Datenbank: ist sie beschaedigt, liefern viele Programme (und diese Diagnose) falsche oder keine Werte
    $wmi = [string](Get-Data 'WmiDatenbank')
    if (-not (Test-QuelleFehlt 'WmiDatenbank' 'WMI' $B 'WMI-Datenbank') -and $wmi) {
        if ($wmi -match '(?i)nicht konsistent|inkonsistent|inconsistent|not consistent') { Add-Finding 'WMI' $B 'WMI-Datenbank' 'Warnung' 'Beschädigt' -Ursache 'Die WMI-Datenbank, aus der Windows und Programme Hardware- und Systemdaten lesen, ist inkonsistent - meist nach Abstürzen während eines Updates.' -Massnahme 'In der Admin-Eingabeaufforderung „winmgmt /salvagerepository“ ausführen und neu starten.' -Empfehlung 'Beheben' -Loesungen @(New-BefehlLoesung -Titel 'WMI-Reparatur vorbereiten' -Befehl 'WmiReparatur') }
        elseif ($wmi -match '(?i)konsistent|consistent') { Add-Finding 'WMI' $B 'WMI-Datenbank' 'OK' 'Konsistent' }
        else { Add-Finding 'WMI' $B 'WMI-Datenbank' 'Info' (Format-Kurz $wmi 160) -Empfehlung 'Prüfen' }
    }

    # TRIM: ohne TRIM werden SSDs mit der Zeit langsamer und verschleissen schneller
    $trim = [string](Get-Data 'Trim')
    if (-not (Test-QuelleFehlt 'Trim' 'Trim' 'Laufwerke' 'TRIM für SSDs') -and $trim) {
        if ($trim -match '(?i)NTFS DisableDeleteNotify\s*=\s*1') { Add-Finding 'Trim' 'Laufwerke' 'TRIM für SSDs' 'Warnung' 'aus' -Ursache 'Windows teilt SSDs nicht mehr mit, welche Bereiche frei sind - die SSD wird mit der Zeit langsamer und verschleißt schneller. Meist von einem Tuning-Tool abgeschaltet.' -Massnahme 'In der Admin-Eingabeaufforderung „fsutil behavior set DisableDeleteNotify 0“ ausführen.' -Empfehlung 'Beheben' }
        elseif ($trim -match '(?i)NTFS DisableDeleteNotify\s*=\s*0') { Add-Finding 'Trim' 'Laufwerke' 'TRIM für SSDs' 'OK' 'an' }
    }

    # Wo der Platz steckt - komplette Begehung jedes Laufwerks (Ebene 1 und 2)
    foreach ($lw in @(Get-Data 'Speicherbelegung')) {
        if (Test-QuelleFehlt 'Speicherbelegung' 'Speicherbelegung' 'Laufwerke' 'Speicherbelegung') { break }
        $prozent = if ($lw.Groesse) { [Math]::Round(100 * $lw.Frei / $lw.Groesse) } else { 0 }
        $oben = @($lw.Ordner | Where-Object { $_.Ebene -eq 1 } | Select-Object -First 6)
        $zeilen = foreach ($o in $oben) {
            [pscustomobject]@{ Ordner = $o.Pfad; Groesse = (Format-Groesse $o.Bytes); Dateien = $o.Dateien }
            foreach ($u in @($lw.Ordner | Where-Object { $_.Ebene -eq 2 -and $_.Pfad.StartsWith($o.Pfad + '\', [System.StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 3)) {
                [pscustomobject]@{ Ordner = '   ' + $u.Pfad; Groesse = (Format-Groesse $u.Bytes); Dateien = $u.Dateien }
            }
        }
        $groesste = (@($oben | Select-Object -First 3 | ForEach-Object { "$(Split-Path -Leaf $_.Pfad) $(Format-Groesse $_.Bytes)" }) -join ', ')
        $pruefung = "Speicherbelegung $($lw.Laufwerk):"
        if ($prozent -lt 15) { Add-Finding 'Speicherbelegung' 'Laufwerke' $pruefung 'Info' "$prozent % frei - größte Ordner: $groesste" -Massnahme 'In den Details stehen die größten Ordner bis zur zweiten Ebene. Spiele lassen sich in Steam/Epic auf ein anderes Laufwerk verschieben; Caches und Protokolle räumt der Reiter „Reinigung“ auf.' -Empfehlung 'Prüfen' -Details @($zeilen) -Hinweis "Ganzes Laufwerk in $($lw.Sekunden) s gelesen" }
        else { Add-Finding 'Speicherbelegung' 'Laufwerke' $pruefung 'OK' "$prozent % frei - größte Ordner: $groesste" -Details @($zeilen) }
    }
}

function Test-SoftwareAutostart {
    $B = 'Software & Autostart'

    $auto = @(Get-Data 'Autostart')
    if ($auto.Count -gt 0) {
        $an = @($auto | Where-Object { $_.Status -eq 'an' })
        $kaputt = @($an | Where-Object { $_.ZielVorhanden -eq $false })
        $zeilen = $auto | Sort-Object Status, Name | ForEach-Object { [pscustomobject]@{ Status = $_.Status; Name = $_.Name; Ort = $_.Ort; Ziel = $_.Ziel; Vorhanden = $(if ($_.ZielVorhanden -eq $false) { 'FEHLT' } elseif ($null -eq $_.ZielVorhanden) { '?' } else { 'ja' }) } }
        if ($kaputt.Count) {
            $loesungen = @($kaputt | Where-Object { $_.Ort -like 'Run*' } | ForEach-Object { New-Loesung -Id 'AutostartDeaktivieren' -Titel ('Autostart „' + $_.Name + '“ deaktivieren') -Parameter @{ Ort = $_.Ort; Name = $_.Name } -Beschreibung "Fehlendes Programm: $($_.Ziel)" })
            Add-Finding 'Autostart' $B 'Autostart' 'Warnung' "$($an.Count) aktiv, davon $($kaputt.Count) mit fehlendem Programm: $(@($kaputt.Name) -join ', ')" -Ursache 'Das Programm wurde deinstalliert oder verschoben, sein Autostart-Eintrag ist geblieben. Windows versucht es bei jeder Anmeldung zu starten.' -Massnahme 'Den Eintrag deaktivieren (umkehrbar, wie im Task-Manager unter „Autostart“).' -Empfehlung 'Entfernen' -Details @($zeilen) -Loesungen $loesungen
        } else {
            Add-Finding 'Autostart' $B 'Autostart' 'OK' "$($an.Count) aktiv, $($auto.Count - $an.Count) deaktiviert" -Hinweis 'Enthält Run-Schlüssel, Autostart-Ordner und Store-Apps - wie der Autostart-Tab im Task-Manager.' -Details @($zeilen)
        }
    }

    $tasks = @(Get-Data 'Tasks')
    if ($tasks.Count -gt 0) {
        $kaputt = @($tasks | Where-Object { $_.Status -ne 'Disabled' -and $_.ProgrammVorhanden -eq $false })
        $zeilen = $tasks | Sort-Object Aufgabe | Select-Object Aufgabe, Status, ZuletztGelaufen, Ergebnis, Programm
        if ($kaputt.Count) {
            $loesungen = @($kaputt | ForEach-Object { New-Loesung -Id 'AufgabeDeaktivieren' -Titel "Aufgabe $($_.Aufgabe) deaktivieren" -Parameter @{ Pfad = $_.Aufgabe } -Beschreibung "Fehlendes Programm: $($_.Programm)" })
            $ursache = 'Die Software, zu der die Aufgabe gehört, ist nicht (mehr) installiert - die Aufgabe startet also ins Leere.'
            if (@($kaputt | Where-Object { $_.Aufgabe -like '\ASUS\*' }).Count) { $ursache += ' Bei ASUS-Aufgaben: Armoury Crate legt solche Aufgaben für Geräte an (z. B. P508 = ROG-Strix-Carry-Maus), auch wenn das Gerät gar nicht vorhanden ist.' }
            Add-Finding 'Aufgaben' $B 'Geplante Aufgaben' 'Warnung' "$($kaputt.Count) aktive Aufgabe(n) starten ein fehlendes Programm: $(@($kaputt.Aufgabe) -join ', ')" -Ursache $ursache -Massnahme 'Die Aufgabe deaktivieren (umkehrbar). Sie richtet keinen Schaden an, erzeugt aber Fehler im Hintergrund.' -Empfehlung 'Entfernen' -Details @($zeilen) -Loesungen $loesungen
        } else {
            Add-Finding 'Aufgaben' $B 'Geplante Aufgaben' 'OK' "$($tasks.Count) Aufgaben von Drittanbietern, alle Programme vorhanden" -Details @($zeilen)
        }
    }

    $dienste = @(Get-Data 'Services')
    if ($dienste.Count -gt 0) {
        foreach ($n in $Script:Profil.DiensteMuessenLaufen.Keys) {
            $d = $dienste | Where-Object { $_.Name -eq $n } | Select-Object -First 1
            $grund = $Script:Profil.DiensteMuessenLaufen[$n]
            if (-not $d) { Add-Finding 'ProfilDienst' $B "Dienst $n" 'Warnung' 'nicht installiert' -Ursache "$grund Der Dienst gehört zum Logitech-Treiber und fehlt - vermutlich nach einem G-HUB-Update oder einer Deinstallation." -Massnahme 'Logitech G HUB neu installieren bzw. reparieren.' -Empfehlung 'Beheben' }
            elseif ($d.Status -ne 'Running') { Add-Finding 'ProfilDienst' $B "Dienst $n" 'Warnung' "Status: $($d.Status), Start: $($d.Start)" -Ursache "$grund Er läuft gerade nicht - entweder abgestürzt oder deaktiviert." -Massnahme 'Den Dienst starten und auf „Automatisch“ stellen. Danach G HUB neu starten, falls das RGB nicht sofort zurückkommt.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'DienstStarten' -Titel "Dienst $n starten und auf Automatisch stellen" -Parameter @{ Name = $n }) }
            else { Add-Finding 'ProfilDienst' $B "Dienst $n" 'OK' 'läuft' -Hinweis $grund }
        }
        foreach ($n in $Script:Profil.DiensteMuessenAusSein.Keys) {
            $d = $dienste | Where-Object { $_.Name -eq $n } | Select-Object -First 1
            $grund = $Script:Profil.DiensteMuessenAusSein[$n]
            if (-not $d) { Add-Finding 'ProfilDienst' $B "Dienst $n" 'OK' 'nicht installiert' -Hinweis $grund }
            elseif ($d.Start -ne 'Disabled') { Add-Finding 'ProfilDienst' $B "Dienst $n" 'Warnung' "wieder aktiv (Start: $($d.Start), Status: $($d.Status))" -Ursache "$grund Er wurde vermutlich durch ein Treiber- oder Armoury-Crate-Update wieder eingeschaltet." -Massnahme 'Den Dienst wieder deaktivieren (umkehrbar).' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'DienstDeaktivieren' -Titel "Dienst $n wieder deaktivieren" -Parameter @{ Name = $n }) }
            else { Add-Finding 'ProfilDienst' $B "Dienst $n" 'OK' 'deaktiviert' -Hinweis $grund }
        }
        $verwaist = @($dienste | Where-Object { $_.Start -ne 'Disabled' -and $_.Vorhanden -eq $false })
        if ($verwaist.Count) {
            $loesungen = @($verwaist | ForEach-Object { New-Loesung -Id 'DienstDeaktivieren' -Titel "Dienst $($_.Name) deaktivieren" -Parameter @{ Name = $_.Name } -Beschreibung "Fehlendes Programm: $($_.Pfad)" })
            Add-Finding 'VerwaisteDienste' $B 'Verwaiste Dienste' 'Warnung' "$($verwaist.Count) Dienst(e) zeigen auf ein fehlendes Programm: $(@($verwaist.Name) -join ', ')" -Ursache 'Die zugehörige Software wurde deinstalliert, der Dienst-Eintrag ist geblieben.' -Massnahme 'Den Dienst deaktivieren (umkehrbar). Ganz entfernen lässt er sich später, wenn sicher ist, dass nichts fehlt.' -Empfehlung 'Entfernen' -Details @($verwaist | Select-Object Name, Anzeige, Start, Pfad) -Loesungen $loesungen
        } else {
            Add-Finding 'VerwaisteDienste' $B 'Verwaiste Dienste' 'OK' "Keine ($($dienste.Count) Dienste geprüft)"
        }
    }

    # --- Hintergrund-Skripte (doppelt laufende Watcher waren hier schon mal die Ursache fuer Flackern) ---
    $proz = @(Get-Data 'Processes')
    $skripte = @(foreach ($p in $proz) {
        if ($p.Name -notin @('powershell.exe', 'pwsh.exe', 'wscript.exe', 'cscript.exe') -or $p.ProcessId -eq $PID) { continue }
        $cmd = [string]$p.CommandLine
        $datei = $null
        if (-not $cmd) { $datei = '(Befehlszeile nicht lesbar - läuft mit Adminrechten)' }
        elseif ($cmd -match '(?i)-File\s+"?([^"]+?\.ps1)') { $datei = $Matches[1] }
        elseif ($cmd -match '(?i)"?([^"]+?\.(vbs|js))"?') { $datei = $Matches[1] }
        if (-not $datei) { continue }
        [pscustomobject]@{ Skript = $datei; Prozess = $p.Name; PID = $p.ProcessId; Seit = Format-Datum $p.CreationDate }
    })
    $skripte = @($skripte | Where-Object { $_.Skript -notmatch '(?i)\\Systemdiagnose\\' })
    if ($skripte.Count -gt 0) {
        $doppelt = @($skripte | Where-Object { $_.Skript -notmatch '^\(' } | Group-Object Skript | Where-Object { $_.Count -gt 1 })
        if ($doppelt.Count) {
            $loesungen = @($doppelt | ForEach-Object { New-Loesung -Id 'DoppeltesSkriptBeenden' -Titel "Doppelte Instanz von $(Split-Path -Leaf $_.Name) beenden (die neueste bleibt)" -Umkehrbar $false -Parameter @{ Skript = $_.Name } })
            Add-Finding 'Skripte' $B 'Hintergrund-Skripte' 'Warnung' "Läuft doppelt: $(@($doppelt.Name | ForEach-Object { Split-Path -Leaf $_ }) -join ', ')" -Ursache ('Zwei Instanzen desselben Skripts arbeiten gegeneinander' + $(if ($Script:IstEigenerPc) { ' - so entstand im September das Flackern bei The Force Unleashed II' } else { '' }) + '. Meist ist eine Instanz von einem früheren Start übrig geblieben.') -Massnahme 'Die ältere Instanz beenden, die neueste weiterlaufen lassen.' -Empfehlung 'Beheben' -Details @($skripte) -Loesungen $loesungen
        } else {
            Add-Finding 'Skripte' $B 'Hintergrund-Skripte' 'OK' "$($skripte.Count) Skript(e) im Hintergrund, keines doppelt" -Details @($skripte)
        }
    }

    $prog = @(Get-Data 'Uninstall')
    if ($prog.Count -gt 0) {
        $ohne = @($prog | Where-Object { -not $_.Deinstallation })
        $zeilen = $prog | Sort-Object Name -Unique | Select-Object Name, Version, Hersteller, Installiert
        $txt = "$(@($prog | Sort-Object Name -Unique).Count) Programme installiert"
        if ($ohne.Count) { $txt += ", $($ohne.Count) Einträge ohne Deinstallationsbefehl ($(@($ohne.Name) -join ', '))" }
        Add-Finding 'Programme' $B 'Installierte Programme' 'Info' $txt -Massnahme 'Nur zur Übersicht - neue und entfernte Programme stehen unter „Änderungen“.' -Empfehlung 'Behalten' -Details @($zeilen) -Loesungen @(New-SeitenLoesung -Titel 'Installierte Apps öffnen' -Seite 'ms-settings:appsfeatures')
    }
}

function Test-Einstellungen {
    $B = 'Einstellungen (PC-Profil)'
    $s = Get-Data 'Settings'
    $f = '<fehlt>'

    $plan = Get-Data 'PowerScheme'
    if ($plan) {
        $name = if ($plan -match '\((.+)\)') { $Matches[1] } else { Format-Kurz $plan 80 }
        if (-not $Script:Profil.EnergieplanMuster) { Add-Finding 'Energieplan' $B 'Energieplan' 'Info' $name -Massnahme 'Noch kein Sollwert - wird beim ersten Lauf auf diesem PC eingelernt.' -Empfehlung 'Behalten' }
        elseif ($name -match $Script:Profil.EnergieplanMuster) { Add-Finding 'Energieplan' $B 'Energieplan' 'OK' $name }
        elseif (-not $Script:IstEigenerPc) { Add-Finding 'Energieplan' $B 'Energieplan' 'Warnung' $name -Ursache 'Der aktive Energieplan ist nicht mehr der eingelernte. Treiber-Updates und manche Tools stellen ihn gelegentlich um.' -Massnahme 'In den Energieoptionen den gewohnten Plan wieder aktivieren - oder, wenn der neue Plan gewollt ist, die Sollwerte in den Einstellungen der App neu einlernen.' -Empfehlung 'Prüfen' -Loesungen @(New-SeitenLoesung -Titel 'Energieoptionen öffnen' -Seite 'ms-settings:powersleep') }
        else {
            # Eingelernter Plan (exakter Name) hat Vorrang vor dem eingebauten Muster
            $soll = if ($Script:Profil['EnergieplanName']) { [string]$Script:Profil['EnergieplanName'] } else { 'AMD Ryzen High Performance' }
            $param = if ($Script:Profil['EnergieplanName']) { @{ Name = $soll } } else { @{ Muster = $Script:Profil.EnergieplanMuster } }
            # Texte per -f aus einfachen Strings: „…“ in "..." beendet in PowerShell den String
            Add-Finding 'Energieplan' $B 'Energieplan' 'Warnung' $name -Ursache ('Sollwert ist der Plan „{0}“. Chipsatz-Updates und manche Spiele-Tools setzen den Plan gelegentlich zurück.' -f $soll) -Massnahme ('Den Plan „{0}“ wieder aktivieren (umkehrbar).' -f $soll) -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'EnergieplanSetzen' -Titel ('Energieplan „{0}“ aktivieren' -f $soll) -Parameter $param)
        }
    }

    if ($s) {
        $hb = [string]$s['Schnellstart (HiberbootEnabled)']
        if ($hb -eq '0') { Add-Finding 'Schnellstart' $B 'Schnellstart' 'OK' 'aus' }
        elseif ($hb -eq '1' -and -not $Script:IstEigenerPc) { Add-Finding 'Schnellstart' $B 'Schnellstart' 'Info' 'an' -Massnahme 'Windows-Standard. Bei Treiber- oder Laufwerksproblemen nach dem Hochfahren lohnt es sich, ihn testweise abzuschalten.' -Empfehlung 'Behalten' }
        elseif ($hb -eq '1') { Add-Finding 'Schnellstart' $B 'Schnellstart' 'Warnung' 'an' -Ursache 'Schnellstart wurde bewusst abgeschaltet: Er war Mitverdächtiger bei den Laufwerks-Erkennungsproblemen und hält Treiberzustände über das Herunterfahren hinweg fest. Ein Windows-Update hat ihn vermutlich wieder eingeschaltet.' -Massnahme 'Schnellstart wieder ausschalten (umkehrbar).' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'SchnellstartAus' -Titel 'Schnellstart ausschalten') }
        else { Add-Finding 'Schnellstart' $B 'Schnellstart' 'Info' 'Wert fehlt (Windows-Standard = an)' -Massnahme 'Schnellstart ausschalten, um dem Sollwert zu entsprechen.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'SchnellstartAus' -Titel 'Schnellstart ausschalten') }

        $hib = [string]$s['Ruhezustand (HibernateEnabled)']
        Add-Finding 'Ruhezustand' $B 'Ruhezustand' 'Info' $(if ($hib -eq '0') { 'aus' } elseif ($hib -eq '1') { 'an' } else { 'Wert fehlt' }) -Massnahme 'Nur zur Übersicht. Aus spart Platz auf C: (hiberfil.sys).' -Empfehlung 'Behalten'

        $dvr1 = [string]$s['GameDVR Aufnahme (AppCaptureEnabled)']; $dvr2 = [string]$s['GameDVR (GameDVR_Enabled)']
        if ($dvr1 -eq '0' -and $dvr2 -eq '0') { Add-Finding 'GameDVR' $B 'Xbox-Aufnahme (GameDVR)' 'OK' 'aus' }
        else { Add-Finding 'GameDVR' $B 'Xbox-Aufnahme (GameDVR)' 'Warnung' "Aufnahme=$dvr1, GameDVR=$dvr2" -Ursache 'Die Hintergrundaufnahme ist wieder an - meist nach einem Update der Xbox-Game-Bar.' -Massnahme 'Wieder ausschalten (umkehrbar). Sie kostet in Spielen Leistung.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'GameDvrAus' -Titel 'Xbox-Hintergrundaufnahme ausschalten') }

        $dl1 = $s['Dynamic Lighting (AmbientLighting)']; $dl2 = $s['Dynamic Lighting (Vordergrund-App)']
        if ([string]$dl1 -eq '0' -and [string]$dl2 -eq '0') { Add-Finding 'DynamicLighting' $B 'Windows Dynamic Lighting' 'OK' 'aus' -Hinweis $(if ($Script:IstEigenerPc) { 'Muss aus bleiben, sonst kämpft Windows mit G HUB um das Logitech-RGB.' } else { '' }) }
        elseif (-not $Script:IstEigenerPc) { Add-Finding 'DynamicLighting' $B 'Windows Dynamic Lighting' 'Info' $(if ($dl1 -eq $f) { 'an (Windows-Standard)' } else { "Ambient=$dl1, Vordergrund-App=$dl2" }) -Massnahme 'Nur zur Übersicht. Leuchtet RGB-Zubehör nicht wie in der Hersteller-Software eingestellt, Dynamic Lighting in den Windows-Einstellungen ausschalten.' -Empfehlung 'Behalten' -Loesungen @(New-SeitenLoesung -Titel 'Dynamic Lighting öffnen' -Seite 'ms-settings:personalization-lighting') }
        else {
            $ursache = if ($dl1 -eq $f) { 'Der Einstellungs-Eintrag fehlt - und ein fehlender Eintrag bedeutet bei Windows nicht „aus“, sondern „Standard“, und der ist hier AN. Folge: Windows beansprucht Maus und Headset, G HUB verliert die Kontrolle und das RGB leuchtet nur kurz beim Start.' } else { 'Dynamic Lighting ist eingeschaltet. Windows beansprucht dann Maus und Headset, G HUB verliert die Kontrolle über das Logitech-RGB.' }
            Add-Finding 'DynamicLighting' $B 'Windows Dynamic Lighting' 'Warnung' $(if ($dl1 -eq $f) { 'Schlüssel fehlt - Windows-Standard ist AN' } else { "Ambient=$dl1, Vordergrund-App=$dl2" }) -Ursache $ursache -Massnahme 'Dynamic Lighting ausschalten (umkehrbar). Danach G HUB einmal neu starten.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'DynamicLightingAus' -Titel 'Dynamic Lighting ausschalten und LampArray-Dienst neu starten')
        }

        $gm = [string]$s['Spielmodus (AutoGameModeEnabled)']
        Add-Finding 'Spielmodus' $B 'Spielmodus' 'Info' $(if ($gm -eq '0') { 'aus' } elseif ($gm -eq '1') { 'an' } else { 'Wert fehlt (Windows-Standard = an)' }) -Massnahme 'Nur zur Übersicht.' -Empfehlung 'Behalten'

        $hags = [string]$s['GPU-Planung HAGS (HwSchMode)']
        Add-Finding 'HAGS' $B 'Hardwarebeschleunigte GPU-Planung' 'Info' $(if ($hags -eq '2') { 'an' } elseif ($hags -eq '1') { 'aus' } else { 'Wert fehlt (Windows-/Treiber-Standard)' }) -Massnahme 'Nur zur Übersicht.' -Empfehlung 'Behalten'

        Add-Finding 'DirectX' $B 'DirectX-Grafikeinstellungen' 'Info' ([string]$s['DirectX global']) -Hinweis 'SwapEffectUpgradeEnable = Optimierungen für Spiele im Fenstermodus, AutoHDREnable = Auto HDR global.' -Massnahme 'Nur zur Übersicht.' -Empfehlung 'Behalten'

        $cd = [string]$s['Absturzabbild-Typ (CrashDumpEnabled)']
        $cdText = switch ($cd) { '0' { 'keine' } '1' { 'vollständig' } '2' { 'Kernel' } '3' { 'klein (Minidump)' } '7' { 'automatisch' } default { $cd } }
        if ($cd -eq '0') { Add-Finding 'DumpTyp' $B 'Absturzabbilder bei Bluescreens' 'Warnung' 'abgeschaltet' -Ursache 'Ohne Abbild lassen sich Bluescreens später nicht auswerten.' -Massnahme 'Auf „klein (Minidump)“ stellen - braucht kaum Platz.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'AbbildTypKlein' -Titel 'Kleine Absturzabbilder (Minidump) einschalten') }
        else { Add-Finding 'DumpTyp' $B 'Absturzabbilder bei Bluescreens' 'Info' $cdText -Massnahme 'Nur zur Übersicht.' -Empfehlung 'Behalten' }
    }

    $pf = @(Get-Data 'PageFile')
    $cs = Get-Data 'ComputerSystem'
    if ($pf.Count -gt 0) {
        $txt = (@($pf | ForEach-Object { '{0}: {1:N1} GB zugewiesen, Spitze {2:N1} GB' -f $_.Name, ($_.AllocatedBaseSize / 1024), ($_.PeakUsage / 1024) }) -join '; ')
        if ($cs) { $txt += $(if ($cs.AutomaticManagedPagefile) { ', automatisch verwaltet' } else { ', manuell festgelegt' }) }
        Add-Finding 'Auslagerung' $B 'Auslagerungsdatei' 'Info' $txt -Massnahme 'Bei 16 GB RAM und großen Spielen automatisch verwaltet lassen.' -Empfehlung 'Behalten'
    }

    # --- Monitore gegen das Profil ---
    $mons = @(Get-Data 'Display')
    if ($mons.Count -gt 0) {
        foreach ($soll in $Script:Profil.Monitore) {
            $ist = $mons | Where-Object { $_.Edid -eq $soll.Edid } | Select-Object -First 1
            if (-not $ist) { Add-Finding 'Monitor' $B "Monitor $($soll.Name)" 'Warnung' 'nicht aktiv oder nicht angeschlossen' -Ursache 'Der Monitor wird nicht erkannt - ausgeschaltet, im Standby oder Kabel locker.' -Massnahme 'Monitor einschalten und Kabel prüfen; danach erneut prüfen.' -Empfehlung 'Prüfen'; continue }
            $probleme = @()
            # Bietet der Monitor den Soll-Modus ueberhaupt an? Fehlt er, ist meist der HDMI-2.1-Handshake weg (Philips am 25.10.)
            $sollModus = $null
            if ($soll.Aufloesung) {
                $angeboten = @(@($ist.Modi) | Where-Object { $_ -match '^(\d+x\d+)@(\d+)$' -and $Matches[1] -eq $soll.Aufloesung -and [Math]::Abs([double]$Matches[2] - [double]$soll.Hz) -le 1 })
                $sollModus = ($angeboten.Count -gt 0)
                if ($ist.Aufloesung -ne $soll.Aufloesung) { $probleme += "läuft mit $($ist.Aufloesung) statt $($soll.Aufloesung)" }
                if (@($ist.Modi).Count -and -not $sollModus) { $probleme += "$($soll.Aufloesung) @ $($soll.Hz) Hz wird gerade gar nicht angeboten" }
            }
            if ([Math]::Abs([double]$ist.Hz - [double]$soll.Hz) -gt 1) { $probleme += "läuft mit $($ist.Hz) Hz statt $($soll.Hz) Hz" }
            if ($null -ne $soll.Hdr -and [bool]$ist.HdrAktiv -ne [bool]$soll.Hdr) { $probleme += "HDR ist $(if ($ist.HdrAktiv) { 'an' } else { 'aus' }) (Sollwert: $(if ($soll.Hdr) { 'an' } else { 'aus' }))" }
            $txt = '{0} @ {1} Hz, HDR {2}, {3} mit {4} Bit{5}' -f $ist.Aufloesung, $ist.Hz, $(if ($ist.HdrAktiv) { 'an' } else { 'aus' }), $ist.Farbformat, $ist.BitProKanal, $(if ($ist.Primaer) { ', Hauptbildschirm' } else { '' })
            if ($probleme.Count -and $sollModus -eq $false) {
                Add-Finding 'Monitor' $B "Monitor $($soll.Name)" 'Warnung' (($probleme -join '; ') + " - $txt") -Ursache 'Der Monitor meldet den Soll-Modus gerade nicht als verfügbar. Beim Fernseher heißt das fast immer: Die schnelle HDMI-2.1-Verbindung ist beim Aushandeln abgebrochen, Windows bekommt nur noch die langsameren Modi angeboten.' -Massnahme '1. Grafiktreiber neu verbinden: Win+Strg+Umschalt+B. 2. Hilft das nicht: Fernseher ganz aus- und wieder einschalten (nicht nur Standby) oder das HDMI-Kabel kurz abziehen. 3. Danach die Auflösung in den Anzeige-Einstellungen prüfen und hier „Aktualisieren“.' -Empfehlung 'Beheben' -Loesungen @(New-SeitenLoesung -Titel 'Anzeige-Einstellungen öffnen' -Seite 'ms-settings:display')
            } elseif ($probleme.Count) {
                Add-Finding 'Monitor' $B "Monitor $($soll.Name)" 'Warnung' (($probleme -join '; ') + " - $txt") -Ursache 'Die Anzeige-Einstellungen weichen vom Sollwert ab - etwa nach einem Spiel, das die Auflösung umgestellt hat, nach einem Treiber-Update oder einem neuen Kabel-Handshake mit dem Fernseher.' -Massnahme 'In den Anzeige-Einstellungen den Monitor wählen > Erweiterte Anzeige > Bildwiederholrate bzw. HDR einstellen. War die Änderung gewollt, den Sollwert im PC-Profil (oben im Skript) anpassen.' -Empfehlung 'Beheben' -Loesungen @(New-SeitenLoesung -Titel 'Anzeige-Einstellungen öffnen' -Seite 'ms-settings:display')
            } else {
                Add-Finding 'Monitor' $B "Monitor $($soll.Name)" 'OK' $txt
            }
        }
        $fremde = @($mons | Where-Object { $e = $_.Edid; -not @($Script:Profil.Monitore | Where-Object { $_.Edid -eq $e }).Count })
        foreach ($m in $fremde) {
            if ($Script:ProfilQuelle -eq 'neu') { Add-Finding 'Monitor' $B "Monitor $($m.Name)" 'OK' "$($m.Aufloesung) @ $($m.Hz) Hz, HDR $(if ($m.HdrAktiv) { 'an' } else { 'aus' })" -Hinweis 'Wird beim ersten Lauf auf diesem PC als Sollwert eingelernt.' }
            else { Add-Finding 'Monitor' $B "Monitor $($m.Name)" 'Info' "$($m.Aufloesung) @ $($m.Hz) Hz (nicht im PC-Profil)" -Massnahme 'Bleibt der Monitor dauerhaft angeschlossen: in den Einstellungen der App „Sollwerte neu einlernen“.' -Empfehlung 'Behalten' }
        }
    }

    # --- Mauszeiger eingesperrt? (ClipCursor-Leiche nach Spielen) ---
    $c = @(Get-Data 'CursorClip')
    if ($c.Count -ge 8) {
        $frei = ($c[0] -le $c[4]) -and ($c[1] -le $c[5]) -and ($c[2] -ge ($c[4] + $c[6])) -and ($c[3] -ge ($c[5] + $c[7]))
        if ($frei) { Add-Finding 'ClipCursor' $B 'Mauszeiger-Bereich' 'OK' 'frei beweglich auf allen Bildschirmen' }
        else { Add-Finding 'ClipCursor' $B 'Mauszeiger-Bereich' 'Warnung' "eingesperrt auf ($($c[0]),$($c[1])) bis ($($c[2]),$($c[3]))" -Ursache 'Ein Spiel hat den Mauszeiger auf seinen Bereich eingesperrt und die Sperre beim Beenden nicht aufgehoben (so geschehen bei Battlefront II).' -Massnahme 'Sperre aufheben - wirkt sofort, ohne Neustart.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'MauszeigerFreigeben' -Titel 'Mauszeiger-Sperre aufheben' -Art 'Sofort') }
    }
}

function Test-Sicherheit {
    $B = 'Sicherheit & Updates'
    $s = Get-Data 'Settings'

    if ($s) {
        $sb = [string]$s['Secure Boot (UEFISecureBootEnabled)']
        if ($sb -eq '1') { Add-Finding 'SecureBoot' $B 'Secure Boot' 'OK' 'an' }
        elseif ($sb -eq '0') { Add-Finding 'SecureBoot' $B 'Secure Boot' 'Warnung' 'aus' -Ursache 'Secure Boot wird nach einem BIOS-Reset oder -Update oft zurückgesetzt.' -Massnahme 'Im BIOS (Entf beim Start): Boot > CSM „Disabled“, Secure Boot > OS Type „Windows UEFI mode“, „Install Default Secure Boot Keys“. Danach DOCP prüfen.' -Empfehlung 'Beheben' }
        else { Add-Finding 'SecureBoot' $B 'Secure Boot' 'Info' 'nicht ermittelbar' }

        $uac = [string]$s['Benutzerkontensteuerung (EnableLUA)']
        if ($uac -eq '0') { Add-Finding 'UAC' $B 'Benutzerkontensteuerung' 'Warnung' 'aus' -Ursache 'Ohne Benutzerkontensteuerung läuft jedes Programm mit vollen Adminrechten.' -Massnahme 'Wieder einschalten (wirksam nach einem Neustart).' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'UacEinschalten' -Titel 'Benutzerkontensteuerung einschalten (Neustart nötig)') }
        elseif ([string]$s['UAC-Nachfrage Admins (ConsentPrompt)'] -eq '0') {
            Add-Finding 'UAC' $B 'Benutzerkontensteuerung' 'Warnung' 'an, fragt aber nie nach (Programme bekommen Adminrechte ohne Rückfrage)' -Ursache 'Die Stufe „Nie benachrichtigen“ ist eingestellt (ConsentPromptBehaviorAdmin = 0). Dann kann jedes Programm, das du startest, still Adminrechte anfordern - auch Schadsoftware.' -Massnahme 'Auf die Windows-Standardstufe zurückstellen: Nachfrage nur, wenn Programme Änderungen vornehmen wollen, auf dem abgedunkelten sicheren Desktop. Spiele und Programme laufen genauso schnell wie vorher.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'UacNachfrageAn' -Titel 'Benutzerkontensteuerung auf Standardstufe (mit Nachfrage)')
        }
        else { Add-Finding 'UAC' $B 'Benutzerkontensteuerung' 'OK' 'an, mit Nachfrage' }

        $bl = [string]$s['Treiber-Sperrliste (VulnerableDriver)']
        Add-Finding 'Sperrliste' $B 'Sperrliste für anfällige Treiber' 'Info' $(if ($bl -eq '1') { 'an' } elseif ($bl -eq '0') { 'aus' } else { 'Windows-Standard' }) -Massnahme 'Nur zur Übersicht.' -Empfehlung 'Behalten'
    }

    if ($Script:AdminModus) {
        $tpm = Get-Data 'Tpm'
        if (Test-QuelleFehlt 'Tpm' 'TPM' $B 'TPM') {
        } elseif (-not $tpm -or $null -eq $tpm.TpmPresent) {
            Add-Finding 'TPM' $B 'TPM' 'Info' 'nicht lesbar (Windows hat keine Werte geliefert)'
        } else {
            $ok = $tpm.TpmPresent -and $tpm.TpmReady -and $tpm.TpmEnabled -and $tpm.TpmActivated
            $txt = 'vorhanden {0}, bereit {1}, aktiviert {2}' -f $(if ($tpm.TpmPresent) { 'ja' } else { 'nein' }), $(if ($tpm.TpmReady) { 'ja' } else { 'nein' }), $(if ($tpm.TpmActivated) { 'ja' } else { 'nein' })
            if ($ok) { Add-Finding 'TPM' $B 'TPM 2.0' 'OK' $txt }
            else { Add-Finding 'TPM' $B 'TPM 2.0' 'Warnung' $txt -Ursache 'Das fTPM wird nach einem BIOS-Reset manchmal abgeschaltet.' -Massnahme 'Im BIOS unter Advanced > AMD fTPM configuration das fTPM einschalten.' -Empfehlung 'Beheben' }
        }
    }

    $dg = Get-Data 'DeviceGuard'
    if ($dg) {
        $vbs = switch ([int]$dg.VirtualizationBasedSecurityStatus) { 0 { 'aus' } 1 { 'aktiviert, läuft nicht' } 2 { 'läuft' } default { '?' } }
        $hvci = if (@($dg.SecurityServicesRunning) -contains 2) { 'an' } else { 'aus' }
        Add-Finding 'VBS' $B 'Virtualisierungsschutz / Speicherintegrität' 'Info' "VBS $vbs, Speicherintegrität $hvci" -Massnahme 'Aus = etwas mehr Spieleleistung, passt zum Ziel maximale Leistung. Mehrere Hardwarezugriffs-Treiber wären mit Speicherintegrität ohnehin nicht kompatibel.' -Empfehlung 'Behalten'
    }

    $def = Get-Data 'Defender'
    if ($def) {
        $probleme = @()
        if (-not $def.AMServiceEnabled -or -not $def.AntivirusEnabled) { $probleme += 'Virenschutz aus' }
        if (-not $def.RealTimeProtectionEnabled) { $probleme += 'Echtzeitschutz aus' }
        $alt = ($def.AntivirusSignatureAge -gt 3)
        if ($alt) { $probleme += "Signaturen $($def.AntivirusSignatureAge) Tage alt" }
        $txt = "Echtzeitschutz $(if ($def.RealTimeProtectionEnabled) { 'an' } else { 'aus' }), Signaturen vom $(Format-Datum $def.AntivirusSignatureLastUpdated)"
        if ($probleme.Count) {
            $loesungen = @()
            if ($alt) { $loesungen += New-Loesung -Id 'DefenderSignaturen' -Umkehrbar $false -Titel 'Virensignaturen jetzt aktualisieren' }
            $loesungen += New-SeitenLoesung -Titel 'Windows-Sicherheit öffnen' -Seite 'windowsdefender:'
            Add-Finding 'Defender' $B 'Microsoft Defender' 'Warnung' (($probleme -join ', ') + " - $txt") -Ursache 'Der Virenschutz ist eingeschränkt - entweder abgeschaltet oder seit Tagen ohne neue Signaturen (z. B. wenn Windows Update hängt).' -Massnahme 'Signaturen aktualisieren bzw. den Echtzeitschutz in „Windows-Sicherheit“ wieder einschalten.' -Empfehlung 'Beheben' -Loesungen $loesungen
        } else { Add-Finding 'Defender' $B 'Microsoft Defender' 'OK' $txt }
    } else {
        Add-Finding 'Defender' $B 'Microsoft Defender' 'Info' 'Status nicht lesbar' -Ursache 'Entweder ist ein anderer Virenschutz aktiv, oder Windows hat die Abfrage abgelehnt.'
    }

    $fw = @(Get-Data 'Firewall')
    if ($fw.Count -gt 0) {
        $aus = @($fw | Where-Object { -not $_.Enabled })
        if ($aus.Count) { Add-Finding 'Firewall' $B 'Firewall' 'Warnung' ('Aus für: ' + (@($aus.Name) -join ', ')) -Ursache 'Die Windows-Firewall ist für mindestens ein Netzwerkprofil abgeschaltet - manchmal von Spielen, VPNs oder Tuning-Tools.' -Massnahme 'Wieder einschalten (umkehrbar).' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'FirewallEinschalten' -Titel ('Firewall einschalten für ' + (@($aus.Name) -join ', ')) -Parameter @{ Profile = @($aus.Name) }) }
        else { Add-Finding 'Firewall' $B 'Firewall' 'OK' ('an für ' + (@($fw.Name) -join ', ')) }
    }

    # Erfolgreiche Installationen stehen im ohnehin gelesenen System-Protokoll (Ereignis 19)
    $erfolg = @(@(Get-Data 'EventsSystem') | Where-Object { $_.ProviderName -eq 'Microsoft-Windows-WindowsUpdateClient' -and $_.Id -eq 19 })
    $winUpd = @(foreach ($e in $erfolg) {
        $t = Get-EventText $e
        if ($t -match '(?i)Defender|Security Intelligence|KB2267602') { continue }
        $titel = if ($t -match '(?s)(?:installiert\.|following update:)\s*(.+)$') { ($Matches[1] -replace '\s+', ' ').Trim() } else { Format-Kurz $t 120 }
        if ($titel -match '^9[A-Z0-9]{11}-') { continue }
        [pscustomobject]@{ Update = $titel; Installiert = Format-Datum $e.TimeCreated }
    })
    if ($winUpd.Count -gt 0) {
        Add-Finding 'Updates' $B 'Windows-Updates' 'OK' "Zuletzt installiert am $($winUpd[0].Installiert): $($winUpd[0].Update)" -Hinweis "Alle Windows-, Treiber- und Office-Updates der letzten $Tage Tage (ohne Defender-Signaturen und Store-Apps) in den Details." -Details @($winUpd)
    } else {
        # Nur wenn im Zeitraum nichts installiert wurde, die (langsamere) Update-Liste lesen
        $hf = @(Get-Data 'HotFix' | Where-Object { $_.InstalledOn } | Sort-Object InstalledOn -Descending)
        if ($hf.Count -gt 0) {
            $alter = ((Get-Date) - [datetime]$hf[0].InstalledOn).Days
            $txt = "Im Zeitraum kein Update; laut Update-Liste zuletzt am $(([datetime]$hf[0].InstalledOn).ToString('dd.MM.yyyy')) (vor $alter Tagen)"
            if ($alter -gt $Script:Profil.UpdateMaxTage) { Add-Finding 'Updates' $B 'Windows-Updates' 'Warnung' $txt -Ursache 'Seit Wochen wurde kein Update installiert - Windows Update hängt oder ist pausiert.' -Massnahme 'In Windows Update nach Updates suchen und eine eventuelle Pause aufheben.' -Empfehlung 'Beheben' -Loesungen @(New-SeitenLoesung -Titel 'Windows Update öffnen' -Seite 'ms-settings:windowsupdate') }
            else { Add-Finding 'Updates' $B 'Windows-Updates' 'OK' $txt }
        }
    }

    $wu = @(@(Get-Data 'EventsSystem') | Where-Object { $_.ProviderName -eq 'Microsoft-Windows-WindowsUpdateClient' -and $_.Id -eq 20 -and -not (Get-NoiseReason $_) })
    if ($wu.Count -gt 0) {
        $zeilen = $wu | Group-Object { $t = Get-EventText $_; if ($t -match ':\s*([^:]+)$') { $Matches[1].Trim() } else { Format-Kurz $t 100 } } | ForEach-Object { [pscustomobject]@{ Update = $_.Name; Fehlversuche = $_.Count; Zuletzt = Format-Datum $_.Group[0].TimeCreated } }
        $oft = @($zeilen | Where-Object { $_.Fehlversuche -ge 3 })
        Add-Finding 'UpdateFehler' $B 'Fehlgeschlagene Updates' $(if ($oft.Count) { 'Warnung' } else { 'Info' }) "$($wu.Count) Fehlversuche" -Ursache 'Einzelne Fehlversuche entstehen, wenn ein Update während der Installation gestört wird. Scheitert dasselbe Update immer wieder, ist meist der Update-Speicher beschädigt.' -Massnahme $(if ($oft.Count) { 'In Windows Update erneut suchen; scheitert es weiter, im Gründlich-Modus DISM prüfen.' } else { 'Holt Windows Update meist von selbst nach.' }) -Empfehlung $(if ($oft.Count) { 'Beheben' } else { 'Beobachten' }) -Details @($zeilen) -Loesungen @(New-SeitenLoesung -Titel 'Windows Update öffnen' -Seite 'ms-settings:windowsupdate') -Zuletzt (Get-Neuestes @($wu.TimeCreated))
    } elseif (-not $Script:QuellenFehler.ContainsKey('EventsSystem')) {
        Add-Finding 'UpdateFehler' $B 'Fehlgeschlagene Updates' 'OK' 'Keine'
    }

    $akt = @(Get-Data 'Activation')
    if (Test-QuelleFehlt 'Activation' 'Aktivierung' $B 'Windows-Aktivierung') {
    } elseif ($akt.Count -gt 0) {
        $ok = @($akt | Where-Object { $_.LicenseStatus -eq 1 }).Count -gt 0
        if ($ok) { Add-Finding 'Aktivierung' $B 'Windows-Aktivierung' 'OK' 'aktiviert' }
        else { Add-Finding 'Aktivierung' $B 'Windows-Aktivierung' 'Warnung' 'nicht aktiviert' -Ursache 'Die digitale Lizenz wurde nicht erkannt - etwa nach einem Hardwaretausch (Mainboard).' -Massnahme 'Unter Einstellungen > Aktivierung die Problembehandlung starten.' -Empfehlung 'Beheben' -Loesungen @(New-SeitenLoesung -Titel 'Aktivierung öffnen' -Seite 'ms-settings:activation') }
    }
}

function Test-SystemBasis {
    $B = 'System'
    $os = Get-Data 'OS'
    $wv = Get-Data 'WinVersion'
    if ($os -and $wv) {
        Add-Finding 'Windows' $B 'Windows' 'Info' ("{0}, Version {1}, Build {2}.{3}" -f $os.Caption, $wv['DisplayVersion'], $wv['CurrentBuild'], $wv['UBR']) -Hinweis 'In der Registry steht weiterhin „Windows 10“ als Produktname - ein bekannter Microsoft-Rest, kein Fehler.' -Massnahme 'Nur zur Übersicht.' -Empfehlung 'Behalten'
        $lauf = (Get-Date) - $os.LastBootUpTime
        $txt = '{0} Tage {1} Std. (Start: {2})' -f [int][Math]::Floor($lauf.TotalDays), $lauf.Hours, (Format-Datum $os.LastBootUpTime)
        if ($lauf.TotalDays -gt 14) { Add-Finding 'Laufzeit' $B 'Seit dem letzten Start' 'Info' $txt -Ursache 'Der PC läuft seit über zwei Wochen ohne Neustart.' -Massnahme 'Gelegentlich neu starten - dabei werden Updates abgeschlossen und Treiber frisch geladen.' -Empfehlung 'Beheben' }
        else { Add-Finding 'Laufzeit' $B 'Seit dem letzten Start' 'OK' $txt }
    }

    $pr = Get-Data 'PendingReboot'
    if ($pr) {
        if (@($pr.Gruende).Count) { Add-Finding 'Neustart' $B 'Ausstehender Neustart' 'Warnung' ('Neustart nötig wegen: ' + (@($pr.Gruende) -join ', ')) -Ursache 'Installierte Updates oder Komponenten werden erst nach dem Neustart wirksam.' -Massnahme 'Bei Gelegenheit neu starten.' -Empfehlung 'Beheben' }
        elseif ($pr.DateiOperationen) { Add-Finding 'Neustart' $B 'Ausstehender Neustart' 'OK' 'Kein Neustart nötig (nur vorgemerkte Datei-Aufräumarbeiten)' }
        else { Add-Finding 'Neustart' $B 'Ausstehender Neustart' 'OK' 'Kein Neustart nötig' }
    }

    $ins = Get-Data 'Insider'
    if ($ins -and ($ins['BranchName'] -or [string]$ins['EnablePreviewBuilds'] -eq '1')) { Add-Finding 'Insider' $B 'Insider-Programm' 'Info' "aktiv (Kanal: $($ins['BranchName']))" -Ursache 'Vorabversionen sind weniger stabil.' -Massnahme 'Wenn nicht gewollt: Einstellungen > Windows Update > Windows-Insider-Programm verlassen.' -Empfehlung 'Prüfen' -Loesungen @(New-SeitenLoesung -Titel 'Insider-Einstellungen öffnen' -Seite 'ms-settings:windowsinsider') }
    else { Add-Finding 'Insider' $B 'Insider-Programm' 'OK' 'nicht aktiv' }

    $bios = Get-Data 'BIOS'
    if ($bios) {
        Add-Finding 'BIOS' $B 'BIOS' 'Info' ("Version {0} vom {1}" -f $bios.SMBIOSBIOSVersion, $(if ($bios.ReleaseDate) { ([datetime]$bios.ReleaseDate).ToString('dd.MM.yyyy') } else { '?' })) -Massnahme 'BIOS-Updates nur bei konkretem Problem: Sie setzen Secure Boot, CSM und DOCP zurück.' -Empfehlung 'Behalten'
    }

    $cpu = @(Get-Data 'CPU')
    $csys = Get-Data 'ComputerSystem'
    if ($cpu.Count -gt 0) {
        $ramGb = if ($csys) { [Math]::Round($csys.TotalPhysicalMemory / 1GB, 1) } else { '?' }
        Add-Finding 'Prozessor' $B 'Prozessor' 'Info' ("{0}, {1} Kerne / {2} Threads, {3} GB nutzbarer RAM" -f ([string]$cpu[0].Name).Trim(), $cpu[0].NumberOfCores, $cpu[0].NumberOfLogicalProcessors, $ramGb) -Hinweis 'Temperaturen kann Windows nicht zuverlässig auslesen - dafür HWiNFO nutzen.' -Massnahme 'Nur zur Übersicht.' -Empfehlung 'Behalten'
    }

    # --- Zeitsynchronisation ---
    $zeit = @(@(Get-Data 'EventsSystem') | Where-Object { $_.ProviderName -eq 'Microsoft-Windows-Time-Service' -and ($_.Id -eq 35 -or $_.Id -eq 37) })
    $quelle = Get-Data 'TimeStatus'
    $q = if ($quelle -and $quelle -match '(?im)^(?:Quelle|Source)\s*:\s*(.+)$') { $Matches[1].Trim() } else { '?' }
    if ($zeit.Count -gt 0) {
        $alter = ((Get-Date) - $zeit[0].TimeCreated).TotalDays
        if ($alter -gt 7) { Add-Finding 'Zeit' $B 'Zeitsynchronisation' 'Warnung' "Letzter Abgleich am $(Format-Datum $zeit[0].TimeCreated)" -Ursache 'Der Zeitdienst gleicht seit über einer Woche nicht mehr ab.' -Massnahme 'Jetzt synchronisieren.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'ZeitSynchronisieren' -Umkehrbar $false -Titel 'Uhrzeit jetzt synchronisieren') }
        else { Add-Finding 'Zeit' $B 'Zeitsynchronisation' 'OK' "Letzter Abgleich am $(Format-Datum $zeit[0].TimeCreated)" -Hinweis ('Aktuelle Quelle: ' + $q + ' („Local CMOS Clock“ ist nur eine Momentaufnahme zwischen zwei Abgleichen).') }
    } elseif (-not $Script:QuellenFehler.ContainsKey('EventsSystem')) {
        Add-Finding 'Zeit' $B 'Zeitsynchronisation' 'Warnung' "Kein erfolgreicher Abgleich in $Tage Tagen (Quelle: $q)" -Ursache 'Der Windows-Zeitdienst hat im Zeitraum keinen Zeitserver erreicht.' -Massnahme 'Jetzt synchronisieren.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'ZeitSynchronisieren' -Umkehrbar $false -Titel 'Uhrzeit jetzt synchronisieren')
    }
}

function Test-Netzwerk {
    $B = 'Netzwerk'
    $ad = @(Get-Data 'NetAdapters')
    if ($ad.Count -gt 0) {
        $an = @($ad | Where-Object { $_.Status -eq 'Up' })
        $zeilen = $ad | Select-Object Name, InterfaceDescription, Status, LinkSpeed
        if ($an.Count -eq 0) { Add-Finding 'Netzwerk' $B 'Netzwerkadapter' 'Kritisch' 'Kein Adapter verbunden' -Ursache 'Weder WLAN noch LAN ist verbunden.' -Massnahme 'WLAN-Verbindung prüfen (Router an? Flugmodus?).' -Empfehlung 'Beheben' -Details @($zeilen) -Loesungen @(New-SeitenLoesung -Titel 'Netzwerkeinstellungen öffnen' -Seite 'ms-settings:network') }
        else { Add-Finding 'Netzwerk' $B 'Netzwerkadapter' 'OK' ('Verbunden: ' + (@($an | ForEach-Object { "$($_.Name) ($($_.LinkSpeed))" }) -join ', ')) -Hinweis 'Der LAN-Anschluss (Intel I226-V) wird bewusst nicht genutzt.' -Details @($zeilen) }
    }

    $w = Get-Data 'Wlan'
    if ($w -and $w.Count -gt 0) {
        $ssid = ($w.Keys | Where-Object { $_ -match '^SSID$' } | ForEach-Object { $w[$_] } | Select-Object -First 1)
        $sigKey = $w.Keys | Where-Object { $_ -match '^Signal' } | Select-Object -First 1
        if ($ssid -and $sigKey) {
            $sig = [int](([string]$w[$sigKey]) -replace '[^\d]', '')
            $rx = $w.Keys | Where-Object { $_ -match 'Empfangsrate|Receive rate' } | ForEach-Object { $w[$_] } | Select-Object -First 1
            $tx = $w.Keys | Where-Object { $_ -match 'bertragungsrate|Transmit rate' } | ForEach-Object { $w[$_] } | Select-Object -First 1
            $typ = $w.Keys | Where-Object { $_ -match 'Funktyp|Radio type' } | ForEach-Object { $w[$_] } | Select-Object -First 1
            $txt = "$ssid, Signal $sig %, Empfang $rx / Senden $tx MBit/s, $typ"
            if ($sig -lt 60) { Add-Finding 'WLAN' $B 'WLAN' $(if ($sig -lt 40) { 'Warnung' } else { 'Info' }) $txt -Ursache 'Das WLAN-Signal ist schwach - Entfernung, Wände oder Störungen durch andere Netze.' -Massnahme 'Router-Position oder die WLAN-Antennen am PC (Mainboard-Antenne) prüfen; ggf. 5-GHz-Band bevorzugen.' -Empfehlung 'Prüfen' }
            else { Add-Finding 'WLAN' $B 'WLAN' 'OK' $txt }
        }
    }

    $dns = Get-Data 'DnsTest'
    if ($dns) {
        if ($dns.Ok) { Add-Finding 'DNS' $B 'Namensauflösung (DNS)' 'OK' 'funktioniert' }
        else { Add-Finding 'DNS' $B 'Namensauflösung (DNS)' 'Warnung' "fehlgeschlagen: $($dns.Text)" -Ursache 'Internet-Adressen konnten nicht aufgelöst werden - Router, Internetverbindung oder ein veralteter DNS-Zwischenspeicher.' -Massnahme 'DNS-Zwischenspeicher leeren; hilft das nicht, Router neu starten.' -Empfehlung 'Beheben' -Loesungen @(New-Loesung -Id 'DnsCacheLeeren' -Umkehrbar $false -Titel 'DNS-Zwischenspeicher leeren') }
    }
    $gw = Get-Data 'GatewayTest'
    if ($gw -and $gw.Router) {
        if ($gw.Ok) { Add-Finding 'Router' $B 'Router' 'OK' "$($gw.Router) erreichbar" }
        else { Add-Finding 'Router' $B 'Router' 'Warnung' "$($gw.Router) antwortet nicht" -Ursache 'Der Router antwortet nicht auf Ping - manche Router blockieren das, sonst ist die Verbindung gestört.' -Massnahme 'Prüfen, ob das Internet funktioniert; falls nicht, Router neu starten.' -Empfehlung 'Prüfen' }
    }
}

# ============================================================================
#  7. VERGLEICH MIT DEM LETZTEN LAUF
# ============================================================================

function New-Snapshot {
    $s = [ordered]@{}
    $s['Zeit'] = $Script:StartZeit.ToString('dd.MM.yyyy HH:mm')
    $s['Maschine'] = $Script:MaschinenId
    $wv = Get-Data 'WinVersion'
    $s['Build'] = if ($wv) { '{0} {1}.{2}' -f $wv['DisplayVersion'], $wv['CurrentBuild'], $wv['UBR'] } else { '' }
    $s['Treiberpakete'] = @(@(Get-Data 'DriverStore') | ForEach-Object { '{0} {1} ({2})' -f $_.Original, $_.Version, $_.Anbieter } | Sort-Object -Unique)
    $s['Kerneltreiber'] = @(@(Get-Data 'SystemDrivers') | Where-Object { $_.Status -eq 'Running' } | ForEach-Object { $_.Name } | Sort-Object -Unique)
    # Benutzerdienste bekommen bei jeder Anmeldung ein neues Kuerzel (CDPUserSvc_ba2ccb) - als "_*" vergleichen
    $s['Dienste'] = @(@(Get-Data 'Services') | ForEach-Object { '{0} | {1}' -f ($_.Name -replace '_[0-9a-f]{4,8}$', '_*'), $_.Start } | Sort-Object -Unique)
    $s['Autostart'] = @(@(Get-Data 'Autostart') | ForEach-Object { '{0}: {1} | {2}' -f $_.Ort, $_.Name, $_.Status } | Sort-Object)
    $s['Aufgaben'] = @(@(Get-Data 'Tasks') | ForEach-Object { '{0} | {1}' -f $_.Aufgabe, $_.Status } | Sort-Object)
    $s['Programme'] = @(@(Get-Data 'Uninstall') | Sort-Object Name -Unique | ForEach-Object { '{0} | {1}' -f $_.Name, $_.Version })
    $s['Monitore'] = @(@(Get-Data 'Display') | ForEach-Object { '{0} | {1} @ {2} Hz, HDR {3}, {4} {5} Bit' -f $_.Edid, $_.Aufloesung, $_.Hz, $(if ($_.HdrAktiv) { 'an' } else { 'aus' }), $_.Farbformat, $_.BitProKanal })
    # Nicht lesbare Quellen vermerken - Compare-Snapshot vergleicht diese Bereiche dann nicht (sonst „alles entfernt“)
    $quelle = [ordered]@{ Treiberpakete = 'DriverStore'; Kerneltreiber = 'SystemDrivers'; Dienste = 'Services'; Autostart = 'Autostart'; Aufgaben = 'Tasks'; Programme = 'Uninstall'; Monitore = 'Display' }
    $s['Unlesbar'] = @(foreach ($b in $quelle.Keys) { if ($Script:QuellenFehler.ContainsKey($quelle[$b])) { $b } })
    $e = [ordered]@{}
    $st = Get-Data 'Settings'
    if ($st) { foreach ($k in $st.Keys) { $e[$k] = [string]$st[$k] } }
    $plan = Get-Data 'PowerScheme'
    if ($plan -and $plan -match '\((.+)\)') { $e['Energieplan'] = $Matches[1] }
    $s['Einstellungen'] = $e
    $l = [ordered]@{}
    foreach ($v in @(Get-Data 'Volumes' | Where-Object { $_.DriveLetter })) { $l[[string]$v.DriveLetter] = [Math]::Round([double]$v.SizeRemaining / 1GB, 1) }
    $s['FreiGB'] = $l
    return $s
}

# Letzter Schnappschuss DIESES PCs - liegen Berichte eines anderen PCs im Ordner (Ordner kopiert), werden sie uebergangen.
# Alte Dateien ohne Kennung stammen von PC 1.
function Get-LastSnapshot {
    if (-not (Test-Path -LiteralPath $Script:BerichtOrdner)) { return $null }
    foreach ($f in @(Get-ChildItem -LiteralPath $Script:BerichtOrdner -Filter 'Snapshot-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 40)) {
        try { $s = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }
        if ([string]$s.Maschine -eq $Script:MaschinenId -or (-not $s.Maschine -and $Script:IstEigenerPc)) { return $s }
    }
    return $null
}

function ConvertTo-KeyMap {
    param([object[]] $Zeilen)
    $karte = @{}
    foreach ($z in $Zeilen) {
        $t = [string]$z
        $i = $t.IndexOf(' | ')
        if ($i -lt 0) { $karte[$t] = '' } else { $karte[$t.Substring(0, $i)] = $t.Substring($i + 3) }
    }
    return $karte
}

function Compare-Snapshot {
    param($Alt, $Neu)
    if (-not $Alt) { return }
    $titel = [ordered]@{ 'Treiberpakete' = 'Treiberpaket'; 'Kerneltreiber' = 'Kernel-Treiber (laufend)'; 'Dienste' = 'Dienst'; 'Autostart' = 'Autostart'; 'Aufgaben' = 'Geplante Aufgabe'; 'Programme' = 'Programm'; 'Monitore' = 'Monitor' }
    $zeilenModus = @('Treiberpakete', 'Kerneltreiber')
    foreach ($bereich in $titel.Keys) {
        # Nur vergleichen, wenn beide Male erfolgreich gelesen. Voll -> komplett leer ist ein stiller Lesefehler
        # (pnputil & Co. liefern dann nichts), keine echte Entfernung aller Eintraege.
        if (@($Alt.Unlesbar) -contains $bereich -or @($Neu['Unlesbar']) -contains $bereich) { continue }
        $a = @($Alt.$bereich)
        $n = @($Neu[$bereich])
        if ($a.Count -gt 0 -and $n.Count -eq 0) { continue }
        if ($bereich -eq 'Dienste') { $a = @($a | ForEach-Object { [string]$_ -replace '^(.+?)_[0-9a-f]{4,8}( \| )', '${1}_*${2}' }) }
        if ($zeilenModus -contains $bereich) {
            $setA = @{}; foreach ($x in $a) { $setA[[string]$x] = $true }
            $setN = @{}; foreach ($x in $n) { $setN[[string]$x] = $true }
            foreach ($x in $setN.Keys) { if (-not $setA.ContainsKey($x)) { [pscustomobject]@{ Bereich = $titel[$bereich]; Art = 'neu'; Eintrag = $x; Vorher = ''; Nachher = '' } } }
            foreach ($x in $setA.Keys) { if (-not $setN.ContainsKey($x)) { [pscustomobject]@{ Bereich = $titel[$bereich]; Art = 'entfernt'; Eintrag = $x; Vorher = ''; Nachher = '' } } }
        } else {
            $kA = ConvertTo-KeyMap $a
            $kN = ConvertTo-KeyMap $n
            foreach ($k in $kN.Keys) {
                if (-not $kA.ContainsKey($k)) { [pscustomobject]@{ Bereich = $titel[$bereich]; Art = 'neu'; Eintrag = $k; Vorher = ''; Nachher = $kN[$k] } }
                elseif ($kA[$k] -ne $kN[$k]) { [pscustomobject]@{ Bereich = $titel[$bereich]; Art = 'geändert'; Eintrag = $k; Vorher = $kA[$k]; Nachher = $kN[$k] } }
            }
            foreach ($k in $kA.Keys) { if (-not $kN.ContainsKey($k)) { [pscustomobject]@{ Bereich = $titel[$bereich]; Art = 'entfernt'; Eintrag = $k; Vorher = $kA[$k]; Nachher = '' } } }
        }
    }
    if ([string]$Alt.Build -ne [string]$Neu['Build']) { [pscustomobject]@{ Bereich = 'Windows'; Art = 'geändert'; Eintrag = 'Build'; Vorher = [string]$Alt.Build; Nachher = [string]$Neu['Build'] } }
    if ($Alt.Einstellungen) {
        foreach ($p in $Alt.Einstellungen.PSObject.Properties) {
            if ($Neu['Einstellungen'].Contains($p.Name) -and [string]$p.Value -ne [string]$Neu['Einstellungen'][$p.Name]) {
                [pscustomobject]@{ Bereich = 'Einstellung'; Art = 'geändert'; Eintrag = $p.Name; Vorher = [string]$p.Value; Nachher = [string]$Neu['Einstellungen'][$p.Name] }
            }
        }
    }
    if ($Alt.FreiGB) {
        foreach ($p in $Alt.FreiGB.PSObject.Properties) {
            if ($Neu['FreiGB'].Contains($p.Name)) {
                $delta = [double]$Neu['FreiGB'][$p.Name] - [double]$p.Value
                if ([Math]::Abs($delta) -ge 2) { [pscustomobject]@{ Bereich = 'Speicherplatz'; Art = $(if ($delta -lt 0) { 'weniger frei' } else { 'mehr frei' }); Eintrag = "Laufwerk $($p.Name):"; Vorher = "$($p.Value) GB"; Nachher = "$($Neu['FreiGB'][$p.Name]) GB" } }
            }
        }
    }
}

# ============================================================================
#  8. AUSGABE: HTML-BERICHT, BEFUNDDATEI (JSON) UND KONSOLE
# ============================================================================

$Script:BerichtCss = @'
:root{--bg:#f5f6f8;--card:#fff;--card2:#f8f9fb;--text:#1c2028;--muted:#5d6876;--line:#e2e6eb;--ok:#1d8a4b;--info:#2a68cc;--warn:#b86f00;--crit:#c52a2a;--okbg:#e6f5ec;--infobg:#e7effc;--warnbg:#fff3dc;--critbg:#fde7e7;--accent:#2a68cc}
@media (prefers-color-scheme: dark){:root{--bg:#13161b;--card:#1c2027;--card2:#22262e;--text:#e5e8ed;--muted:#99a4b2;--line:#2c323b;--ok:#4cc27c;--info:#6aa1ff;--warn:#f0b13c;--crit:#ff6b6b;--okbg:#17301f;--infobg:#16233a;--warnbg:#382a11;--critbg:#3b1919;--accent:#6aa1ff}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:15px/1.55 "Segoe UI",system-ui,sans-serif}
main{max-width:1180px;margin:0 auto;padding:28px 16px 72px}
h1{font-size:27px;margin:0 0 4px}h2{font-size:19px;margin:36px 0 12px;padding-bottom:6px;border-bottom:1px solid var(--line)}
.meta{color:var(--muted);font-size:13.5px}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:20px 0 6px}
.tile{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 16px}
.tile b{display:block;font-size:30px;line-height:1.2}.tile span{color:var(--muted);font-size:13px}
.t-Kritisch b{color:var(--crit)}.t-Warnung b{color:var(--warn)}.t-Info b{color:var(--info)}.t-OK b{color:var(--ok)}
.wrap{overflow-x:auto;border:1px solid var(--line);border-radius:10px;background:var(--card)}
table{width:100%;border-collapse:collapse}
th,td{text-align:left;padding:8px 11px;border-bottom:1px solid var(--line);vertical-align:top}
tr:last-child>td{border-bottom:0}
th{font-size:11.5px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);font-weight:600}
.badge{display:inline-block;padding:1px 9px;border-radius:999px;font-size:12px;font-weight:600;white-space:nowrap}
.b-Kritisch{background:var(--critbg);color:var(--crit)}.b-Warnung{background:var(--warnbg);color:var(--warn)}.b-Info{background:var(--infobg);color:var(--info)}.b-OK{background:var(--okbg);color:var(--ok)}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin:10px 0}
.card.s-Kritisch{border-left:4px solid var(--crit)}.card.s-Warnung{border-left:4px solid var(--warn)}.card.s-Info{border-left:4px solid var(--info)}.card.s-OK{border-left:4px solid var(--ok)}
.card h3{font-size:16px;margin:4px 0 2px}.card .erg{margin:0 0 8px}
.sec{display:grid;grid-template-columns:150px 1fr;gap:4px 14px;font-size:14px;margin-top:8px}
.sec dt{color:var(--muted);font-weight:600}.sec dd{margin:0}
.emp{display:inline-block;padding:1px 9px;border-radius:999px;font-size:12px;font-weight:600;background:var(--card2);border:1px solid var(--line)}
.hint{color:var(--muted);font-size:13px;margin-top:4px}
details{margin-top:8px}summary{cursor:pointer;color:var(--accent);font-size:13px}
table.det{font-size:12.5px;margin-top:6px}
table.det td,table.det th{padding:4px 8px}
.mehr,.none{color:var(--muted);font-size:13px;font-style:italic}
pre{white-space:pre-wrap;font:12.5px Consolas,monospace;margin:6px 0 0}
ul.sec2{margin:6px 0 0;padding-left:20px}ul.sec2 li{margin:3px 0}
.ok-box{background:var(--okbg);color:var(--ok);border-radius:10px;padding:12px 16px;font-weight:600}
.okrow{display:flex;gap:10px;align-items:baseline;padding:6px 2px;border-bottom:1px solid var(--line);font-size:14px}
.okrow b{min-width:260px}
@media (max-width:640px){.sec{grid-template-columns:1fr}.okrow{flex-direction:column}.okrow b{min-width:0}}
'@

function ConvertTo-HtmlText {
    param($Wert)
    if ($null -eq $Wert) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Wert)
}

function ConvertTo-HtmlTableSnippet {
    param([object[]] $Zeilen, [int] $Max = 250)
    if (-not $Zeilen -or $Zeilen.Count -eq 0) { return '' }
    $spalten = @($Zeilen[0].PSObject.Properties | ForEach-Object { $_.Name })
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="wrap"><table class="det"><thead><tr>')
    foreach ($s in $spalten) { [void]$sb.Append('<th>' + (ConvertTo-HtmlText $s) + '</th>') }
    [void]$sb.Append('</tr></thead><tbody>')
    $n = 0
    foreach ($z in $Zeilen) {
        if ($n -ge $Max) { break }
        $n++
        [void]$sb.Append('<tr>')
        foreach ($s in $spalten) { [void]$sb.Append('<td>' + (ConvertTo-HtmlText $z.$s) + '</td>') }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table></div>')
    if ($Zeilen.Count -gt $Max) { [void]$sb.Append('<p class="mehr">... und ' + ($Zeilen.Count - $Max) + ' weitere</p>') }
    return $sb.ToString()
}

function Format-DetailsHtml {
    param($Details)
    if ($null -eq $Details) { return '' }
    if ($Details -is [string]) {
        if (-not $Details.Trim()) { return '' }
        return '<details><summary>Details</summary><pre>' + (ConvertTo-HtmlText $Details) + '</pre></details>'
    }
    $zeilen = @($Details | Where-Object { $null -ne $_ })
    if ($zeilen.Count -eq 0) { return '' }
    return '<details><summary>Details (' + $zeilen.Count + ')</summary>' + (ConvertTo-HtmlTableSnippet $zeilen) + '</details>'
}

# Befund-Karte: Was ist das? / Woran liegt es? / Was ist zu tun? / Empfehlung / Loesungen
function ConvertTo-FindingCard {
    param($Befund, [switch] $MitBereich)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="card s-' + $Befund.Status + '"><span class="badge b-' + $Befund.Status + '">' + $Befund.Status + '</span>')
    if ($MitBereich) { [void]$sb.Append(' <span class="meta">' + (ConvertTo-HtmlText $Script:BereichTitel[$Befund.Bereich]) + '</span>') }
    [void]$sb.Append('<h3>' + (ConvertTo-HtmlText $Befund.Pruefung) + '</h3><p class="erg">' + (ConvertTo-HtmlText $Befund.Ergebnis) + '</p><dl class="sec">')
    if ($Befund.Was) { [void]$sb.Append('<dt>Was ist das?</dt><dd>' + (ConvertTo-HtmlText $Befund.Was) + '</dd>') }
    if ($Befund.Ursache) { [void]$sb.Append('<dt>Woran liegt es?</dt><dd>' + (ConvertTo-HtmlText $Befund.Ursache) + '</dd>') }
    if ($Befund.Massnahme) { [void]$sb.Append('<dt>Was ist zu tun?</dt><dd>' + (ConvertTo-HtmlText $Befund.Massnahme) + '</dd>') }
    if ($Befund.Empfehlung) { [void]$sb.Append('<dt>Empfehlung</dt><dd><span class="emp">' + (ConvertTo-HtmlText $Befund.Empfehlung) + '</span></dd>') }
    $bereinigung = @($Befund.Loesungen | Where-Object { $_.Art -eq 'Bereinigung' })
    if ($bereinigung.Count) { [void]$sb.Append('<dt>In der App lösbar</dt><dd>' + (ConvertTo-HtmlText ((@($bereinigung | ForEach-Object { $_.Titel }) -join ' · '))) + '</dd>') }
    [void]$sb.Append('</dl>')
    if ($Befund.Hinweis) { [void]$sb.Append('<div class="hint">' + (ConvertTo-HtmlText $Befund.Hinweis) + '</div>') }
    [void]$sb.Append((Format-DetailsHtml $Befund.Details))
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function Get-StatusCounts {
    $z = [ordered]@{ Kritisch = 0; Warnung = 0; Info = 0; OK = 0 }
    foreach ($f in $Script:Befunde) { $z[$f.Status]++ }
    return [pscustomobject]$z
}

function Build-Report {
    param([object[]] $Aenderungen, [string] $AltZeit, [timespan] $Dauer, $Audit)
    $z = Get-StatusCounts
    $os = Get-Data 'OS'
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<!doctype html><html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">')
    [void]$sb.Append('<title>Systemdiagnose ' + (ConvertTo-HtmlText $Script:StartZeit.ToString('dd.MM.yyyy HH:mm')) + '</title><style>' + $Script:BerichtCss + '</style></head><body><main>')
    [void]$sb.Append('<h1>Systemdiagnose</h1>')
    $rechte = if ($Script:IstAdmin) { 'mit Adminrechten' } else { 'ohne Adminrechte' }
    $pc = if ($os) { $os.CSName } else { $env:COMPUTERNAME }
    [void]$sb.Append('<div class="meta">' + (ConvertTo-HtmlText ("$pc · $($Script:Profil.Name)")) + '<br>' + (ConvertTo-HtmlText ("{0} · Modus {1} ({2}) · Ereignisse der letzten {3} Tage · Dauer {4:N0} s · Version {5}" -f $Script:StartZeit.ToString('dd.MM.yyyy HH:mm'), $Modus, $rechte, $Tage, $Dauer.TotalSeconds, $Script:Version)) + '</div>')

    [void]$sb.Append('<div class="tiles">')
    foreach ($k in @('Kritisch', 'Warnung', 'Info', 'OK')) { [void]$sb.Append('<div class="tile t-' + $k + '"><b>' + $z.$k + '</b><span>' + $k + '</span></div>') }
    [void]$sb.Append('</div>')

    $offen = @($Script:Befunde | Where-Object { $_.Status -in @('Kritisch', 'Warnung') } | Sort-Object @{ E = { Get-StatusRang $_.Status } }, Nr)
    [void]$sb.Append('<h2>Handlungsbedarf</h2>')
    if ($offen.Count -eq 0) { [void]$sb.Append('<div class="ok-box">Nichts zu tun - keine Warnungen und nichts Kritisches.</div>') }
    else { foreach ($f in $offen) { [void]$sb.Append((ConvertTo-FindingCard $f -MitBereich)) } }

    [void]$sb.Append('<h2>Änderungen seit dem letzten Lauf</h2>')
    if (-not $AltZeit) { [void]$sb.Append('<p class="none">Erster Lauf - ab dem nächsten Lauf steht hier, was sich verändert hat (Treiber, Dienste, Autostart, Programme, Einstellungen, Monitore, Speicherplatz).</p>') }
    elseif (@($Aenderungen).Count -eq 0) { [void]$sb.Append('<p class="none">Keine Änderungen seit ' + (ConvertTo-HtmlText $AltZeit) + '.</p>') }
    else { [void]$sb.Append('<p class="meta">Verglichen mit dem Lauf vom ' + (ConvertTo-HtmlText $AltZeit) + '</p>' + (ConvertTo-HtmlTableSnippet @($Aenderungen | Sort-Object Bereich, Art, Eintrag) 400)) }

    foreach ($bereich in $Script:BereichReihenfolge) {
        $liste = @($Script:Befunde | Where-Object { $_.Bereich -eq $bereich } | Sort-Object @{ E = { Get-StatusRang $_.Status } }, Nr)
        if ($liste.Count -eq 0) { continue }
        [void]$sb.Append('<h2>' + (ConvertTo-HtmlText $Script:BereichTitel[$bereich]) + '</h2>')
        foreach ($f in $liste) {
            if ($f.Status -in @('Kritisch', 'Warnung')) {
                [void]$sb.Append('<p class="meta">' + (ConvertTo-HtmlText $f.Pruefung) + ': siehe „Handlungsbedarf“ oben.</p>')
            } elseif ($f.Status -eq 'Info' -and $f.Ursache) {
                [void]$sb.Append((ConvertTo-FindingCard $f))
            } else {
                [void]$sb.Append('<div class="okrow"><span class="badge b-' + $f.Status + '">' + $f.Status + '</span><b>' + (ConvertTo-HtmlText $f.Pruefung) + '</b><span>' + (ConvertTo-HtmlText $f.Ergebnis) + $(if ($f.Hinweis) { '<div class="hint">' + (ConvertTo-HtmlText $f.Hinweis) + '</div>' } else { '' }) + (Format-DetailsHtml $f.Details) + '</span></div>')
            }
        }
    }

    [void]$sb.Append('<h2>Datenquellen</h2><p class="meta">Jede Quelle wurde genau einmal gelesen. Alle Prüfungen werten diesen Zwischenspeicher aus.</p>')
    [void]$sb.Append((ConvertTo-HtmlTableSnippet $Script:Quellen.ToArray() 200))

    [void]$sb.Append('<h2>Sicherheit dieses Werkzeugs</h2><ul class="sec2">')
    [void]$sb.Append('<li>Selbstprüfung vor dem Start bestanden: ' + $Audit.Befehle + ' Befehle und ' + $Audit.Methoden + ' Methodenaufrufe geprüft, keine schreibenden Zugriffe.</li>')
    [void]$sb.Append('<li>WhatIf-Sicherheitsnetz aktiv: Jeder Windows-Befehl mit Änderungswirkung hätte nur simuliert.</li>')
    [void]$sb.Append('<li>Registry ausschließlich über schreibgeschützte Handles gelesen.</li>')
    [void]$sb.Append('<li>Externe Programme nur aus System32 mit freigegebenen, lesenden Parametern: ' + (ConvertTo-HtmlText ((@($Script:ProgrammFreigaben.Keys) | Sort-Object) -join ', ')) + '.</li>')
    [void]$sb.Append('<li>Lösungen werden hier nur beschrieben. Ausgeführt werden sie ausschließlich vom getrennten Bereinigungsteil in der App - nach deiner Bestätigung und erneuter Prüfung jeder einzelnen Aktion.</li>')
    [void]$sb.Append('<li>Einziger Schreibvorgang: dieser Bericht, die Befunddatei und die Vergleichsdatei im Ordner ' + (ConvertTo-HtmlText $Script:BerichtOrdner) + '.</li></ul>')
    [void]$sb.Append('</main></body></html>')
    return $sb.ToString()
}

# Befunddatei fuer die App (reine Daten)
function Build-ErgebnisJson {
    param([object[]] $Aenderungen, [string] $AltZeit, [timespan] $Dauer, [string] $BerichtPfad)
    $befunde = foreach ($f in $Script:Befunde) {
        $det = $f.Details
        if ($det -is [string]) { $det = @([pscustomobject]@{ Text = $det }) }
        [ordered]@{
            Nr = $f.Nr; Typ = $f.Typ; Bereich = $f.Bereich; BereichTitel = $Script:BereichTitel[$f.Bereich]
            BereichRang = [Array]::IndexOf($Script:BereichReihenfolge, $f.Bereich); Pruefung = $f.Pruefung; Status = $f.Status
            Rang = (Get-StatusRang $f.Status); Ergebnis = $f.Ergebnis; Was = $f.Was; Ursache = $f.Ursache; Massnahme = $f.Massnahme
            Empfehlung = $f.Empfehlung; Hinweis = $f.Hinweis; Details = @($det | Where-Object { $null -ne $_ }); Loesungen = @($f.Loesungen)
            Zuletzt = $f.Zuletzt
        }
    }
    $obj = [ordered]@{
        Format = 'Systemdiagnose-Befunde/2'; Version = $Script:Version; Zeit = $Script:StartZeit.ToString('yyyy-MM-dd HH:mm:ss')
        ZeitText = $Script:StartZeit.ToString('dd.MM.yyyy HH:mm'); Modus = $Modus; Admin = $Script:IstAdmin; Tage = $Tage
        DauerSekunden = [Math]::Round($Dauer.TotalSeconds); Profil = $Script:Profil.Name; Computer = $env:COMPUTERNAME
        Maschine = (Get-MaschinenDaten)
        Bericht = $BerichtPfad; Zaehler = (Get-StatusCounts); AltZeit = $AltZeit
        Befunde = @($befunde); Aenderungen = @($Aenderungen); Quellen = $Script:Quellen.ToArray()
    }
    return ($obj | ConvertTo-Json -Depth 8)
}

# Messdatei fuer den Reiter "Reinigung" der App (reine Daten)
function Build-ReinigungJson {
    $obj = [ordered]@{
        Format = 'Systemdiagnose-Reinigungsmessung/1'; Version = $Script:Version; Zeit = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        ZeitText = (Get-Date).ToString('dd.MM.yyyy HH:mm'); Admin = $Script:IstAdmin; Maschine = $Script:MaschinenId
        Fehler = $(if ($Script:QuellenFehler.ContainsKey('Reinigung')) { [string]$Script:QuellenFehler['Reinigung'] } else { '' })
        Kategorien = @(Get-Data 'Reinigung')
    }
    return ($obj | ConvertTo-Json -Depth 5)
}

function Write-Summary {
    param([string] $Pfad, [string] $JsonPfad)
    $z = Get-StatusCounts
    Write-Host ''
    Write-Host '  ----------------------------------------------------------------' -ForegroundColor Cyan
    Write-Host '  Ergebnis:  ' -NoNewline
    Write-Host ('{0} kritisch' -f $z.Kritisch) -ForegroundColor $(if ($z.Kritisch) { 'Red' } else { 'DarkGray' }) -NoNewline
    Write-Host '  |  ' -NoNewline
    Write-Host ('{0} Warnungen' -f $z.Warnung) -ForegroundColor $(if ($z.Warnung) { 'Yellow' } else { 'DarkGray' }) -NoNewline
    Write-Host '  |  ' -NoNewline
    Write-Host ('{0} Info' -f $z.Info) -ForegroundColor Cyan -NoNewline
    Write-Host '  |  ' -NoNewline
    Write-Host ('{0} OK' -f $z.OK) -ForegroundColor Green
    $offen = @($Script:Befunde | Where-Object { $_.Status -in @('Kritisch', 'Warnung') } | Sort-Object @{ E = { Get-StatusRang $_.Status } }, Nr)
    foreach ($f in $offen) {
        $farbe = if ($f.Status -eq 'Kritisch') { 'Red' } else { 'Yellow' }
        Write-Host ('   [{0}] {1}: {2} - {3}' -f $f.Status.ToUpper(), $Script:BereichTitel[$f.Bereich], $f.Pruefung, (Format-Kurz $f.Ergebnis 110)) -ForegroundColor $farbe
        if ($f.Massnahme) { Write-Host ('       Was tun: ' + (Format-Kurz $f.Massnahme 150)) -ForegroundColor DarkGray }
    }
    if ($offen.Count -eq 0) { Write-Host '   Nichts zu tun.' -ForegroundColor Green }
    Write-Host ''
    Write-Host "  Bericht: $Pfad"
    Write-Host "  Befunddatei: $JsonPfad"
    Write-Host '  ----------------------------------------------------------------' -ForegroundColor Cyan
}

function Show-AuditDetails {
    Write-Host ''
    Write-Host '  Freigegebene Befehle:' -ForegroundColor Cyan
    Write-Host ('    ' + (($Script:ErlaubteBefehle | Sort-Object) -join ', '))
    Write-Host '  Freigegebene Programme (nur aus System32, nur diese Parameter):' -ForegroundColor Cyan
    foreach ($k in ($Script:ProgrammFreigaben.Keys | Sort-Object)) { Write-Host ('    {0,-14} {1}' -f $k, ($Script:ProgrammFreigaben[$k] -join ' | ')) }
    Write-Host '  Gesperrte Methoden:' -ForegroundColor Cyan
    Write-Host ('    ' + ($Script:VerboteneMethoden -join ', ') + ' sowie alles, was mit ' + ($Script:VerbotenePraefixe -replace '[\^\(\)]', '' -replace '\|', ', ') + ' beginnt')
    Write-Host '  Einziger Schreibzugriff:' -ForegroundColor Cyan
    Write-Host "    neue Dateien in $Script:BerichtOrdner (werden nie überschrieben)"
}

# ============================================================================
#  9. HAUPTABLAUF
# ============================================================================

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host "   Systemdiagnose $Script:Version - rein lesend" -ForegroundColor Cyan
Write-Host '  ================================================================' -ForegroundColor Cyan

$audit = Invoke-SelfAudit
if (-not $audit.Ok) {
    Write-Host '  [ABBRUCH] Die Sicherheitsprüfung hat nicht freigegebene Befehle gefunden:' -ForegroundColor Red
    foreach ($v in $audit.Verstoesse) { Write-Host "    - $v" -ForegroundColor Red }
    Write-Host '  Es wurde nichts ausgeführt.' -ForegroundColor Red
    if ($Pause) { [void](Read-Host '  Enter zum Schließen') }
    return
}
Write-Host ('  [OK] Sicherheitsprüfung: {0} Befehle und {1} Methodenaufrufe geprüft - nur lesend.' -f $audit.Befehle, $audit.Methoden) -ForegroundColor Green

if ($Selbsttest) {
    Show-AuditDetails
    if ($Pause) { [void](Read-Host '  Enter zum Schließen') }
    return
}

$Script:IstAdmin = Test-IstAdmin
$Script:AdminModus = ($Modus -notin @('Schnell', 'Reinigung')) -and $Script:IstAdmin
Initialize-Profil
Write-Host ("  PC: {0}  |  Sollwerte: {1}" -f $Script:Profil.Name, $(switch ($Script:ProfilQuelle) { 'eingebaut' { 'eingebaut (PC 1)' } 'eingelernt' { "eingelernt $Script:ProfilStand" } default { 'noch keine (neuer PC)' } }))

# Nur messen, was die Reinigungskategorien freigeben wuerden (Reiter "Reinigung" der App)
if ($Modus -eq 'Reinigung') {
    Write-Host ("  Modus: Reinigung (nur messen)  |  Adminrechte: {0}" -f $(if ($Script:IstAdmin) { 'ja' } else { 'nein' }))
    Write-Host ''
    [void](Get-Data 'Reinigung')
    $reinPfad = Save-ReportFile -DateiName ('Reinigung-' + $Script:StartZeit.ToString('yyyy-MM-dd_HHmmss') + '.json') -Inhalt (Build-ReinigungJson)
    Write-Host ''
    Write-Host "  Reinigungsdatei: $reinPfad"
    if ($Pause) { [void](Read-Host '  Enter zum Schließen') }
    return
}

Write-Host ("  Modus: {0}  |  Adminrechte: {1}  |  Ereignisse: letzte {2} Tage" -f $Modus, $(if ($Script:IstAdmin) { 'ja' } else { 'nein' }), $Tage)
if ($Modus -ne 'Schnell' -and -not $Script:IstAdmin) {
    Write-Host '  Hinweis: ohne Adminrechte gestartet - Tiefenprüfungen werden übersprungen.' -ForegroundColor Yellow
}
Write-Host ''
Write-Host '  Datenquellen (jede wird genau einmal gelesen):' -ForegroundColor Cyan

Test-Stabilitaet
Test-HardwareTreiber
Test-Laufwerke
Test-SoftwareAutostart
Test-Einstellungen
Test-Sicherheit
Test-SystemBasis
Test-Netzwerk

$letzter = Get-LastSnapshot
$schnappschuss = New-Snapshot
$aenderungen = @(Compare-Snapshot $letzter $schnappschuss)
Write-Progress -Activity 'Systemdiagnose' -Completed

$dauer = (Get-Date) - $Script:StartZeit
$stempel = $Script:StartZeit.ToString('yyyy-MM-dd_HHmmss')
$altZeit = if ($letzter) { [string]$letzter.Zeit } else { '' }
$html = Build-Report -Aenderungen $aenderungen -AltZeit $altZeit -Dauer $dauer -Audit $audit
$berichtPfad = Save-ReportFile -DateiName "Diagnose-$stempel.html" -Inhalt $html
$jsonPfad = Save-ReportFile -DateiName "Befunde-$stempel.json" -Inhalt (Build-ErgebnisJson -Aenderungen $aenderungen -AltZeit $altZeit -Dauer $dauer -BerichtPfad $berichtPfad)
[void](Save-ReportFile -DateiName "Snapshot-$stempel.json" -Inhalt ($schnappschuss | ConvertTo-Json -Depth 5))
$reinPfad = Save-ReportFile -DateiName "Reinigung-$stempel.json" -Inhalt (Build-ReinigungJson)

Write-Summary $berichtPfad $jsonPfad
Write-Host "  Reinigungsdatei: $reinPfad"
if (-not $KeinBrowser) { Open-Report $berichtPfad }
if ($Pause) { [void](Read-Host '  Enter zum Schließen') }
