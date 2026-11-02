# Reinigungskategorien der Systemdiagnose - reine Daten (wird mit Import-PowerShellDataFile gelesen,
# darin kann kein Code ausgefuehrt werden). Die Diagnose misst damit, Bereinigung.ps1 loescht damit.
#
# Ziele: Basis   = fester Name, den beide Skripte selbst aufloesen (LocalAppData, AppData, Windows,
#                  ProgramData, Systemlaufwerk) - hier steht nie ein absoluter Pfad.
#        Pfad    = Unterordner der Basis; '*' als ganzer Ordnername = alle Unterordner (z. B. Browser-Profile).
#        Muster  = welche Dateien (z. B. *.log); Rekursiv = auch in Unterordnern.
#        MinAlterStunden = nur Dateien, die mindestens so lange nicht mehr geaendert wurden.
# Prozesse  = Programme, die geschlossen sein muessen (sonst wird die Kategorie uebersprungen).
# Art       = Dateien (Standard), Papierkorb oder Lieferoptimierung (eigene Windows-Funktion).
# Bedingung = KeinNeustartAusstehend: nicht reinigen, solange Updates auf einen Neustart warten.
# Bereinigung.ps1 prueft zusaetzlich hart: nie Laufwerkswurzel, Windows-/Programmordner, Benutzer-
# ordner (Dokumente, Desktop, Downloads ...), nie Verknuepfungen (Junctions/Symlinks) folgen.
@{
    Format     = 'Systemdiagnose-Reinigung/1'
    Kategorien = @(
        # ------------------------------------------------------------------ Windows
        @{
            Id = 'BenutzerTemp'; Gruppe = 'Windows'; Name = 'Temporäre Dateien (Benutzer)'; Vorauswahl = $true
            Beschreibung = 'Reste von Installationen, Updates und Programmen, älter als 24 Stunden.'
            Ziele = @( @{ Basis = 'LocalAppData'; Pfad = 'Temp'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 24 } )
        }
        @{
            Id = 'WindowsTemp'; Gruppe = 'Windows'; Name = 'Temporäre Dateien (Windows)'; Vorauswahl = $true
            Beschreibung = 'Temp-Ordner von Windows selbst, älter als 24 Stunden.'
            Ziele = @( @{ Basis = 'Windows'; Pfad = 'Temp'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 24 } )
        }
        @{
            Id = 'UpdateDownloads'; Gruppe = 'Windows'; Name = 'Windows-Update-Downloads'; Vorauswahl = $true
            Beschreibung = 'Bereits installierte Update-Pakete, älter als 3 Tage.'
            Hinweis = 'Wird nicht angefasst, solange Updates auf einen Neustart warten.'
            Bedingung = 'KeinNeustartAusstehend'
            Ziele = @( @{ Basis = 'Windows'; Pfad = 'SoftwareDistribution\Download'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 72 } )
        }
        @{
            Id = 'Lieferoptimierung'; Gruppe = 'Windows'; Name = 'Übermittlungsoptimierung'; Vorauswahl = $true; Art = 'Lieferoptimierung'
            Beschreibung = 'Zwischenspeicher, über den Windows Update-Dateien mit anderen PCs im Netzwerk teilt.'
            Ziele = @()
        }
        @{
            Id = 'Upgradeprotokolle'; Gruppe = 'Windows'; Name = 'Windows-Upgrade-Protokolle'; Vorauswahl = $true
            Beschreibung = 'Protokolle früherer Windows-Upgrades (setupact.log, MigLog.xml), älter als 30 Tage.'
            Ziele = @(
                @{ Basis = 'Windows'; Pfad = 'Panther'; Muster = '*.log'; Rekursiv = $false; MinAlterStunden = 720 }
                @{ Basis = 'Windows'; Pfad = 'Panther'; Muster = 'MigLog*.xml'; Rekursiv = $false; MinAlterStunden = 720 }
                @{ Basis = 'Windows'; Pfad = 'Panther\monitor'; Muster = '*.log'; Rekursiv = $false; MinAlterStunden = 720 }
            )
        }
        @{
            Id = 'Protokolle'; Gruppe = 'Windows'; Name = 'Windows-Protokolldateien'; Vorauswahl = $true
            Beschreibung = 'Alte Protokolle der Komponentenwartung (CBS) und von DISM, älter als 14 Tage.'
            Ziele = @(
                @{ Basis = 'Windows'; Pfad = 'Logs\CBS'; Muster = '*.log'; Rekursiv = $false; MinAlterStunden = 336 }
                @{ Basis = 'Windows'; Pfad = 'Logs\CBS'; Muster = '*.cab'; Rekursiv = $false; MinAlterStunden = 336 }
                @{ Basis = 'Windows'; Pfad = 'Logs\DISM'; Muster = '*.log'; Rekursiv = $false; MinAlterStunden = 336 }
            )
        }
        @{
            Id = 'Miniaturansichten'; Gruppe = 'Windows'; Name = 'Miniaturansichten-Cache'; Vorauswahl = $false
            Beschreibung = 'Vorschaubilder des Explorers. Werden beim nächsten Öffnen der Ordner neu erstellt.'
            Hinweis = 'Der Explorer hält die Dateien meist offen - die werden übersprungen.'
            Ziele = @( @{ Basis = 'LocalAppData'; Pfad = 'Microsoft\Windows\Explorer'; Muster = 'thumbcache_*.db'; Rekursiv = $false; MinAlterStunden = 0 } )
        }
        @{
            Id = 'Fehlerberichte'; Gruppe = 'Windows'; Name = 'Windows-Fehlerberichte'; Vorauswahl = $false
            Beschreibung = 'Gesendete und wartende Fehlerberichte an Microsoft.'
            Hinweis = 'Danach fehlen im Zuverlässigkeitsverlauf die technischen Details zu alten Abstürzen.'
            Ziele = @(
                @{ Basis = 'ProgramData'; Pfad = 'Microsoft\Windows\WER\ReportQueue'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'ProgramData'; Pfad = 'Microsoft\Windows\WER\ReportArchive'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Microsoft\Windows\WER\ReportQueue'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Microsoft\Windows\WER\ReportArchive'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
            )
        }
        @{
            Id = 'Absturzabbilder'; Gruppe = 'Windows'; Name = 'Absturzabbilder (älter als 14 Tage)'; Vorauswahl = $false
            Beschreibung = 'Speicherauszüge alter Bluescreens, Grafik-Hänger und Programmabstürze.'
            Hinweis = 'Neuere Abbilder bleiben für die Fehlersuche erhalten.'
            Ziele = @(
                @{ Basis = 'Windows'; Pfad = 'Minidump'; Muster = '*.dmp'; Rekursiv = $false; MinAlterStunden = 336 }
                @{ Basis = 'Windows'; Pfad = ''; Muster = 'MEMORY.DMP'; Rekursiv = $false; MinAlterStunden = 336 }
                @{ Basis = 'Windows'; Pfad = 'LiveKernelReports'; Muster = '*.dmp'; Rekursiv = $true; MinAlterStunden = 336 }
                @{ Basis = 'LocalAppData'; Pfad = 'CrashDumps'; Muster = '*.dmp'; Rekursiv = $false; MinAlterStunden = 336 }
            )
        }
        @{
            Id = 'Papierkorb'; Gruppe = 'Windows'; Name = 'Papierkorb'; Vorauswahl = $false; Art = 'Papierkorb'
            Beschreibung = 'Gelöschte Dateien auf allen Laufwerken - auch verwaiste Reste, die Windows selbst nicht mehr anzeigt.'
            Hinweis = 'Endgültig - danach lassen sich die Dateien nicht mehr wiederherstellen.'
            Ziele = @()
        }

        # ------------------------------------------------------------------ Browser (nur Cache!)
        @{
            Id = 'Chrome'; Gruppe = 'Browser'; Name = 'Google Chrome - Cache'; Vorauswahl = $true; Prozesse = @('chrome')
            Beschreibung = 'Zwischengespeicherte Webseiten und Shader aller Profile. Cookies, Passwörter, Verlauf und Tabs bleiben.'
            Ziele = @(
                @{ Basis = 'LocalAppData'; Pfad = 'Google\Chrome\User Data\*\Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Google\Chrome\User Data\*\Code Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Google\Chrome\User Data\*\GPUCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Google\Chrome\User Data\*\DawnWebGPUCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Google\Chrome\User Data\GrShaderCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Google\Chrome\User Data\ShaderCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Google\Chrome\User Data\GraphiteDawnCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
            )
        }
        @{
            Id = 'Edge'; Gruppe = 'Browser'; Name = 'Microsoft Edge - Cache'; Vorauswahl = $true; Prozesse = @('msedge')
            Beschreibung = 'Zwischengespeicherte Webseiten aller Profile. Cookies, Passwörter, Verlauf und Tabs bleiben.'
            Hinweis = 'Edge läuft oft unsichtbar im Hintergrund (Startbeschleunigung) - dann wird übersprungen.'
            Ziele = @(
                @{ Basis = 'LocalAppData'; Pfad = 'Microsoft\Edge\User Data\*\Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Microsoft\Edge\User Data\*\Code Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Microsoft\Edge\User Data\*\GPUCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Microsoft\Edge\User Data\GrShaderCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Microsoft\Edge\User Data\ShaderCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
            )
        }
        @{
            Id = 'Firefox'; Gruppe = 'Browser'; Name = 'Mozilla Firefox - Cache'; Vorauswahl = $true; Prozesse = @('firefox')
            Beschreibung = 'Zwischengespeicherte Webseiten aller Profile. Cookies, Passwörter, Verlauf und Tabs bleiben.'
            Ziele = @(
                @{ Basis = 'LocalAppData'; Pfad = 'Mozilla\Firefox\Profiles\*\cache2'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Mozilla\Firefox\Profiles\*\startupCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Mozilla\Firefox\Profiles\*\thumbnails'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
            )
        }
        @{
            Id = 'Brave'; Gruppe = 'Browser'; Name = 'Brave - Cache'; Vorauswahl = $true; Prozesse = @('brave')
            Beschreibung = 'Zwischengespeicherte Webseiten aller Profile. Cookies, Passwörter, Verlauf und Tabs bleiben.'
            Ziele = @(
                @{ Basis = 'LocalAppData'; Pfad = 'BraveSoftware\Brave-Browser\User Data\*\Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'BraveSoftware\Brave-Browser\User Data\*\Code Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'BraveSoftware\Brave-Browser\User Data\*\GPUCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
            )
        }

        # ------------------------------------------------------------------ Programme
        @{
            Id = 'Discord'; Gruppe = 'Programme'; Name = 'Discord - Cache und Protokolle'; Vorauswahl = $true; Prozesse = @('Discord')
            Beschreibung = 'Zwischengespeicherte Bilder, Videos und Programmprotokolle. Anmeldung und Einstellungen bleiben.'
            Ziele = @(
                @{ Basis = 'AppData'; Pfad = 'discord\Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'AppData'; Pfad = 'discord\Code Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'AppData'; Pfad = 'discord\GPUCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'AppData'; Pfad = 'discord\logs'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 24 }
            )
        }
        @{
            Id = 'Blitz'; Gruppe = 'Programme'; Name = 'Blitz - Cache'; Vorauswahl = $true; Prozesse = @('Blitz')
            Beschreibung = 'Zwischengespeicherte Daten der LoL-Hilfe Blitz. Anmeldung und Einstellungen bleiben.'
            Ziele = @(
                @{ Basis = 'AppData'; Pfad = 'Blitz\Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'AppData'; Pfad = 'Blitz\Code Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'AppData'; Pfad = 'Blitz\GPUCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
            )
        }
        @{
            Id = 'Steam'; Gruppe = 'Programme'; Name = 'Steam - Browser-Cache'; Vorauswahl = $true; Prozesse = @('steam', 'steamwebhelper')
            Beschreibung = 'Cache des eingebauten Steam-Browsers (Shop, Community). Spiele und Anmeldung bleiben.'
            Ziele = @(
                @{ Basis = 'LocalAppData'; Pfad = 'Steam\htmlcache\*\Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Steam\htmlcache\*\Code Cache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Steam\htmlcache\*\GPUCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'Steam\htmlcache\GrShaderCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
            )
        }

        # ------------------------------------------------------------------ Grafik & Treiber
        @{
            Id = 'ShaderCache'; Gruppe = 'Grafik & Treiber'; Name = 'DirectX- und Grafik-Shader-Cache'; Vorauswahl = $false
            Beschreibung = 'Vorübersetzte Shader von Spielen (DirectX, AMD, Vulkan).'
            Hinweis = 'Spiele übersetzen die Shader beim nächsten Start neu - kurze Ruckler. Sinnvoll nach einem Treiberwechsel oder bei Grafikfehlern.'
            Ziele = @(
                @{ Basis = 'LocalAppData'; Pfad = 'D3DSCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'AMD\DxCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'AMD\DxcCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'AMD\VkCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'AMD\GLCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'NVIDIA\DXCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
                @{ Basis = 'LocalAppData'; Pfad = 'NVIDIA\GLCache'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 0 }
            )
        }
        @{
            Id = 'AmdInstaller'; Gruppe = 'Grafik & Treiber'; Name = 'AMD-Installationsdateien (C:\AMD)'; Vorauswahl = $false
            Beschreibung = 'Entpackte Treiber-Installer von AMD und ASUS, älter als 7 Tage. Die installierten Treiber bleiben.'
            Hinweis = 'Enthält auch selbst angelegte Sicherungen (z. B. Chipset_Software_Backup) - nur löschen, wenn die Treiber laufen.'
            Ziele = @( @{ Basis = 'Systemlaufwerk'; Pfad = 'AMD'; Muster = '*'; Rekursiv = $true; MinAlterStunden = 168 } )
        }
    )
}
