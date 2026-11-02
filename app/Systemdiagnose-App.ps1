#Requires -Version 5.1
<#
.SYNOPSIS
    Programmfenster der Systemdiagnose.

.DESCRIPTION
    - Startet die (rein lesende) Diagnose-Engine Systemdiagnose.ps1 und zeigt die Befunde als Karten:
      Was ist das? / Woran liegt es? / Was ist zu tun? / Empfehlung / Loesung.
    - Loesungen der Art "Bereinigung" werden ausgewaehlt und gesammelt an den getrennten
      Bereinigungsteil Bereinigung.ps1 uebergeben (mit Adminrechten, nach Bestaetigung).
      Das Fenster selbst aendert nichts am System - ausser harmlosen Sofort-Aktionen
      (Einstellungsseite oeffnen, Mauszeiger-Sperre aufheben).
    - Verlauf der Bereinigungen mit "Rueckgaengig".
    - Reiter "Reinigung" (wie CCleaner): Kategorien aus Reinigung.psd1 mit Groesse ankreuzen und reinigen.
      Gemessen wird von der Engine (-Modus Reinigung), geloescht nur von Bereinigung.ps1 - die App
      uebergibt dabei nur Kategorie-Kennungen, nie Pfade.

.PARAMETER Vorschau
    Nur fuer Tests: rendert das Fenster (und den Bestaetigungsdialog) in PNG-Dateien statt es anzuzeigen.
#>
param([string] $Vorschau)

$ErrorActionPreference = 'Stop'

# Nur ein Fenster gleichzeitig: zwei Instanzen koennten parallel pruefen, bereinigen oder Archiv/Einstellungen schreiben.
# Laeuft schon eins, wird es nach vorn geholt und dieses beendet sich.
if (-not $Vorschau) {
    $Script:Einzelinstanz = New-Object System.Threading.Mutex($false, 'Local\Systemdiagnose-App')
    $frei = $false
    try { $frei = $Script:Einzelinstanz.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $frei = $true }
    if (-not $frei) {
        try { [void](New-Object -ComObject WScript.Shell).AppActivate('Systemdiagnose') } catch { }
        exit
    }
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Data, System.Drawing

$Script:Ordner             = Split-Path -Parent $PSCommandPath
# Nur lesen (Vorschau, Testkopien): nichts speichern, keine Pruefung/Behebung/Reinigung starten.
# Testkopien setzen den Schalter per Textersatz dieser Zeile auf $true, statt einzelne Funktionen stillzulegen.
$Script:NurLesen           = [bool]$Vorschau
# Datenordner wie in Systemdiagnose.ps1 (Store-Fassung: LocalState des Pakets statt schreibgeschuetztem Programmordner)
$Script:DatenOrdner        = $Script:Ordner
if ($Script:Ordner -match '\\WindowsApps\\([^\\_]+)_[^\\_]*_[^\\_]*_[^\\_]*_([^\\_]+)$') { $Script:DatenOrdner = Join-Path $env:LOCALAPPDATA ('Packages\' + $Matches[1] + '_' + $Matches[2] + '\LocalState') }
$Script:BerichtOrdner      = Join-Path $Script:DatenOrdner 'Berichte'
$Script:BereinigungsOrdner = Join-Path $Script:DatenOrdner 'Bereinigung'
$Script:Engine             = Join-Path $Script:Ordner 'Systemdiagnose.ps1'
$Script:Bereiniger         = Join-Path $Script:Ordner 'Bereinigung.ps1'
$Script:Lauf               = $null
$Script:Daten              = $null
$Script:ErgebnisPfad       = $null
$Script:Liste              = $null
$Script:Ansicht            = $null
$Script:Filter             = 'Offen'
$Script:SuchText           = ''
$Script:LetzterModus       = 'Schnell'
$Script:ReinigungStand     = ''
$Script:EinstellungenDatei = Join-Path $Script:DatenOrdner 'Einstellungen.json'
$Script:Einstellungen      = [pscustomobject]@{ Design = 'Standard'; AutoPruefen = $true; LetzterFehlerbericht = '' }
$Script:ArchivDatei        = Join-Path $Script:DatenOrdner 'Archiv.json'
# Updates (02.11.): GitHub-Repo mit app\ + manifest.json, signiert von Installer\Update-veroeffentlichen.ps1 auf PC 1.
# Leer = keine Update-Pruefung. Nur Manifeste mit gueltiger Signatur dieses Schluessels werden angenommen.
$Script:UpdateQuelle       = 'https://raw.githubusercontent.com/xdockner-debug/systemdiagnose-updates/main'
$Script:UpdateSchluessel   = '<RSAKeyValue><Modulus>3OXzifzON/ig9gB+xgm9JmCjPPjeJW6TSECMQAUpsKEKsqkCXnOqYN+it0WYoq6kqJhJa7RwTND/r+3dZ1OT4+8AmtTGB7jjNPTDJkF8s+dUkKgYdYovQpTK4tuCA/HZybMRTXF6ogKg/c5fwDSs+mgEB5Qz8D+3GMkK0/Ut++YTjWOLZd2eUtoMAYOENE/pgnUJymoj+UkzLY2f37f6ABy4QhFHPiEb/+hC6w4XdcKOjWCaOsWZV+EEnv6QDME0tgaT+mNPrt8M3Ks47MpUAW6ujAFRcJfjBGEQvisWskAG4gxtHvt+GPOAwgYbJVLi6XIormyT+pXpyWUVS4SNkFGPBe1V5OEomoFiXC0U2ijzvYHaG8jOm4TRvQkktg7b7K5JyFicvG7PadbXmctANPufCQiwF+w3ow9cim8klv5Fxzvyi2jhrJXgXJDurHJgkhWSYi1E4A153OxVgx9qT5m+x/a1tZwPsVW61EAXCpqqpQZ3uBZtzf70E7hySPNF</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>'
$Script:Update             = $null
$Script:Archiv             = @{}
# Seit dem letzten Aktualisieren ausgefuehrte Massnahmen: 'B:<Typ>|<Pruefung>' bzw. 'K:<Kategorie>' -> Status und Meldung
$Script:Ausgefuehrt        = @{}

# ============================================================================
#  Datenmodelle (fuer saubere Bindungen in WPF) und Fensterhelfer
# ============================================================================
if (-not ('SdBefund' -as [type])) {
    Add-Type -ReferencedAssemblies 'System.Data', 'System.Xml' -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Runtime.InteropServices;

public class SdLoesung : INotifyPropertyChanged {
    public string Id { get; set; }
    public string Titel { get; set; }
    public string Art { get; set; }
    public bool Umkehrbar { get; set; }
    public string Dauer { get; set; }
    public string Beschreibung { get; set; }
    public string ParameterJson { get; set; }
    public int BefundNr { get; set; }
    public string BefundTitel { get; set; }
    public string Empfehlung { get; set; }
    public bool IstBereinigung { get { return Art == "Bereinigung"; } }
    public bool IstSofort { get { return Art == "Sofort"; } }
    public string Kennzeichen { get { return Umkehrbar ? "umkehrbar" : "nicht umkehrbar"; } }
    public bool HatBeschreibung { get { return !string.IsNullOrEmpty(Beschreibung); } }
    public bool HatDauer { get { return !string.IsNullOrEmpty(Dauer); } }
    private bool _ausgewaehlt;
    public bool Ausgewaehlt {
        get { return _ausgewaehlt; }
        set {
            if (_ausgewaehlt == value) return;
            _ausgewaehlt = value;
            var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs("Ausgewaehlt"));
            var a = AuswahlGeaendert; if (a != null) a(this, EventArgs.Empty);
        }
    }
    public static event EventHandler AuswahlGeaendert;
    public event PropertyChangedEventHandler PropertyChanged;
}

public class SdBefund : INotifyPropertyChanged {
    public int Nr { get; set; }
    public string Bereich { get; set; }
    public string BereichTitel { get; set; }
    public int BereichRang { get; set; }
    public string Pruefung { get; set; }
    public string Status { get; set; }
    public int Rang { get; set; }
    public string Ergebnis { get; set; }
    public string Was { get; set; }
    public string Ursache { get; set; }
    public string Massnahme { get; set; }
    public string Empfehlung { get; set; }
    public string Hinweis { get; set; }
    public string Schluessel { get; set; }
    public string Zuletzt { get; set; }
    // Archiv: ausgeblendet, bis ein neues Ereignis kommt oder der Befund schlimmer wird
    private bool _archiviert;
    public bool IstArchiviert {
        get { return _archiviert; }
        set {
            if (_archiviert == value) return;
            _archiviert = value;
            foreach (var n in new[] { "IstArchiviert", "IstOffen", "ArchivKopf", "ArchivErklaerung", "ArchivKnopfText", "ArchivKnopfGlyph", "ArchivHinweis", "HatArchivHinweis" }) Melde(n);
        }
    }
    public string ArchiviertAm { get; set; }
    // Automatisch archiviert: Warnung/Hinweis, fuer den die App nichts tun kann (keine Loesung, nicht kritisch)
    public bool AutoArchiviert { get; set; }
    public string ArchivHinweis { get; set; }
    public bool HatArchivHinweis { get { return !string.IsNullOrEmpty(ArchivHinweis); } }
    public bool KannArchivieren { get { return Status != "OK"; } }
    public string ArchivKopf { get { return _archiviert ? (AutoArchiviert ? "Automatisch archiviert" : "Archiviert") : HatBereinigung ? "Nicht mehr relevant?" : "In der App nicht behebbar"; } }
    public string ArchivErklaerung { get {
        const string zurueck = "kommt von selbst zurück, sobald ein neues Ereignis auftritt oder er schlimmer wird.";
        if (_archiviert && AutoArchiviert) return "Die App kann hier nichts beheben, deshalb liegt der Befund gleich im Archiv und zählt nicht zum Handlungsbedarf. Er " + zurueck + " Mit „Wiederherstellen“ bleibt er dauerhaft sichtbar.";
        if (_archiviert) return "Archiviert am " + ArchiviertAm + ". Der Befund zählt nicht zum Handlungsbedarf und " + zurueck;
        if (HatBereinigung) return "Willst du das nicht beheben, kannst du den Befund archivieren: Er verschwindet aus „Handlungsbedarf“ und den Zählern und " + zurueck;
        return "Die Schritte unter „Was ist zu tun?“ erledigst du selbst - die App kann das nicht automatisch. Ist es erledigt oder nicht mehr relevant, archiviere den Befund: Er verschwindet aus „Handlungsbedarf“ und den Zählern und " + zurueck;
    } }
    public string ArchivKnopfText { get { return _archiviert ? "Wiederherstellen" : "Archivieren"; } }
    public string ArchivKnopfGlyph { get { return _archiviert ? "" : ""; } }
    public event PropertyChangedEventHandler PropertyChanged;
    private void Melde(string n) { var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(n)); }
    public DataTable Details { get; set; }
    public List<SdLoesung> Loesungen { get; set; }
    public SdBefund() { Loesungen = new List<SdLoesung>(); }
    public bool HatWas { get { return !string.IsNullOrEmpty(Was); } }
    public bool HatUrsache { get { return !string.IsNullOrEmpty(Ursache); } }
    public bool HatMassnahme { get { return !string.IsNullOrEmpty(Massnahme); } }
    public bool HatHinweis { get { return !string.IsNullOrEmpty(Hinweis); } }
    public bool HatEmpfehlung { get { return !string.IsNullOrEmpty(Empfehlung); } }
    public bool HatDetails { get { return Details != null && Details.Rows.Count > 0; } }
    public bool HatLoesungen { get { return Loesungen.Count > 0; } }
    public bool HatBereinigung { get { foreach (var l in Loesungen) if (l.IstBereinigung) return true; return false; } }
    public bool IstOffen { get { return !_archiviert && (Status == "Kritisch" || Status == "Warnung"); } }
    public string DetailsTitel { get { return Details == null ? "" : "Details anzeigen (" + Details.Rows.Count + ")"; } }
    public string Suchtext { get { return (Pruefung + " " + Ergebnis + " " + BereichTitel + " " + Ursache + " " + Massnahme).ToLowerInvariant(); } }
    public string StatusGlyph { get { switch (Status) { case "Kritisch": return ""; case "Warnung": return ""; case "Info": return ""; default: return ""; } } }
}

public class SdAktion {
    public string Titel { get; set; }
    public string Meldung { get; set; }
    public string Status { get; set; }
    public bool Umkehrbar { get; set; }
    public string Gruppe { get; set; }
    public string Glyph { get { switch (Status) { case "behoben": return ""; case "offen": return ""; case "neu": return ""; case "erledigt": return ""; case "fehlgeschlagen": return ""; case "abgelehnt": return ""; case "simuliert": return ""; default: return ""; } } }
}

public class SdProgramm {
    public string Programm { get; set; }
    public string Anleitung { get; set; }
    public List<string> Kategorien { get; set; }
    public SdProgramm() { Kategorien = new List<string>(); }
    public string FuerText { get { return "nötig für: " + string.Join(", ", Kategorien); } }
}

public class SdLauf {
    public string Datei { get; set; }
    public string Art { get; set; }
    public string Zeit { get; set; }
    public string Zusammenfassung { get; set; }
    public string Wiederherstellungspunkt { get; set; }
    public string Fehler { get; set; }
    public bool KannRueckgaengig { get; set; }
    public List<SdAktion> Aktionen { get; set; }
    public SdLauf() { Aktionen = new List<SdAktion>(); }
    public string Titel { get { return Art + " vom " + Zeit; } }
    // Symbol je Art (gleiche wie im Fortschritts-Overlay: Papierkorb, Schraubenschluessel, Rueckgaengig)
    public string Glyph { get { switch (Art) { case "Reinigung": return ""; case "Rückgängig": return ""; case "Simulation": return ""; default: return ""; } } }
    public bool HatFehler { get { return !string.IsNullOrEmpty(Fehler); } }
}

public static class SdFormat {
    public static string Groesse(long bytes) {
        if (bytes >= 1073741824L) return (bytes / 1073741824.0).ToString("N1") + " GB";
        if (bytes < 1048576L) return "< 1 MB";
        return Math.Round(bytes / 1048576.0).ToString("N0") + " MB";
    }
}

public class SdKategorie : INotifyPropertyChanged {
    public string Id { get; set; }
    public string Gruppe { get; set; }
    public string Name { get; set; }
    public string Beschreibung { get; set; }
    public string Hinweis { get; set; }
    public bool Vorauswahl { get; set; }
    public long Bytes { get; set; }
    public int Dateien { get; set; }
    public bool Messbar { get; set; }
    public string Laeuft { get; set; }
    public string Sperre { get; set; }
    public double Anteil { get; set; }
    // Mindestens: ohne Adminrechte war ein Teil der Ordner nicht lesbar - die Groesse ist ein Mindestwert
    public bool Mindestens { get; set; }
    // GruppeSumme: Groesse der ganzen Gruppe fuer die Ueberschrift (gesetzt in ConvertTo-Kategorien)
    public string GruppeSumme { get; set; }
    public bool IstLeer { get { return Messbar && Dateien == 0 && !Mindestens; } }
    public bool IstAuswaehlbar { get { return !IstLeer; } }
    public bool HatHinweis { get { return !string.IsNullOrEmpty(Hinweis); } }
    public bool HatWarnung { get { return !string.IsNullOrEmpty(Laeuft) || !string.IsNullOrEmpty(Sperre); } }
    public string WarnText { get { return !string.IsNullOrEmpty(Laeuft) ? "läuft gerade: " + Laeuft : Sperre; } }
    public string GroesseText { get { return !Messbar ? "?" : IstLeer ? "leer" : (Mindestens ? "mind. " : "") + SdFormat.Groesse(Bytes); } }
    public string DateienText { get { return !Messbar ? "nur als Admin messbar" : IstLeer ? "nichts zu löschen" : Dateien.ToString("N0") + " Dateien" + (Mindestens ? ", Rest nur als Admin lesbar" : ""); } }
    public double BalkenBreite { get { return Math.Max(0.0, Math.Min(1.0, Anteil)) * 150.0; } }
    private bool _ausgewaehlt;
    public bool Ausgewaehlt {
        get { return _ausgewaehlt; }
        set {
            if (_ausgewaehlt == value) return;
            _ausgewaehlt = value;
            var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs("Ausgewaehlt"));
            var a = AuswahlGeaendert; if (a != null) a(this, EventArgs.Empty);
        }
    }
    public static event EventHandler AuswahlGeaendert;
    public event PropertyChangedEventHandler PropertyChanged;
}

public static class SdFenster {
    [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    [DllImport("user32.dll")] static extern bool ClipCursor(IntPtr rect);
    public static void Titelleiste(IntPtr hwnd, bool dunkel, int hintergrund, int text) {
        int d = dunkel ? 1 : 0;
        DwmSetWindowAttribute(hwnd, 20, ref d, 4);
        DwmSetWindowAttribute(hwnd, 35, ref hintergrund, 4);
        DwmSetWindowAttribute(hwnd, 36, ref text, 4);
    }
    public static bool MauszeigerFreigeben() { return ClipCursor(IntPtr.Zero); }
}
'@
}

# ============================================================================
#  Design: Farben, Schriften, Steuerelement-Vorlagen
# ============================================================================
$Script:DesignXaml = @'
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <SolidColorBrush x:Key="Bg" Color="#16181D"/>
  <SolidColorBrush x:Key="Surface" Color="#1F2228"/>
  <SolidColorBrush x:Key="Surface2" Color="#262A31"/>
  <SolidColorBrush x:Key="SurfaceHover" Color="#2D323A"/>
  <SolidColorBrush x:Key="Linie" Color="#323741"/>
  <SolidColorBrush x:Key="Text" Color="#E8EAED"/>
  <SolidColorBrush x:Key="Muted" Color="#9AA3AE"/>
  <SolidColorBrush x:Key="Accent" Color="#4C8DFF"/>
  <SolidColorBrush x:Key="AccentHover" Color="#6AA1FF"/>
  <SolidColorBrush x:Key="AccentText" Color="#FFFFFF"/>
  <SolidColorBrush x:Key="AccentSoft" Color="#1B2940"/>
  <SolidColorBrush x:Key="Crit" Color="#FF6B6B"/>
  <SolidColorBrush x:Key="CritBg" Color="#3B1F22"/>
  <SolidColorBrush x:Key="Warn" Color="#F2B33D"/>
  <SolidColorBrush x:Key="WarnBg" Color="#3A2E16"/>
  <SolidColorBrush x:Key="Info" Color="#6AA1FF"/>
  <SolidColorBrush x:Key="InfoBg" Color="#1C2A42"/>
  <SolidColorBrush x:Key="Ok" Color="#4CC27C"/>
  <SolidColorBrush x:Key="OkBg" Color="#183022"/>
  <SolidColorBrush x:Key="Overlay" Color="#B3101216"/>
  <SolidColorBrush x:Key="Spalte" Color="#00000000"/>
  <FontFamily x:Key="Icons">Segoe Fluent Icons, Segoe MDL2 Assets</FontFamily>
  <FontFamily x:Key="Anzeige">Segoe UI Variable Display, Segoe UI</FontFamily>
  <BooleanToVisibilityConverter x:Key="BoolSicht"/>

  <Style TargetType="ScrollBar">
    <Setter Property="Width" Value="10"/>
    <Setter Property="MinWidth" Value="10"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ScrollBar">
          <Track x:Name="PART_Track" Orientation="Vertical" IsDirectionReversed="True">
            <Track.Thumb>
              <Thumb>
                <Thumb.Template>
                  <ControlTemplate TargetType="Thumb">
                    <Border CornerRadius="4" Background="{DynamicResource Linie}" Margin="2"/>
                  </ControlTemplate>
                </Thumb.Template>
              </Thumb>
            </Track.Thumb>
          </Track>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="Orientation" Value="Horizontal">
        <Setter Property="Width" Value="Auto"/>
        <Setter Property="MinWidth" Value="0"/>
        <Setter Property="Height" Value="10"/>
        <Setter Property="MinHeight" Value="10"/>
        <Setter Property="Template">
          <Setter.Value>
            <ControlTemplate TargetType="ScrollBar">
              <Track x:Name="PART_Track" Orientation="Horizontal">
                <Track.Thumb>
                  <Thumb>
                    <Thumb.Template>
                      <ControlTemplate TargetType="Thumb">
                        <Border CornerRadius="4" Background="{DynamicResource Linie}" Margin="2"/>
                      </ControlTemplate>
                    </Thumb.Template>
                  </Thumb>
                </Track.Thumb>
              </Track>
            </ControlTemplate>
          </Setter.Value>
        </Setter>
      </Trigger>
    </Style.Triggers>
  </Style>

  <Style x:Key="Knopf" TargetType="Button">
    <Setter Property="Foreground" Value="{DynamicResource Text}"/>
    <Setter Property="Background" Value="{DynamicResource Surface2}"/>
    <Setter Property="BorderBrush" Value="{DynamicResource Linie}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="Padding" Value="14,7"/>
    <Setter Property="FontSize" Value="13.5"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="Rahmen" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Rahmen" Property="Background" Value="{DynamicResource SurfaceHover}"/></Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="PrimaerKnopf" TargetType="Button">
    <Setter Property="Foreground" Value="{DynamicResource AccentText}"/>
    <Setter Property="Background" Value="{DynamicResource Accent}"/>
    <Setter Property="Padding" Value="18,8"/>
    <Setter Property="FontSize" Value="14"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="Rahmen" Background="{TemplateBinding Background}" CornerRadius="10" Padding="{TemplateBinding Padding}">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Rahmen" Property="Background" Value="{DynamicResource AccentHover}"/></Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="ModusKarte" TargetType="RadioButton">
    <Setter Property="Foreground" Value="{DynamicResource Text}"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border x:Name="Rahmen" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1.5" CornerRadius="12" Padding="16,11">
            <ContentPresenter/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Rahmen" Property="Background" Value="{DynamicResource SurfaceHover}"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="Rahmen" Property="BorderBrush" Value="{DynamicResource Accent}"/>
              <Setter TargetName="Rahmen" Property="Background" Value="{DynamicResource AccentSoft}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="Chip" TargetType="RadioButton">
    <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Grid>
            <Border x:Name="Rahmen" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="16"/>
            <!-- Klinge: die Akzentfuellung faehrt beim Anklicken von links aus wie ein Lichtschwert (einmalig, kein Dauereffekt).
                 Schein und heller Kern nur in den Star-Wars-Designs (ReiterKlingeSichtbar), sonst eine schlichte Fuellung -->
            <Grid x:Name="Klinge" RenderTransformOrigin="0,0.5" IsHitTestVisible="False" Opacity="0">
              <Grid.RenderTransform><ScaleTransform ScaleX="1"/></Grid.RenderTransform>
              <Border CornerRadius="16" Background="{DynamicResource Accent}">
                <Border.Effect><DropShadowEffect Color="{DynamicResource KlingeGlut}" ShadowDepth="0" BlurRadius="12" Opacity="{DynamicResource ReiterKlingeSichtbar}"/></Border.Effect>
              </Border>
              <Border CornerRadius="16" BorderThickness="1.2" BorderBrush="{DynamicResource KlingeKernHell}" Opacity="{DynamicResource ReiterKlingeSichtbar}"/>
            </Grid>
            <ContentPresenter Margin="13,5"/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Rahmen" Property="Background" Value="{DynamicResource SurfaceHover}"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <!-- Endzustand per Setter (sichtbar), die Animation faehrt nur aus: ohne Animationstakt bleibt der Filter trotzdem gefuellt -->
              <Setter TargetName="Rahmen" Property="BorderBrush" Value="{DynamicResource Accent}"/>
              <Setter TargetName="Klinge" Property="Opacity" Value="1"/>
              <Trigger.EnterActions>
                <BeginStoryboard>
                  <Storyboard>
                    <DoubleAnimation Storyboard.TargetName="Klinge" Storyboard.TargetProperty="(UIElement.RenderTransform).(ScaleTransform.ScaleX)" From="0" To="1" Duration="0:0:0.4" FillBehavior="Stop">
                      <DoubleAnimation.EasingFunction><CubicEase EasingMode="EaseOut"/></DoubleAnimation.EasingFunction>
                    </DoubleAnimation>
                  </Storyboard>
                </BeginStoryboard>
              </Trigger.EnterActions>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="IsChecked" Value="True"><Setter Property="Foreground" Value="{DynamicResource AccentText}"/></Trigger>
    </Style.Triggers>
  </Style>

  <!-- Knopf im Filter-Look (Pruefung starten): transparent, beim Darueberfahren faehrt dieselbe Klinge aus wie beim Filter -->
  <Style x:Key="ChipKnopf" TargetType="Button">
    <Setter Property="Foreground" Value="{DynamicResource Text}"/>
    <Setter Property="FontSize" Value="16"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Button">
          <Grid>
            <Border x:Name="Rahmen" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="22"/>
            <Grid x:Name="Klinge" RenderTransformOrigin="0,0.5" IsHitTestVisible="False" Opacity="0">
              <Grid.RenderTransform><ScaleTransform ScaleX="1"/></Grid.RenderTransform>
              <Border CornerRadius="22" Background="{DynamicResource Accent}">
                <Border.Effect><DropShadowEffect Color="{DynamicResource KlingeGlut}" ShadowDepth="0" BlurRadius="12" Opacity="{DynamicResource ReiterKlingeSichtbar}"/></Border.Effect>
              </Border>
              <Border CornerRadius="22" BorderThickness="1.2" BorderBrush="{DynamicResource KlingeKernHell}" Opacity="{DynamicResource ReiterKlingeSichtbar}"/>
            </Grid>
            <ContentPresenter Margin="40,12"/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Rahmen" Property="BorderBrush" Value="{DynamicResource Accent}"/>
              <Setter TargetName="Klinge" Property="Opacity" Value="1"/>
              <Setter Property="Foreground" Value="{DynamicResource AccentText}"/>
              <Trigger.EnterActions>
                <BeginStoryboard>
                  <Storyboard>
                    <DoubleAnimation Storyboard.TargetName="Klinge" Storyboard.TargetProperty="(UIElement.RenderTransform).(ScaleTransform.ScaleX)" From="0" To="1" Duration="0:0:0.4" FillBehavior="Stop">
                      <DoubleAnimation.EasingFunction><CubicEase EasingMode="EaseOut"/></DoubleAnimation.EasingFunction>
                    </DoubleAnimation>
                  </Storyboard>
                </BeginStoryboard>
              </Trigger.EnterActions>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- Stufen-Schalter: Schnell (aus) - Gruendlich (an); die aktive Seite ist hell, die andere grau -->
  <Style x:Key="Schalter" TargetType="CheckBox">
    <Setter Property="FontSize" Value="14.5"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="CheckBox">
          <StackPanel Orientation="Horizontal" Background="Transparent">
            <TextBlock x:Name="Links" Text="Schnell" Foreground="{DynamicResource Text}" VerticalAlignment="Center"/>
            <Grid Width="48" Height="26" Margin="11,0">
              <Border x:Name="Spur" CornerRadius="13" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1"/>
              <Ellipse x:Name="Punkt" Width="18" Height="18" Fill="{DynamicResource Muted}" HorizontalAlignment="Left" Margin="4,0"/>
            </Grid>
            <TextBlock x:Name="Rechts" Text="Gründlich" Foreground="{DynamicResource Muted}" VerticalAlignment="Center"/>
          </StackPanel>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Spur" Property="BorderBrush" Value="{DynamicResource Accent}"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="Spur" Property="Background" Value="{DynamicResource Accent}"/>
              <Setter TargetName="Spur" Property="BorderBrush" Value="{DynamicResource Accent}"/>
              <Setter TargetName="Punkt" Property="HorizontalAlignment" Value="Right"/>
              <Setter TargetName="Punkt" Property="Fill" Value="{DynamicResource AccentText}"/>
              <Setter TargetName="Links" Property="Foreground" Value="{DynamicResource Muted}"/>
              <Setter TargetName="Rechts" Property="Foreground" Value="{DynamicResource Text}"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- Kleines ⓘ zum Auf- und Zuklappen einer Erklaerung -->
  <Style x:Key="InfoSchalter" TargetType="ToggleButton">
    <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ToggleButton">
          <Border Background="Transparent" Padding="2">
            <TextBlock Text="&#xE946;" FontFamily="{StaticResource Icons}" FontSize="17" Foreground="{TemplateBinding Foreground}" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="{DynamicResource Text}"/></Trigger>
            <Trigger Property="IsChecked" Value="True"><Setter Property="Foreground" Value="{DynamicResource Accent}"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- Leiser Filter (Archiv): ohne Rahmen, Symbol + graue Schrift, ausgewaehlt in Akzentfarbe -->
  <Style x:Key="ChipLeise" TargetType="RadioButton">
    <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
    <Setter Property="FontSize" Value="12.5"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border x:Name="Rahmen" Background="Transparent" CornerRadius="16" Padding="9,6">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="&#xE7B8;" FontFamily="{DynamicResource Icons}" FontSize="12" Margin="0,2,6,0"/>
              <ContentPresenter/>
            </StackPanel>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Rahmen" Property="Background" Value="{DynamicResource SurfaceHover}"/></Trigger>
            <Trigger Property="IsChecked" Value="True"><Setter TargetName="Rahmen" Property="Background" Value="{DynamicResource AccentSoft}"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="{DynamicResource Text}"/></Trigger>
      <Trigger Property="IsChecked" Value="True"><Setter Property="Foreground" Value="{DynamicResource Accent}"/></Trigger>
    </Style.Triggers>
  </Style>

  <Style TargetType="CheckBox">
    <Setter Property="Foreground" Value="{DynamicResource Text}"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="CheckBox">
          <Grid Background="Transparent">
            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border x:Name="Box" Width="20" Height="20" CornerRadius="5" BorderThickness="1.5" BorderBrush="{DynamicResource Muted}"
                    Background="Transparent" VerticalAlignment="Top" Margin="0,1,12,0">
              <TextBlock x:Name="Haken" Text="&#xE73E;" FontFamily="{DynamicResource Icons}" FontSize="12" Foreground="{DynamicResource AccentText}"
                         HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
            </Border>
            <ContentPresenter Grid.Column="1" VerticalAlignment="Top"/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Box" Property="BorderBrush" Value="{DynamicResource Accent}"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="Box" Property="Background" Value="{DynamicResource Accent}"/>
              <Setter TargetName="Box" Property="BorderBrush" Value="{DynamicResource Accent}"/>
              <Setter TargetName="Haken" Property="Visibility" Value="Visible"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="TabControl">
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="TabControl">
          <Grid>
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
            <!-- Tag = Inhalt rechts in der Reiterzeile (dort sitzt „Pruefung starten“).
                 Keine Trenn- oder Schienenlinie: nur der Akzent-Strich unter dem offenen Reiter (schlichte Designs) bzw.
                 die Lichtschwert-Rahmen (Star Wars) -->
            <Border Padding="20,0">
              <DockPanel>
                <ContentPresenter DockPanel.Dock="Right" Content="{TemplateBinding Tag}" VerticalAlignment="Center"/>
                <Grid HorizontalAlignment="Left">
                  <TabPanel IsItemsHost="True"/>
                </Grid>
              </DockPanel>
            </Border>
            <ContentPresenter Grid.Row="1" ContentSource="SelectedContent"/>
          </Grid>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- Schrift und Farbe nur fuer die Reiterbeschriftung - der Seiteninhalt erbt sie so nicht -->
  <Style TargetType="TabItem">
    <Setter Property="Cursor" Value="Hand"/>
    <!-- Ohne eigene Farbe erbt der Reiterinhalt Schwarz aus dem Windows-Standarddesign -->
    <Setter Property="Foreground" Value="{DynamicResource Text}"/>
    <Setter Property="FontWeight" Value="Normal"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="TabItem">
          <!-- Abstaende kommen aus dem Design (Set-Design): im Star-Wars-Design Platz fuer den Lichtschwert-Rahmen -->
          <Grid Margin="{DynamicResource ReiterAbstand}" Background="Transparent">
            <!-- Lichtschwert-Rahmen: Kern fast weiss, roter (Clone Wars: blauer) Schein, Farben aus Set-Design; nur sichtbar, wenn ReiterKlingeSichtbar = 1 -->
            <Border x:Name="KlingeHuelle" Opacity="{DynamicResource ReiterKlingeSichtbar}" Margin="0,6,0,6" IsHitTestVisible="False">
             <Grid>
              <!-- Dezente Innenflaeche nur beim offenen Reiter: so bleibt er erkennbar, auch wenn die Maus einen anderen aufhellt -->
              <Border x:Name="Flaeche" CornerRadius="8" Margin="1.5" Background="{DynamicResource AccentSoft}" Opacity="0"/>
              <Grid x:Name="Klinge" Opacity="0.5">
                <!-- Schein: breiter roter Rand, weichgezeichnet -->
                <Border x:Name="Schein" CornerRadius="9" BorderThickness="3" BorderBrush="{DynamicResource KlingeSchein}" Opacity="0.55">
                  <Border.Effect><BlurEffect Radius="6"/></Border.Effect>
                </Border>
                <!-- Kern: duenne rote Linie; darueber der helle Kern (KernHell), der nur bei Maus darauf oder ausgewaehlt leuchtet.
                     Eigene Ebene mit Opacity statt Farbwechsel: Trigger setzen nur einfache Zahlen, keine Pinsel -->
                <Border x:Name="Kern" CornerRadius="8" BorderThickness="1.4" BorderBrush="{DynamicResource KlingeKern}" Margin="0.8">
                  <Border.Effect><DropShadowEffect Color="{DynamicResource KlingeGlut}" ShadowDepth="0" BlurRadius="6" Opacity="1"/></Border.Effect>
                </Border>
                <Border x:Name="KernHell" CornerRadius="8" BorderThickness="1.4" BorderBrush="{DynamicResource KlingeKernHell}" Margin="0.8" Opacity="0"/>
              </Grid>
             </Grid>
            </Border>
            <ContentPresenter x:Name="Kopf" ContentSource="Header" Margin="{DynamicResource ReiterKopfRand}"
                              TextElement.FontWeight="SemiBold" TextElement.FontSize="16" TextElement.Foreground="{DynamicResource Muted}"/>
            <Border x:Name="Strich" Height="3" CornerRadius="2" VerticalAlignment="Bottom" Background="Transparent" Opacity="{DynamicResource ReiterStrichSichtbar}"/>
          </Grid>
          <ControlTemplate.Triggers>
            <!-- Hell (Neon), solange die Maus darauf ist oder der Reiter offen ist; sonst gedimmt -->
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Kopf" Property="TextElement.Foreground" Value="{DynamicResource Text}"/>
              <Setter TargetName="Klinge" Property="Opacity" Value="1"/>
              <Setter TargetName="Schein" Property="Opacity" Value="1"/>
              <Setter TargetName="KernHell" Property="Opacity" Value="1"/>
            </Trigger>
            <Trigger Property="IsSelected" Value="True">
              <Setter TargetName="Strich" Property="Background" Value="{DynamicResource Accent}"/>
              <Setter TargetName="Flaeche" Property="Opacity" Value="1"/>
              <Setter TargetName="Kopf" Property="TextElement.Foreground" Value="{DynamicResource Text}"/>
              <Setter TargetName="Klinge" Property="Opacity" Value="1"/>
              <Setter TargetName="Schein" Property="Opacity" Value="1"/>
              <Setter TargetName="KernHell" Property="Opacity" Value="1"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="ListBox">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Foreground" Value="{DynamicResource Text}"/>
    <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
  </Style>

  <Style TargetType="ListBoxItem">
    <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
    <Setter Property="Margin" Value="0,1"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ListBoxItem">
          <Border x:Name="Rahmen" CornerRadius="9" Background="Transparent" Padding="6,0">
            <ContentPresenter/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Rahmen" Property="Background" Value="{DynamicResource SurfaceHover}"/></Trigger>
            <Trigger Property="IsSelected" Value="True"><Setter TargetName="Rahmen" Property="Background" Value="{DynamicResource Surface2}"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="DataGrid">
    <Setter Property="Background" Value="{DynamicResource Surface}"/>
    <Setter Property="Foreground" Value="{DynamicResource Text}"/>
    <Setter Property="BorderBrush" Value="{DynamicResource Linie}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="RowBackground" Value="Transparent"/>
    <Setter Property="AlternatingRowBackground" Value="{DynamicResource Surface2}"/>
    <Setter Property="GridLinesVisibility" Value="None"/>
    <Setter Property="HeadersVisibility" Value="Column"/>
    <Setter Property="IsReadOnly" Value="True"/>
    <Setter Property="AutoGenerateColumns" Value="True"/>
    <Setter Property="CanUserAddRows" Value="False"/>
    <Setter Property="SelectionMode" Value="Single"/>
    <Setter Property="FontSize" Value="12.5"/>
    <Setter Property="MinRowHeight" Value="26"/>
    <Setter Property="ColumnHeaderHeight" Value="30"/>
    <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
    <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
  </Style>
  <Style TargetType="DataGridColumnHeader">
    <Setter Property="Background" Value="{DynamicResource Surface2}"/>
    <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Padding" Value="8,0"/>
    <Setter Property="BorderBrush" Value="{DynamicResource Linie}"/>
    <Setter Property="BorderThickness" Value="0,0,0,1"/>
  </Style>
  <Style TargetType="DataGridRow">
    <Style.Triggers>
      <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="{DynamicResource AccentSoft}"/></Trigger>
    </Style.Triggers>
  </Style>
  <Style TargetType="DataGridCell">
    <Setter Property="Padding" Value="8,3"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="DataGridCell">
          <Border Padding="{TemplateBinding Padding}" Background="Transparent">
            <ContentPresenter VerticalAlignment="Center"/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="IsSelected" Value="True"><Setter Property="Foreground" Value="{DynamicResource Text}"/></Trigger>
    </Style.Triggers>
  </Style>

  <Style TargetType="Expander">
    <Setter Property="Foreground" Value="{DynamicResource Text}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Expander">
          <StackPanel>
            <ToggleButton IsChecked="{Binding IsExpanded, RelativeSource={RelativeSource TemplatedParent}, Mode=TwoWay}"
                          Content="{TemplateBinding Header}" Cursor="Hand" HorizontalAlignment="Left">
              <ToggleButton.Template>
                <ControlTemplate TargetType="ToggleButton">
                  <Border Background="Transparent" Padding="0,4">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock x:Name="Pfeil" Text="&#xE76C;" FontFamily="{DynamicResource Icons}" FontSize="11" Foreground="{DynamicResource Accent}" VerticalAlignment="Center" Margin="0,1,8,0"/>
                      <TextBlock Text="{Binding Content, RelativeSource={RelativeSource TemplatedParent}}" Foreground="{DynamicResource Accent}" FontWeight="SemiBold" FontSize="13"/>
                    </StackPanel>
                  </Border>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsChecked" Value="True"><Setter TargetName="Pfeil" Property="Text" Value="&#xE70D;"/></Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </ToggleButton.Template>
            </ToggleButton>
            <ContentPresenter x:Name="Inhalt" Visibility="Collapsed"/>
          </StackPanel>
          <ControlTemplate.Triggers>
            <Trigger Property="IsExpanded" Value="True"><Setter TargetName="Inhalt" Property="Visibility" Value="Visible"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="ProgressBar">
    <Setter Property="Height" Value="6"/>
    <Setter Property="Foreground" Value="{DynamicResource Accent}"/>
    <Setter Property="Background" Value="{DynamicResource Surface2}"/>
    <Setter Property="BorderThickness" Value="0"/>
  </Style>

  <!-- Text, der aussieht wie ein TextBlock, sich aber mit der Maus markieren und kopieren laesst (Strg+C, Rechtsklick).
       Decorator statt ScrollViewer als Inhalt: sonst schluckt jedes Textfeld das Mausrad der Karte -->
  <Style x:Key="Markierbar" TargetType="TextBox">
    <Setter Property="IsReadOnly" Value="True"/>
    <Setter Property="IsReadOnlyCaretVisible" Value="False"/>
    <Setter Property="IsUndoEnabled" Value="False"/>
    <Setter Property="TextWrapping" Value="Wrap"/>
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Padding" Value="-2,0,0,0"/>
    <Setter Property="Foreground" Value="{DynamicResource Text}"/>
    <Setter Property="SelectionBrush" Value="{DynamicResource Accent}"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="TextBox">
          <Decorator x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="Suche" TargetType="TextBox">
    <Setter Property="Background" Value="{DynamicResource Surface}"/>
    <Setter Property="Foreground" Value="{DynamicResource Text}"/>
    <Setter Property="CaretBrush" Value="{DynamicResource Text}"/>
    <Setter Property="BorderBrush" Value="{DynamicResource Linie}"/>
    <Setter Property="FontSize" Value="13.5"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="TextBox">
          <Border x:Name="Rahmen" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="9">
            <Grid>
              <TextBlock Text="&#xE721;" FontFamily="{DynamicResource Icons}" Foreground="{DynamicResource Muted}" FontSize="12" Margin="12,0,0,0" VerticalAlignment="Center"/>
              <ScrollViewer x:Name="PART_ContentHost" Margin="32,0,10,0" VerticalAlignment="Center"/>
              <TextBlock x:Name="Platzhalter" Text="In den Befunden suchen …" Foreground="{DynamicResource Muted}" Margin="34,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False" Visibility="Collapsed"/>
            </Grid>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="Text" Value=""><Setter TargetName="Platzhalter" Property="Visibility" Value="Visible"/></Trigger>
            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Rahmen" Property="BorderBrush" Value="{DynamicResource Accent}"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="StatusBalken" TargetType="Border">
    <Setter Property="Background" Value="{DynamicResource Ok}"/>
    <Style.Triggers>
      <DataTrigger Binding="{Binding Status}" Value="Kritisch"><Setter Property="Background" Value="{DynamicResource Crit}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="Warnung"><Setter Property="Background" Value="{DynamicResource Warn}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="Info"><Setter Property="Background" Value="{DynamicResource Info}"/></DataTrigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="StatusPille" TargetType="Border">
    <Setter Property="Background" Value="{DynamicResource OkBg}"/>
    <Setter Property="CornerRadius" Value="12"/>
    <Setter Property="Padding" Value="10,3"/>
    <Style.Triggers>
      <DataTrigger Binding="{Binding Status}" Value="Kritisch"><Setter Property="Background" Value="{DynamicResource CritBg}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="Warnung"><Setter Property="Background" Value="{DynamicResource WarnBg}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="Info"><Setter Property="Background" Value="{DynamicResource InfoBg}"/></DataTrigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="StatusText" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{DynamicResource Ok}"/>
    <Style.Triggers>
      <DataTrigger Binding="{Binding Status}" Value="Kritisch"><Setter Property="Foreground" Value="{DynamicResource Crit}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="Warnung"><Setter Property="Foreground" Value="{DynamicResource Warn}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="Info"><Setter Property="Foreground" Value="{DynamicResource Info}"/></DataTrigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="EmpfehlungPille" TargetType="Border">
    <Setter Property="Background" Value="{DynamicResource Surface2}"/>
    <Setter Property="CornerRadius" Value="12"/>
    <Setter Property="Padding" Value="11,3"/>
    <Setter Property="HorizontalAlignment" Value="Left"/>
    <Style.Triggers>
      <DataTrigger Binding="{Binding Empfehlung}" Value="Beheben"><Setter Property="Background" Value="{DynamicResource WarnBg}"/></DataTrigger>
      <DataTrigger Binding="{Binding Empfehlung}" Value="Prüfen"><Setter Property="Background" Value="{DynamicResource WarnBg}"/></DataTrigger>
      <DataTrigger Binding="{Binding Empfehlung}" Value="Entfernen"><Setter Property="Background" Value="{DynamicResource AccentSoft}"/></DataTrigger>
      <DataTrigger Binding="{Binding Empfehlung}" Value="Behalten"><Setter Property="Background" Value="{DynamicResource OkBg}"/></DataTrigger>
      <DataTrigger Binding="{Binding Empfehlung}" Value="Nichts zu tun"><Setter Property="Background" Value="{DynamicResource OkBg}"/></DataTrigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="EmpfehlungText" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="FontSize" Value="12.5"/>
    <Style.Triggers>
      <DataTrigger Binding="{Binding Empfehlung}" Value="Beheben"><Setter Property="Foreground" Value="{DynamicResource Warn}"/></DataTrigger>
      <DataTrigger Binding="{Binding Empfehlung}" Value="Prüfen"><Setter Property="Foreground" Value="{DynamicResource Warn}"/></DataTrigger>
      <DataTrigger Binding="{Binding Empfehlung}" Value="Entfernen"><Setter Property="Foreground" Value="{DynamicResource Accent}"/></DataTrigger>
      <DataTrigger Binding="{Binding Empfehlung}" Value="Behalten"><Setter Property="Foreground" Value="{DynamicResource Ok}"/></DataTrigger>
      <DataTrigger Binding="{Binding Empfehlung}" Value="Nichts zu tun"><Setter Property="Foreground" Value="{DynamicResource Ok}"/></DataTrigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="ErgebnisGlyph" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
    <Setter Property="FontFamily" Value="{DynamicResource Icons}"/>
    <Setter Property="FontSize" Value="14"/>
    <Style.Triggers>
      <DataTrigger Binding="{Binding Status}" Value="erledigt"><Setter Property="Foreground" Value="{DynamicResource Ok}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="fehlgeschlagen"><Setter Property="Foreground" Value="{DynamicResource Crit}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="abgelehnt"><Setter Property="Foreground" Value="{DynamicResource Warn}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="simuliert"><Setter Property="Foreground" Value="{DynamicResource Info}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="behoben"><Setter Property="Foreground" Value="{DynamicResource Ok}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="offen"><Setter Property="Foreground" Value="{DynamicResource Warn}"/></DataTrigger>
      <DataTrigger Binding="{Binding Status}" Value="neu"><Setter Property="Foreground" Value="{DynamicResource Crit}"/></DataTrigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="Abschnitt" TargetType="Border">
    <Setter Property="Background" Value="{DynamicResource Surface2}"/>
    <Setter Property="CornerRadius" Value="12"/>
    <Setter Property="Padding" Value="16,13"/>
    <Setter Property="Margin" Value="0,12,0,0"/>
  </Style>
  <Style x:Key="AbschnittKopf" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="FontSize" Value="12.5"/>
  </Style>
  <Style x:Key="Kennzeichen" TargetType="Border">
    <Setter Property="Background" Value="{DynamicResource OkBg}"/>
    <Setter Property="CornerRadius" Value="10"/>
    <Setter Property="Padding" Value="8,1"/>
    <Setter Property="Margin" Value="8,1,0,0"/>
    <Setter Property="VerticalAlignment" Value="Top"/>
    <Style.Triggers>
      <DataTrigger Binding="{Binding Umkehrbar}" Value="False"><Setter Property="Background" Value="{DynamicResource WarnBg}"/></DataTrigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="WarnChip" TargetType="Border">
    <Setter Property="Background" Value="{DynamicResource WarnBg}"/>
    <Setter Property="CornerRadius" Value="10"/>
    <Setter Property="Padding" Value="8,1"/>
    <Setter Property="Margin" Value="8,1,0,0"/>
    <Setter Property="VerticalAlignment" Value="Top"/>
  </Style>
  <Style x:Key="KennzeichenText" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{DynamicResource Ok}"/>
    <Setter Property="FontSize" Value="11.5"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Style.Triggers>
      <DataTrigger Binding="{Binding Umkehrbar}" Value="False"><Setter Property="Foreground" Value="{DynamicResource Warn}"/></DataTrigger>
    </Style.Triggers>
  </Style>
</ResourceDictionary>
'@

# ============================================================================
#  Hauptfenster
# ============================================================================
$Script:FensterXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Systemdiagnose" Width="1280" Height="860" MinWidth="1020" MinHeight="680"
        WindowStartupLocation="CenterScreen" Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="False" SnapsToDevicePixels="True">
  <Grid x:Name="Wurzel" Background="{DynamicResource Bg}">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <DockPanel Grid.Row="0" Margin="28,20,28,6">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
        <Button x:Name="UpdateKnopf" Style="{StaticResource PrimaerKnopf}" Margin="0,0,8,0" Visibility="Collapsed" ToolTip="Neue Version verfügbar - ohne Neuinstallation aktualisieren">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE896;" FontFamily="{StaticResource Icons}" Margin="0,2,8,0"/><TextBlock x:Name="UpdateText" Text="Update"/></StackPanel>
        </Button>
        <Button x:Name="OrdnerKnopf" Style="{StaticResource Knopf}" ToolTip="Ordner mit allen Berichten öffnen - je Prüfung ein HTML-Bericht (Diagnose-…html)">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE838;" FontFamily="{StaticResource Icons}" Margin="0,2,8,0"/><TextBlock Text="Berichte"/></StackPanel>
        </Button>
        <Button x:Name="FehlerKnopf" Style="{StaticResource Knopf}" Margin="8,0,0,0" ToolTip="Fehler melden - Fehlercodes dieses PCs für die Weiterentwicklung in eine Datei packen">
          <Grid>
            <TextBlock Text="&#xED15;" FontFamily="{StaticResource Icons}" Margin="0,2,0,0"/>
            <Ellipse x:Name="FehlerPunkt" Width="8" Height="8" Fill="{DynamicResource Crit}" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,-3,-6,0" Visibility="Collapsed"/>
          </Grid>
        </Button>
        <Button x:Name="QuellenKnopf" Style="{StaticResource Knopf}" Margin="8,0,0,0" ToolTip="Datenquellen &amp; Sicherheit - was gelesen wird und wie das Werkzeug abgesichert ist">
          <TextBlock Text="&#xEA18;" FontFamily="{StaticResource Icons}" Margin="0,2,0,0"/>
        </Button>
        <Button x:Name="EinstellungenKnopf" Style="{StaticResource Knopf}" Margin="8,0,0,0" ToolTip="Einstellungen - z. B. das Design der App">
          <TextBlock Text="&#xE713;" FontFamily="{StaticResource Icons}" Margin="0,2,0,0"/>
        </Button>
      </StackPanel>
      <!-- DockPanel statt waagrechtem StackPanel: nur so bekommen Note und Untertitel eine feste Breite und kuerzen mit „…“,
           statt unter die Knoepfe zu laufen -->
      <DockPanel Margin="0,0,16,0">
        <Border DockPanel.Dock="Left" Width="46" Height="46" CornerRadius="13" Background="{DynamicResource AccentSoft}" Margin="0,0,14,0">
          <TextBlock Text="&#xE9D9;" FontFamily="{StaticResource Icons}" FontSize="22" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <StackPanel VerticalAlignment="Center">
          <DockPanel>
            <TextBlock DockPanel.Dock="Left" Text="Systemdiagnose" FontFamily="{DynamicResource Anzeige}" FontSize="24" FontWeight="SemiBold"/>
            <Border x:Name="Note" CornerRadius="11" Padding="10,2,12,3" Margin="14,3,0,0" VerticalAlignment="Center" HorizontalAlignment="Left" Background="{DynamicResource OkBg}" Visibility="Collapsed">
              <DockPanel>
                <Ellipse x:Name="NotePunkt" DockPanel.Dock="Left" Width="8" Height="8" Fill="{DynamicResource Ok}" Margin="0,1,7,0" VerticalAlignment="Center"/>
                <TextBlock x:Name="NoteText" FontSize="12.5" FontWeight="SemiBold" Foreground="{DynamicResource Ok}" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"
                           ToolTip="{Binding Text, RelativeSource={RelativeSource Self}}"/>
              </DockPanel>
            </Border>
          </DockPanel>
          <TextBlock x:Name="Untertitel" Text="Noch keine Prüfung durchgeführt" Foreground="{DynamicResource Muted}" FontSize="13" TextTrimming="CharacterEllipsis"
                     ToolTip="{Binding Text, RelativeSource={RelativeSource Self}}"/>
        </StackPanel>
      </DockPanel>
    </DockPanel>


    <TabControl x:Name="Reiter" Grid.Row="1" Margin="8,10,8,0">
      <!-- In einer Flucht mit den Reitern; der Stufen-Schalter darunter fluchtet mit den Filtern. Nur im Reiter Befunde -->
      <TabControl.Tag>
        <!-- Minimiert nimmt der Fortschritts-Chip genau den Platz von „Pruefung starten“ ein (gleiche Zelle, Breite, Hoehe, Rundung);
             der Startknopf ist dann nur unsichtbar (Hidden), damit die Breite - und der Schalter darunter - in der Flucht bleiben -->
        <Grid>
        <Button x:Name="StartKnopf">
          <Button.Style>
            <Style TargetType="Button" BasedOn="{StaticResource ChipKnopf}">
              <Setter Property="Visibility" Value="Collapsed"/>
              <Style.Triggers>
                <DataTrigger Binding="{Binding SelectedIndex, ElementName=Reiter}" Value="0"><Setter Property="Visibility" Value="Visible"/></DataTrigger>
              </Style.Triggers>
            </Style>
          </Button.Style>
          <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
            <TextBlock Text="&#xE768;" FontFamily="{StaticResource Icons}" Margin="0,3,10,0"/>
            <TextBlock Text="Prüfung starten" FontWeight="SemiBold"/>
          </StackPanel>
        </Button>
        <Button x:Name="LaufMini" Width="{Binding ActualWidth, ElementName=StartKnopf}" Cursor="Hand" Focusable="False"
                Visibility="Collapsed" ToolTip="Fortschritt wieder anzeigen">
          <Button.Template>
            <ControlTemplate TargetType="Button">
              <Border x:Name="Flaeche" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Accent}" BorderThickness="1" CornerRadius="22" Padding="20,0">
                <ContentPresenter VerticalAlignment="Center"/>
              </Border>
              <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Flaeche" Property="Background" Value="{DynamicResource Surface2}"/></Trigger>
              </ControlTemplate.Triggers>
            </ControlTemplate>
          </Button.Template>
          <StackPanel>
            <DockPanel>
              <TextBlock DockPanel.Dock="Right" Text="&#xE740;" FontFamily="{StaticResource Icons}" FontSize="11" Foreground="{DynamicResource Muted}" Margin="10,3,0,0"/>
              <TextBlock x:Name="LaufMiniText" Text="Prüfung" FontSize="13" FontWeight="SemiBold" Foreground="{DynamicResource Text}" TextTrimming="CharacterEllipsis"/>
            </DockPanel>
            <ProgressBar x:Name="LaufMiniBalken" Height="3" Margin="0,5,0,0" IsIndeterminate="True" Maximum="100"/>
          </StackPanel>
        </Button>
        </Grid>
      </TabControl.Tag>
      <TabItem x:Name="ReiterBefunde">
        <TabItem.Header><TextBlock Text="Befunde"/></TabItem.Header>
        <Grid>
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
          <!-- Seit 28.10. schlank: links die Filter (zugleich die Zaehler), rechts „Pruefung starten“ mit dem Stufen-Schalter darunter.
               Die Erklaerung der Stufen klappt per ⓘ (oder Rechtsklick auf den Schalter) darunter auf -->
          <Grid Grid.Row="0" Margin="20,14,20,0">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" VerticalAlignment="Top">
              <WrapPanel>
                <RadioButton x:Name="FilterOffen" GroupName="Filter" Style="{StaticResource Chip}" Margin="0,0,6,6" ToolTip="Kritische Befunde und Warnungen">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="Handlungsbedarf"/><TextBlock x:Name="ZahlOffen" Text="-" FontWeight="SemiBold" Margin="8,0,0,0"/></StackPanel>
                </RadioButton>
                <RadioButton x:Name="FilterInfo" GroupName="Filter" Style="{StaticResource Chip}" Margin="0,0,6,6" ToolTip="Hinweise - lohnt einen Blick, aber nicht dringend">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="Hinweise"/><TextBlock x:Name="ZahlInfo" Text="-" FontWeight="SemiBold" Margin="8,0,0,0"/></StackPanel>
                </RadioButton>
                <RadioButton x:Name="FilterAlle" GroupName="Filter" Style="{StaticResource Chip}" Margin="0,0,6,6" ToolTip="Jede einzelne Prüfung">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="Alle Prüfungen"/><TextBlock x:Name="ZahlAlle" Text="-" FontWeight="SemiBold" Margin="8,0,0,0"/></StackPanel>
                </RadioButton>
              </WrapPanel>
              <TextBlock Foreground="{DynamicResource Muted}" FontSize="12" Margin="6,0,0,0"><Run x:Name="ZahlKritisch" Text="0"/><Run Text=" kritisch · "/><Run x:Name="ZahlWarnung" Text="0"/><Run Text=" Warnungen"/></TextBlock>
            </StackPanel>
            <StackPanel Grid.Column="1" Margin="16,0,0,0" VerticalAlignment="Top" Width="{Binding ActualWidth, ElementName=StartKnopf}">
              <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,4,0,0">
                <CheckBox x:Name="ModusSchalter" Style="{StaticResource Schalter}" ToolTip="Schnell: ca. 20 s, ohne Adminrechte · Gründlich: ca. 5-10 Min., mit Adminrechten - Rechtsklick oder ⓘ für Details"/>
                <ToggleButton x:Name="ModusInfoKnopf" Style="{StaticResource InfoSchalter}" Margin="12,0,0,0" ToolTip="Was die beiden Stufen prüfen"/>
              </StackPanel>
            </StackPanel>
          </Grid>
          <Border x:Name="ModusInfo" Grid.Row="1" Visibility="Collapsed" Margin="20,10,20,0" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="12" Padding="16,10">
            <StackPanel>
              <TextBlock TextWrapping="Wrap" FontSize="12.5" Foreground="{DynamicResource Muted}"><Run Text="Schnell" FontWeight="SemiBold" Foreground="{DynamicResource Text}"/><Run Text="  ca. 20 Sekunden, ohne Adminrechte. Alles, was schnell zu lesen ist: Abstürze, Bluescreens und Treiberprobleme der letzten 14 Tage, Geräte, Laufwerke, Autostart, Einstellungen, Defender-Funde, Netzwerk - und die genaue Reinigungsmessung. Läuft automatisch, wenn der letzte Stand älter als 3 Tage ist."/></TextBlock>
              <TextBlock TextWrapping="Wrap" FontSize="12.5" Foreground="{DynamicResource Muted}" Margin="0,6,0,0"><Run Text="Gründlich" FontWeight="SemiBold" Foreground="{DynamicResource Text}"/><Run Text="  ca. 5-10 Minuten, mit Adminrechten. Zusätzlich alles, was Windows ohne Schreibzugriff hergibt: SMART-Werte, TPM, Systemdateien (SFC), Komponentenspeicher (DISM inkl. Tiefenscan), Dateisystem und TRIM, Startkonfiguration, WMI-Datenbank, Signaturen aller Treiber, Dienste, Aufgaben und Autostart-Programme, Speicherdiagnose - und eine komplette Begehung jedes Laufwerks, die zeigt, wo der Platz steckt. Für Abstürze, nach großen Updates oder einmal im Monat."/></TextBlock>
            </StackPanel>
          </Border>
          <Grid Grid.Row="2">
            <!-- Liste waechst mit dem Fenster mit (390-560), die Karte bekommt den Rest -->
            <Grid.ColumnDefinitions><ColumnDefinition Width="0.3*" MinWidth="400" MaxWidth="560"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Grid.Column="0" Margin="20,14,12,12" Background="{DynamicResource Spalte}" CornerRadius="14" Padding="{DynamicResource SpaltenRand}">
            <Grid>
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <!-- Archiv bewusst leise neben der Suche: selten gebraucht, soll nicht wie ein vierter Filter wirken -->
              <DockPanel Grid.Row="1" Margin="0,6,0,12">
                <RadioButton x:Name="FilterArchiv" DockPanel.Dock="Right" GroupName="Filter" Style="{StaticResource ChipLeise}" Content="Archiv" Margin="6,0,0,0" Visibility="Collapsed"
                             VerticalAlignment="Center" ToolTip="Ausgeblendete Befunde - kommen von selbst zurück, sobald ein neues Ereignis auftritt oder sie schlimmer werden"/>
                <TextBox x:Name="Suche" Style="{StaticResource Suche}" Height="36"/>
              </DockPanel>
              <ListBox x:Name="BefundListe" Grid.Row="2" Padding="0,0,0,24">
                <ListBox.ItemTemplate>
                  <DataTemplate>
                    <Grid x:Name="Zeile" Margin="0,2">
                      <Grid.ColumnDefinitions><ColumnDefinition Width="4"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                      <Border Grid.Column="0" CornerRadius="2" Style="{StaticResource StatusBalken}" Margin="0,6"/>
                      <StackPanel Grid.Column="1" Margin="12,8,8,8">
                        <TextBlock Text="{Binding Pruefung}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                        <TextBlock Text="{Binding Ergebnis}" Foreground="{DynamicResource Muted}" FontSize="12.5" TextTrimming="CharacterEllipsis" Margin="0,2,0,0"/>
                      </StackPanel>
                      <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                        <Border Width="26" Height="26" CornerRadius="13" Background="{DynamicResource AccentSoft}" Margin="0,0,4,0"
                                ToolTip="In der App behebbar" Visibility="{Binding HatBereinigung, Converter={StaticResource BoolSicht}}">
                          <TextBlock Text="&#xE90F;" FontFamily="{StaticResource Icons}" FontSize="12" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <Border Width="26" Height="26" CornerRadius="13" Background="{DynamicResource Surface2}" Margin="0,0,4,0"
                                ToolTip="Archiviert" Visibility="{Binding IstArchiviert, Converter={StaticResource BoolSicht}}">
                          <TextBlock Text="&#xE7B8;" FontFamily="{StaticResource Icons}" FontSize="12" Foreground="{DynamicResource Muted}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                      </StackPanel>
                    </Grid>
                    <DataTemplate.Triggers>
                      <DataTrigger Binding="{Binding IstArchiviert}" Value="True"><Setter TargetName="Zeile" Property="Opacity" Value="0.55"/></DataTrigger>
                    </DataTemplate.Triggers>
                  </DataTemplate>
                </ListBox.ItemTemplate>
                <ListBox.GroupStyle>
                  <GroupStyle>
                    <GroupStyle.HeaderTemplate>
                      <DataTemplate>
                        <TextBlock Text="{Binding Name}" Foreground="{DynamicResource Muted}" FontSize="12" FontWeight="SemiBold" Margin="8,14,0,4"/>
                      </DataTemplate>
                    </GroupStyle.HeaderTemplate>
                  </GroupStyle>
                </ListBox.GroupStyle>
              </ListBox>
              <TextBlock x:Name="ListeLeer" Grid.Row="2" Visibility="Collapsed" Foreground="{DynamicResource Muted}" TextWrapping="Wrap" Margin="8,16,8,0"/>
            </Grid>
            </Border>

            <Border Grid.Column="1" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="16" Margin="0,14,20,12">
              <Grid>
                <ScrollViewer x:Name="Detail" DataContext="{Binding SelectedItem, ElementName=BefundListe}" VerticalScrollBarVisibility="Auto" Visibility="Collapsed"
                             >
                  <!-- Hoechstens 1100 breit: sonst werden die Zeilen auf breiten Fenstern unlesbar lang -->
                  <StackPanel Margin="30,24,30,36" MaxWidth="1100" HorizontalAlignment="Left">
                    <StackPanel Orientation="Horizontal">
                      <Border Style="{StaticResource StatusPille}">
                        <StackPanel Orientation="Horizontal">
                          <TextBlock Text="{Binding StatusGlyph}" FontFamily="{StaticResource Icons}" FontSize="12" Style="{StaticResource StatusText}" Margin="0,2,7,0"/>
                          <TextBlock Text="{Binding Status}" Style="{StaticResource StatusText}" FontWeight="SemiBold" FontSize="12.5"/>
                        </StackPanel>
                      </Border>
                      <TextBlock Text="{Binding BereichTitel}" Foreground="{DynamicResource Muted}" VerticalAlignment="Center" Margin="12,0,0,0" FontSize="13"/>
                    </StackPanel>
                    <TextBox Style="{StaticResource Markierbar}" Text="{Binding Pruefung, Mode=OneWay}" FontFamily="{DynamicResource Anzeige}" FontSize="25" FontWeight="SemiBold" Margin="0,12,0,4"/>
                    <TextBox Style="{StaticResource Markierbar}" Text="{Binding Ergebnis, Mode=OneWay}" FontSize="15"/>
                    <TextBox Style="{StaticResource Markierbar}" Text="{Binding Hinweis, Mode=OneWay}" Foreground="{DynamicResource Muted}" FontSize="13" Margin="0,8,0,0"
                               Visibility="{Binding HatHinweis, Converter={StaticResource BoolSicht}}"/>
                    <Border Background="{DynamicResource WarnBg}" CornerRadius="10" Padding="12,8" Margin="0,10,0,0" Visibility="{Binding HatArchivHinweis, Converter={StaticResource BoolSicht}}">
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#xE7BA;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Warn}" Margin="0,2,8,0"/>
                        <TextBlock Text="{Binding ArchivHinweis}" Foreground="{DynamicResource Warn}" FontSize="13" TextWrapping="Wrap" MaxWidth="620"/>
                      </StackPanel>
                    </Border>

                    <Border Style="{StaticResource Abschnitt}" Margin="0,18,0,0" Visibility="{Binding HatWas, Converter={StaticResource BoolSicht}}">
                      <StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                          <TextBlock Text="&#xE946;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" FontSize="13" Margin="0,2,8,0"/>
                          <TextBlock Text="Was ist das?" Style="{StaticResource AbschnittKopf}"/>
                        </StackPanel>
                        <TextBox Style="{StaticResource Markierbar}" Text="{Binding Was, Mode=OneWay}" TextBlock.LineHeight="21"/>
                      </StackPanel>
                    </Border>
                    <Border Style="{StaticResource Abschnitt}" Visibility="{Binding HatUrsache, Converter={StaticResource BoolSicht}}">
                      <StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                          <TextBlock Text="&#xE82F;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Warn}" FontSize="13" Margin="0,2,8,0"/>
                          <TextBlock Text="Woran liegt es?" Style="{StaticResource AbschnittKopf}"/>
                        </StackPanel>
                        <TextBox Style="{StaticResource Markierbar}" Text="{Binding Ursache, Mode=OneWay}" TextBlock.LineHeight="21"/>
                      </StackPanel>
                    </Border>
                    <Border Style="{StaticResource Abschnitt}" Visibility="{Binding HatMassnahme, Converter={StaticResource BoolSicht}}">
                      <StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                          <TextBlock Text="&#xE8FD;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Ok}" FontSize="13" Margin="0,2,8,0"/>
                          <TextBlock Text="Was ist zu tun?" Style="{StaticResource AbschnittKopf}"/>
                        </StackPanel>
                        <TextBox Style="{StaticResource Markierbar}" Text="{Binding Massnahme, Mode=OneWay}" TextBlock.LineHeight="21"/>
                        <StackPanel Orientation="Horizontal" Margin="0,12,0,0" Visibility="{Binding HatEmpfehlung, Converter={StaticResource BoolSicht}}">
                          <TextBlock Text="Empfehlung" Style="{StaticResource AbschnittKopf}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                          <Border Style="{StaticResource EmpfehlungPille}">
                            <TextBlock Text="{Binding Empfehlung}" Style="{StaticResource EmpfehlungText}"/>
                          </Border>
                        </StackPanel>
                      </StackPanel>
                    </Border>

                    <Border Style="{StaticResource Abschnitt}" BorderBrush="{DynamicResource Accent}" BorderThickness="1" Visibility="{Binding HatLoesungen, Converter={StaticResource BoolSicht}}">
                      <StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                          <TextBlock Text="&#xE90F;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" FontSize="13" Margin="0,2,8,0"/>
                          <TextBlock Text="Lösung" Style="{StaticResource AbschnittKopf}"/>
                        </StackPanel>
                        <ItemsControl ItemsSource="{Binding Loesungen}">
                          <ItemsControl.ItemTemplate>
                            <DataTemplate>
                              <Grid Margin="0,8,0,0">
                                <CheckBox IsChecked="{Binding Ausgewaehlt, Mode=TwoWay}" Visibility="{Binding IstBereinigung, Converter={StaticResource BoolSicht}}">
                                  <StackPanel>
                                    <WrapPanel>
                                      <TextBlock Text="{Binding Titel}" FontWeight="SemiBold" TextWrapping="Wrap"/>
                                      <Border Style="{StaticResource Kennzeichen}"><TextBlock Text="{Binding Kennzeichen}" Style="{StaticResource KennzeichenText}"/></Border>
                                      <Border Style="{StaticResource Kennzeichen}" Background="{DynamicResource InfoBg}" Visibility="{Binding HatDauer, Converter={StaticResource BoolSicht}}">
                                        <TextBlock Text="{Binding Dauer}" Foreground="{DynamicResource Info}" FontSize="11.5" FontWeight="SemiBold"/>
                                      </Border>
                                    </WrapPanel>
                                    <TextBlock Text="{Binding Beschreibung}" Foreground="{DynamicResource Muted}" FontSize="12.5" TextWrapping="Wrap" Margin="0,3,0,0"
                                               Visibility="{Binding HatBeschreibung, Converter={StaticResource BoolSicht}}"/>
                                  </StackPanel>
                                </CheckBox>
                                <Button Style="{StaticResource Knopf}" Tag="{Binding}" HorizontalAlignment="Left" Visibility="{Binding IstSofort, Converter={StaticResource BoolSicht}}">
                                  <StackPanel Orientation="Horizontal">
                                    <TextBlock Text="&#xE8A7;" FontFamily="{StaticResource Icons}" FontSize="12" Margin="0,2,8,0"/>
                                    <TextBlock Text="{Binding Titel}"/>
                                  </StackPanel>
                                </Button>
                              </Grid>
                            </DataTemplate>
                          </ItemsControl.ItemTemplate>
                        </ItemsControl>
                      </StackPanel>
                    </Border>

                    <Border Style="{StaticResource Abschnitt}" Visibility="{Binding KannArchivieren, Converter={StaticResource BoolSicht}}">
                      <StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                          <TextBlock Text="&#xE7B8;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Muted}" FontSize="13" Margin="0,2,8,0"/>
                          <TextBlock Text="{Binding ArchivKopf}" Style="{StaticResource AbschnittKopf}"/>
                        </StackPanel>
                        <TextBox Style="{StaticResource Markierbar}" Text="{Binding ArchivErklaerung, Mode=OneWay}" TextBlock.LineHeight="21" Foreground="{DynamicResource Muted}"/>
                        <Button x:Name="ArchivKnopf" Style="{StaticResource Knopf}" HorizontalAlignment="Left" Margin="0,12,0,0">
                          <StackPanel Orientation="Horizontal">
                            <TextBlock Text="{Binding ArchivKnopfGlyph}" FontFamily="{StaticResource Icons}" Margin="0,2,8,0"/>
                            <TextBlock Text="{Binding ArchivKnopfText}"/>
                          </StackPanel>
                        </Button>
                      </StackPanel>
                    </Border>

                    <Expander Header="{Binding DetailsTitel}" Margin="0,16,0,0" Visibility="{Binding HatDetails, Converter={StaticResource BoolSicht}}">
                      <DataGrid ItemsSource="{Binding Details.DefaultView}" MaxHeight="380" Margin="0,8,0,0"/>
                    </Expander>
                  </StackPanel>
                </ScrollViewer>
                <StackPanel x:Name="DetailLeer" HorizontalAlignment="Center" VerticalAlignment="Center">
                  <TextBlock Text="&#xE9D9;" FontFamily="{StaticResource Icons}" FontSize="40" Foreground="{DynamicResource Linie}" HorizontalAlignment="Center"/>
                  <TextBlock x:Name="DetailLeerText" Text="Links einen Befund auswählen" Foreground="{DynamicResource Muted}" Margin="0,12,0,0" HorizontalAlignment="Center"/>
                </StackPanel>
              </Grid>
            </Border>
          </Grid>

          <Border Grid.Row="3" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="0,1,0,0" Padding="22,12">
            <DockPanel>
              <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                <Button x:Name="EmpfohleneKnopf" Style="{StaticResource Knopf}" Margin="0,0,8,0" ToolTip="Alle Lösungen auswählen, deren Empfehlung „Entfernen“ oder „Beheben“ lautet">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE762;" FontFamily="{StaticResource Icons}" Margin="0,2,8,0"/><TextBlock Text="Empfohlene auswählen"/></StackPanel>
                </Button>
                <Button x:Name="AuswahlLeerenKnopf" Style="{StaticResource Knopf}" Margin="0,0,10,0" Content="Auswahl aufheben"/>
                <Button x:Name="BereinigenKnopf" Style="{StaticResource PrimaerKnopf}" IsEnabled="False">
                  <StackPanel Orientation="Horizontal">
                    <TextBlock Text="&#xE90F;" FontFamily="{StaticResource Icons}" Margin="0,2,9,0"/>
                    <TextBlock x:Name="BereinigenText" Text="Beheben"/>
                  </StackPanel>
                </Button>
              </StackPanel>
              <DockPanel VerticalAlignment="Center" Margin="0,0,16,0">
                <TextBlock DockPanel.Dock="Left" Text="&#xE72E;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Muted}" Margin="0,2,10,0"/>
                <TextBlock x:Name="AuswahlText" Foreground="{DynamicResource Muted}" TextWrapping="Wrap" VerticalAlignment="Center" MaxHeight="40" TextTrimming="CharacterEllipsis"
                           Text="Keine Maßnahme ausgewählt - wähle bei den Befunden unter „Lösung“ aus, was behoben werden soll."/>
              </DockPanel>
            </DockPanel>
          </Border>
        </Grid>
      </TabItem>

      <TabItem x:Name="ReiterReinigung">
        <TabItem.Header><TextBlock x:Name="ReiterReinigungText" Text="Reinigung"/></TabItem.Header>
        <Grid>
          <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
          <Grid Grid.Row="0">
            <!-- Liste hoechstens 1150 breit, sonst stehen Name und Groesse auf breiten Fenstern zu weit auseinander -->
            <Grid.ColumnDefinitions><ColumnDefinition Width="*" MaxWidth="1150"/><ColumnDefinition Width="360"/></Grid.ColumnDefinitions>
            <Grid Grid.Column="0" Margin="20,14,12,12">
              <!-- Zwei Ansichten wie die Befund-Filter: was man immer loeschen kann (Vorauswahl im Katalog) und was man sich
                   vorher ansieht. Leere Kategorien sind ausgeblendet - nach dem Reinigen ist „Bedenkenlos“ wirklich leer -->
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <WrapPanel Grid.Row="0" Margin="0,0,0,10">
                <RadioButton x:Name="ReinigungFilterImmer" GroupName="ReinigungFilter" Style="{StaticResource Chip}" Margin="0,0,6,0" IsChecked="True"
                             ToolTip="Caches, temporäre Dateien, Update-Reste - kann man immer löschen, Windows und Programme legen sie bei Bedarf neu an">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="Bedenkenlos löschen"/><TextBlock x:Name="ZahlReinigungImmer" Text="-" FontWeight="SemiBold" Margin="8,0,0,0"/></StackPanel>
                </RadioButton>
                <RadioButton x:Name="ReinigungFilterAnsehen" GroupName="ReinigungFilter" Style="{StaticResource Chip}" Margin="0,0,6,0"
                             ToolTip="Könnte man löschen, muss man aber nicht - z. B. Shader-Cache (Spiele kompilieren neu), Treiber-Installer, Fehlerberichte, Papierkorb">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="Zum Ansehen"/><TextBlock x:Name="ZahlReinigungAnsehen" Text="-" FontWeight="SemiBold" Margin="8,0,0,0"/></StackPanel>
                </RadioButton>
              </WrapPanel>
              <Border Grid.Row="1" Background="{DynamicResource Spalte}" CornerRadius="14"/>
              <ScrollViewer x:Name="ReinigungScroll" Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="{DynamicResource SpaltenRand}">
                <ItemsControl x:Name="ReinigungListe" Margin="0,0,0,28">
                  <ItemsControl.ItemTemplate>
                    <DataTemplate>
                      <Border x:Name="Zeile" CornerRadius="10" Padding="12,9" Margin="0,1,6,1" Background="Transparent">
                        <Grid>
                          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                          <CheckBox IsChecked="{Binding Ausgewaehlt, Mode=TwoWay}" IsEnabled="{Binding IstAuswaehlbar}">
                            <StackPanel>
                              <WrapPanel>
                                <TextBlock Text="{Binding Name}" FontWeight="SemiBold"/>
                                <Border Style="{StaticResource WarnChip}" Visibility="{Binding HatWarnung, Converter={StaticResource BoolSicht}}">
                                  <TextBlock Text="{Binding WarnText}" Foreground="{DynamicResource Warn}" FontSize="11.5" FontWeight="SemiBold"/>
                                </Border>
                              </WrapPanel>
                              <TextBlock Text="{Binding Beschreibung}" Foreground="{DynamicResource Muted}" FontSize="12.5" TextWrapping="Wrap" Margin="0,2,0,0"/>
                              <TextBlock Text="{Binding Hinweis}" Foreground="{DynamicResource Warn}" FontSize="12" TextWrapping="Wrap" Margin="0,3,0,0"
                                         Visibility="{Binding HatHinweis, Converter={StaticResource BoolSicht}}"/>
                            </StackPanel>
                          </CheckBox>
                          <StackPanel Grid.Column="1" Margin="18,0,2,0" Width="175">
                            <TextBlock Text="{Binding GroesseText}" FontFamily="{DynamicResource Anzeige}" FontSize="16" FontWeight="SemiBold" HorizontalAlignment="Right" TextAlignment="Right" TextWrapping="Wrap"/>
                            <Border Height="4" CornerRadius="2" Background="{DynamicResource Surface2}" Width="150" HorizontalAlignment="Right" Margin="0,5,0,4">
                              <Border Height="4" CornerRadius="2" Background="{DynamicResource Accent}" Width="{Binding BalkenBreite}" HorizontalAlignment="Left"/>
                            </Border>
                            <TextBlock Text="{Binding DateienText}" Foreground="{DynamicResource Muted}" FontSize="12" HorizontalAlignment="Right" TextAlignment="Right" TextWrapping="Wrap"/>
                          </StackPanel>
                        </Grid>
                      </Border>
                      <DataTemplate.Triggers>
                        <Trigger SourceName="Zeile" Property="IsMouseOver" Value="True"><Setter TargetName="Zeile" Property="Background" Value="{DynamicResource SurfaceHover}"/></Trigger>
                        <DataTrigger Binding="{Binding IstLeer}" Value="True"><Setter TargetName="Zeile" Property="Opacity" Value="0.5"/></DataTrigger>
                      </DataTemplate.Triggers>
                    </DataTemplate>
                  </ItemsControl.ItemTemplate>
                  <ItemsControl.GroupStyle>
                    <GroupStyle>
                      <GroupStyle.HeaderTemplate>
                        <DataTemplate>
                          <!-- Summe rechtsbuendig auf Hoehe der Groessenspalte (Zeile: Rand 6 + Innenabstand 12 + 2) -->
                          <DockPanel Margin="12,14,20,4">
                            <TextBlock DockPanel.Dock="Right" Text="{Binding Items[0].GruppeSumme}" Foreground="{DynamicResource Muted}" FontSize="12" FontWeight="SemiBold" Margin="12,0,0,0"/>
                            <TextBlock Text="{Binding Name}" Foreground="{DynamicResource Muted}" FontSize="12" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                          </DockPanel>
                        </DataTemplate>
                      </GroupStyle.HeaderTemplate>
                    </GroupStyle>
                  </ItemsControl.GroupStyle>
                </ItemsControl>
              </ScrollViewer>
              <StackPanel x:Name="ReinigungLeer" Grid.Row="1" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed" MaxWidth="440">
                <TextBlock Text="&#xE74D;" FontFamily="{StaticResource Icons}" FontSize="40" Foreground="{DynamicResource Linie}" HorizontalAlignment="Center"/>
                <TextBlock x:Name="ReinigungLeerText" Text="Noch nicht gemessen." Foreground="{DynamicResource Muted}" TextWrapping="Wrap" TextAlignment="Center" Margin="0,12,0,16"/>
                <Button x:Name="AnalysierenLeerKnopf" Style="{StaticResource PrimaerKnopf}" HorizontalAlignment="Center">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE721;" FontFamily="{StaticResource Icons}" Margin="0,2,9,0"/><TextBlock Text="Jetzt analysieren"/></StackPanel>
                </Button>
              </StackPanel>
            </Grid>

            <Border Grid.Column="1" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="16" Margin="0,14,20,12">
              <ScrollViewer x:Name="ReinigungInfoScroll" VerticalScrollBarVisibility="Auto">
                <StackPanel Margin="24,20,24,28">
                  <TextBlock Text="Freigebbar" Style="{StaticResource AbschnittKopf}"/>
                  <TextBlock x:Name="ReinigungSumme" Text="-" FontFamily="{DynamicResource Anzeige}" FontSize="38" FontWeight="SemiBold" Foreground="{DynamicResource Accent}"/>
                  <TextBlock x:Name="ReinigungSummeText" Foreground="{DynamicResource Muted}" FontSize="13" TextWrapping="Wrap"/>
                  <Border Style="{StaticResource Abschnitt}" Margin="0,14,0,0">
                    <StackPanel>
                      <TextBlock Text="Ausgewählt" Style="{StaticResource AbschnittKopf}"/>
                      <TextBlock x:Name="ReinigungAuswahlGross" Text="-" FontFamily="{DynamicResource Anzeige}" FontSize="22" FontWeight="SemiBold" Margin="0,2,0,0"/>
                      <TextBlock x:Name="ReinigungAuswahlText" Foreground="{DynamicResource Muted}" FontSize="12.5" TextWrapping="Wrap"/>
                      <TextBlock x:Name="ReinigungSchliessenText" Foreground="{DynamicResource Warn}" FontSize="12.5" TextWrapping="Wrap" Margin="0,6,0,0" Visibility="Collapsed"/>
                    </StackPanel>
                  </Border>
                  <DockPanel Margin="0,14,0,0">
                    <Button x:Name="AnalysierenKnopf" DockPanel.Dock="Left" Style="{StaticResource Knopf}" ToolTip="Alle Kategorien neu messen (ohne Adminrechte, wenige Sekunden)">
                      <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE72C;" FontFamily="{StaticResource Icons}" Margin="0,2,8,0"/><TextBlock Text="Neu messen"/></StackPanel>
                    </Button>
                    <TextBlock x:Name="ReinigungStand" Foreground="{DynamicResource Muted}" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center" Margin="12,0,0,0"/>
                  </DockPanel>
                  <!-- Eingeklappt, damit die Spalte bei normaler Fenstergroesse nicht unten abgeschnitten wirkt -->
                  <Expander Header="Was bleibt erhalten, wie wird gelöscht?" Margin="0,16,0,0">
                  <Border Style="{StaticResource Abschnitt}" Margin="0,8,0,0">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                        <TextBlock Text="&#xEA18;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Ok}" FontSize="13" Margin="0,2,8,0"/>
                        <TextBlock Text="Bleibt immer erhalten" Style="{StaticResource AbschnittKopf}"/>
                      </StackPanel>
                      <TextBlock TextWrapping="Wrap" Foreground="{DynamicResource Muted}" FontSize="12.5" LineHeight="19">
                        <Run Text="• Cookies, Passwörter, Verlauf, offene Tabs und Erweiterungen"/><LineBreak/>
                        <Run Text="• Dokumente, Bilder, Downloads, Desktop - diese Ordner sind gesperrt"/><LineBreak/>
                        <Run Text="• Dateien, die gerade benutzt werden, und junge Temp-Dateien"/><LineBreak/>
                        <Run Text="• Kategorien, deren Programm läuft (z. B. Chrome) - die werden übersprungen"/>
                      </TextBlock>
                      <StackPanel Orientation="Horizontal" Margin="0,12,0,6">
                        <TextBlock Text="&#xE72E;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" FontSize="13" Margin="0,2,8,0"/>
                        <TextBlock Text="So wird gelöscht" Style="{StaticResource AbschnittKopf}"/>
                      </StackPanel>
                      <TextBlock TextWrapping="Wrap" Foreground="{DynamicResource Muted}" FontSize="12.5" LineHeight="19">
                        <Run Text="Nur in den fest hinterlegten Ordnern jeder Kategorie, Datei für Datei. Vor jedem Löschen wird der tatsächliche Speicherort geprüft - Ordner-Verknüpfungen können nichts nach außen umleiten. Gelöschtes ist nicht umkehrbar."/>
                      </TextBlock>
                    </StackPanel>
                  </Border>
                  </Expander>
                </StackPanel>
              </ScrollViewer>
            </Border>
          </Grid>

          <Border Grid.Row="1" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="0,1,0,0" Padding="22,12">
            <DockPanel>
              <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                <Button x:Name="ReinigungEmpfohlenKnopf" Style="{StaticResource Knopf}" Margin="0,0,8,0" ToolTip="Nur die Kategorien ohne Nachteile auswählen">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE762;" FontFamily="{StaticResource Icons}" Margin="0,2,8,0"/><TextBlock Text="Empfohlene"/></StackPanel>
                </Button>
                <Button x:Name="ReinigungKeineKnopf" Style="{StaticResource Knopf}" Margin="0,0,10,0" Content="Keine"/>
                <Button x:Name="ReinigenKnopf" Style="{StaticResource PrimaerKnopf}" IsEnabled="False">
                  <StackPanel Orientation="Horizontal">
                    <TextBlock Text="&#xE74D;" FontFamily="{StaticResource Icons}" Margin="0,2,9,0"/>
                    <TextBlock x:Name="ReinigenText" Text="Reinigen"/>
                  </StackPanel>
                </Button>
              </StackPanel>
              <DockPanel VerticalAlignment="Center" Margin="0,0,16,0">
                <TextBlock DockPanel.Dock="Left" Text="&#xE946;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Muted}" Margin="0,2,10,0"/>
                <TextBlock x:Name="ReinigungFussText" Foreground="{DynamicResource Muted}" TextWrapping="Wrap" VerticalAlignment="Center" MaxHeight="40" TextTrimming="CharacterEllipsis"
                           Text="Vorausgewählt sind nur Kategorien ohne Nachteile. Hinweise in Orange vor dem Ankreuzen lesen."/>
              </DockPanel>
            </DockPanel>
          </Border>
        </Grid>
      </TabItem>

      <!-- Verlauf = Massnahmen (Behebungen/Reinigungen), Systemaenderungen und Entwicklung; vorher drei eigene Reiter -->
      <TabItem x:Name="ReiterVerlauf">
        <TabItem.Header><TextBlock Text="Verlauf"/></TabItem.Header>
        <Grid>
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <WrapPanel Margin="20,14,20,0">
            <RadioButton x:Name="VerlaufMassnahmen" GroupName="VerlaufAnsicht" Style="{StaticResource Chip}" Content="Maßnahmen" Margin="0,0,6,0" IsChecked="True"
                         ToolTip="Behebungen und Reinigungen - umkehrbare lassen sich rückgängig machen"/>
            <RadioButton x:Name="VerlaufAenderungen" GroupName="VerlaufAnsicht" Style="{StaticResource Chip}" Content="Systemänderungen" Margin="0,0,6,0"
                         ToolTip="Was sich am System seit der vorigen Prüfung geändert hat (Dienste, Autostart, Treiber …)"/>
            <RadioButton x:Name="VerlaufEntwicklung" GroupName="VerlaufAnsicht" Style="{StaticResource Chip}" Content="Entwicklung" Margin="0,0,6,0"
                         ToolTip="Zuverlässigkeitsindex und freier Speicher über alle bisherigen Prüfungen"/>
          </WrapPanel>
          <Grid Grid.Row="1">
            <Grid x:Name="AnsichtMassnahmen" Margin="20,12,20,18">
              <Grid.ColumnDefinitions><ColumnDefinition Width="340"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <!-- Wie die Befundliste: abgedunkelte, abgerundete Flaeche, je Eintrag ein Symbol der Art -->
              <Border Grid.Column="0" Margin="0,0,14,0" Background="{DynamicResource Spalte}" CornerRadius="14" Padding="{DynamicResource SpaltenRand}">
                <Grid>
                  <ListBox x:Name="VerlaufListe">
                    <ListBox.ItemTemplate>
                      <DataTemplate>
                        <Grid Margin="4,9,8,9">
                          <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                          <TextBlock Text="{Binding Glyph}" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" FontSize="14" Margin="0,2,0,0"/>
                          <StackPanel Grid.Column="1">
                            <TextBlock Text="{Binding Titel}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                            <TextBlock Text="{Binding Zusammenfassung}" Foreground="{DynamicResource Muted}" FontSize="12.5" Margin="0,2,0,0" TextWrapping="Wrap"/>
                          </StackPanel>
                        </Grid>
                      </DataTemplate>
                    </ListBox.ItemTemplate>
                  </ListBox>
                  <TextBlock x:Name="VerlaufLeer" Text="Noch nichts behoben oder gereinigt." Foreground="{DynamicResource Muted}" Margin="10,10,0,0" Visibility="Collapsed"/>
                </Grid>
              </Border>
              <Border Grid.Column="1" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="16"
                      DataContext="{Binding SelectedItem, ElementName=VerlaufListe}">
                <Grid Margin="26,22">
                  <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                  <TextBlock Text="{Binding Titel}" FontFamily="{DynamicResource Anzeige}" FontSize="21" FontWeight="SemiBold"/>
                  <StackPanel Grid.Row="1" Margin="0,6,0,14">
                    <TextBlock Text="{Binding Wiederherstellungspunkt, StringFormat='Wiederherstellungspunkt: {0}'}" Foreground="{DynamicResource Muted}" FontSize="13" TextWrapping="Wrap"/>
                    <TextBlock Text="{Binding Fehler}" Foreground="{DynamicResource Crit}" TextWrapping="Wrap" Margin="0,6,0,0" Visibility="{Binding HatFehler, Converter={StaticResource BoolSicht}}"/>
                  </StackPanel>
                  <!-- Abschnitt wie in der Befund-Karte („Was ist zu tun?“): Flaeche Surface2, Kopf mit Symbol -->
                  <Border Grid.Row="2" Background="{DynamicResource Surface2}" CornerRadius="12" Padding="16,12,16,6" VerticalAlignment="Top">
                  <DockPanel>
                    <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,4">
                      <TextBlock Text="&#xE9D5;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" FontSize="13" Margin="0,2,8,0"/>
                      <TextBlock Text="Einzelne Maßnahmen" Style="{StaticResource AbschnittKopf}"/>
                    </StackPanel>
                  <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <ItemsControl ItemsSource="{Binding Aktionen}">
                      <ItemsControl.ItemTemplate>
                        <DataTemplate>
                          <Grid Margin="0,7">
                            <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                            <TextBlock Text="{Binding Glyph}" Style="{StaticResource ErgebnisGlyph}" Margin="0,2,0,0"/>
                            <StackPanel Grid.Column="1">
                              <TextBlock Text="{Binding Titel}" FontWeight="SemiBold" TextWrapping="Wrap"/>
                              <TextBlock Text="{Binding Meldung}" Foreground="{DynamicResource Muted}" FontSize="12.5" TextWrapping="Wrap" Margin="0,2,0,0"/>
                            </StackPanel>
                          </Grid>
                        </DataTemplate>
                      </ItemsControl.ItemTemplate>
                    </ItemsControl>
                  </ScrollViewer>
                  </DockPanel>
                  </Border>
                  <Button x:Name="RueckgaengigKnopf" Grid.Row="3" Style="{StaticResource Knopf}" HorizontalAlignment="Left" Margin="0,12,0,0"
                          Visibility="{Binding KannRueckgaengig, Converter={StaticResource BoolSicht}}">
                    <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE7A7;" FontFamily="{StaticResource Icons}" Margin="0,2,8,0"/><TextBlock Text="Umkehrbare Maßnahmen rückgängig machen"/></StackPanel>
                  </Button>
                </Grid>
              </Border>
            </Grid>
            <Grid x:Name="AnsichtAenderungen" Margin="20,12,20,18" Visibility="Collapsed">
              <!-- Karte wie die Befund-Karte: Kopf mit Symbol + grauer Abschnittsschrift, Tabelle ohne eigenen Rahmen -->
              <Border Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="16" Padding="18,14,18,12">
                <DockPanel>
                  <DockPanel DockPanel.Dock="Top" Margin="2,0,2,10">
                    <TextBlock Text="&#xE81C;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" FontSize="13" Margin="0,2,8,0"/>
                    <TextBlock x:Name="AenderungenInfo" Style="{StaticResource AbschnittKopf}" TextWrapping="Wrap"/>
                  </DockPanel>
                  <DataGrid x:Name="AenderungenGrid" Background="Transparent" BorderThickness="0"/>
                </DockPanel>
              </Border>
            </Grid>
            <Grid x:Name="AnsichtEntwicklung" Margin="20,12,20,18" Visibility="Collapsed">
              <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <Border Grid.Row="0" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="16" Padding="18,14,18,12" Margin="0,0,0,7">
                <DockPanel>
                  <DockPanel DockPanel.Dock="Top" Margin="2,0,2,8">
                    <TextBlock x:Name="EntwicklungInfoIndex" DockPanel.Dock="Right" Foreground="{DynamicResource Muted}" FontSize="12.5" VerticalAlignment="Center"/>
                    <TextBlock Text="&#xE9D9;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" FontSize="13" Margin="0,2,8,0"/>
                    <TextBlock Text="Zuverlässigkeitsindex von Windows" Style="{StaticResource AbschnittKopf}" TextTrimming="CharacterEllipsis"/>
                  </DockPanel>
                  <Canvas x:Name="KurveIndex" ClipToBounds="True"/>
                </DockPanel>
              </Border>
              <Border Grid.Row="1" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="16" Padding="18,14,18,12" Margin="0,7,0,0">
                <DockPanel>
                  <DockPanel DockPanel.Dock="Top" Margin="2,0,2,8">
                    <TextBlock x:Name="EntwicklungInfoSpeicher" DockPanel.Dock="Right" Foreground="{DynamicResource Muted}" FontSize="12.5" VerticalAlignment="Center"/>
                    <TextBlock Text="&#xEDA2;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" FontSize="13" Margin="0,2,8,0"/>
                    <TextBlock Text="Freier Speicher je Laufwerk (in %)" Style="{StaticResource AbschnittKopf}" TextTrimming="CharacterEllipsis"/>
                  </DockPanel>
                  <Canvas x:Name="KurveSpeicher" ClipToBounds="True"/>
                </DockPanel>
              </Border>
            </Grid>
          </Grid>
        </Grid>
      </TabItem>

      <!-- Ohne sichtbaren Reiter: geoeffnet ueber das Schild-Symbol oben, zurueck ueber „Zurück“ oder jeden Reiter -->
      <TabItem x:Name="ReiterQuellen" Visibility="Collapsed">
        <TabItem.Header><TextBlock Text="Datenquellen &amp; Sicherheit"/></TabItem.Header>
        <Grid Margin="20,14,20,18">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="380"/></Grid.ColumnDefinitions>
          <DockPanel Grid.ColumnSpan="2" Margin="0,0,0,12">
            <Button x:Name="QuellenZurueckKnopf" Style="{StaticResource Knopf}" Margin="0,0,14,0" ToolTip="Zurück zum vorigen Reiter  ·  Esc">
              <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE72B;" FontFamily="{StaticResource Icons}" Margin="0,2,8,0"/><TextBlock Text="Zurück"/></StackPanel>
            </Button>
            <TextBlock Text="Datenquellen &amp; Sicherheit" FontFamily="{DynamicResource Anzeige}" FontSize="19" FontWeight="SemiBold" VerticalAlignment="Center"/>
          </DockPanel>
          <Border Grid.Row="1" Grid.Column="0" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="16" Padding="18,14,18,12">
            <DockPanel>
              <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="2,0,2,10">
                <TextBlock Text="&#xE8F1;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" FontSize="13" Margin="0,2,8,0"/>
                <TextBlock Text="Gelesene Datenquellen der letzten Prüfung" Style="{StaticResource AbschnittKopf}"/>
              </StackPanel>
              <DataGrid x:Name="QuellenGrid" Background="Transparent" BorderThickness="0"/>
            </DockPanel>
          </Border>
          <Border Grid.Row="1" Grid.Column="1" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="16" Margin="14,0,0,0" Padding="22,18">
            <StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                <TextBlock Text="&#xEA18;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Ok}" FontSize="16" Margin="0,3,10,0"/>
                <TextBlock Text="So ist das Werkzeug abgesichert" FontWeight="SemiBold" FontSize="15"/>
              </StackPanel>
              <TextBlock TextWrapping="Wrap" Foreground="{DynamicResource Muted}" FontSize="13" LineHeight="20">
                <Run Text="Diagnose (nur lesend)" FontWeight="SemiBold" Foreground="{DynamicResource Text}"/><LineBreak/>
                <Run Text="• prüft vor jedem Start ihren eigenen Code und startet nicht, wenn darin etwas Schreibendes steht"/><LineBreak/>
                <Run Text="• WhatIf-Netz: ändernde Befehle würden nur simulieren"/><LineBreak/>
                <Run Text="• Registry nur mit Lese-Handles, externe Programme nur freigegeben und lesend"/><LineBreak/>
                <Run Text="• jede Datenquelle wird genau einmal gelesen (Liste links)"/><LineBreak/><LineBreak/>
                <Run Text="Beheben (getrennter Teil)" FontWeight="SemiBold" Foreground="{DynamicResource Text}"/><LineBreak/>
                <Run Text="• nur Aktionen aus einem festen Katalog, jede von dir bestätigt"/><LineBreak/>
                <Run Text="• prüft jede Aktion unmittelbar vorher erneut am System"/><LineBreak/>
                <Run Text="• Windows-Dienste und Microsoft-Aufgaben werden nie angefasst"/><LineBreak/>
                <Run Text="• Treiberpakete nur mit Wiederherstellungspunkt"/><LineBreak/>
                <Run Text="• umkehrbare Maßnahmen lassen sich rückgängig machen"/><LineBreak/><LineBreak/>
                <Run Text="Reinigung" FontWeight="SemiBold" Foreground="{DynamicResource Text}"/><LineBreak/>
                <Run Text="• nur Ordner aus der festen Kategorienliste (Reinigung.psd1) - die App übergibt nie Pfade"/><LineBreak/>
                <Run Text="• Laufwerkswurzel, Windows-, Programm- und persönliche Ordner sind hart gesperrt"/><LineBreak/>
                <Run Text="• Datei für Datei mit Prüfung des tatsächlichen Speicherorts - Verknüpfungen werden nie verfolgt"/>
              </TextBlock>
            </StackPanel>
          </Border>
        </Grid>
      </TabItem>
    </TabControl>

    <Grid x:Name="Overlay" Grid.RowSpan="4" Background="{DynamicResource Overlay}" Visibility="Collapsed">
      <Border Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="18" Padding="32,28" Width="500"
              HorizontalAlignment="Center" VerticalAlignment="Center">
        <StackPanel>
          <DockPanel>
          <!-- Minimieren: der Lauf geht weiter, oben rechts bleibt ein kleiner Fortschritts-Chip (Klick holt das Fenster zurueck) -->
          <Button x:Name="OverlayMinimieren" DockPanel.Dock="Right" VerticalAlignment="Top" Margin="12,-12,-18,0" Width="36" Height="30" Cursor="Hand"
                  Focusable="False" ToolTip="Minimieren - die App bleibt benutzbar, der Lauf geht weiter">
            <Button.Template>
              <ControlTemplate TargetType="Button">
                <Border x:Name="Flaeche" CornerRadius="8" Background="Transparent">
                  <TextBlock Text="&#xE921;" FontFamily="{StaticResource Icons}" FontSize="12" Foreground="{DynamicResource Muted}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <ControlTemplate.Triggers>
                  <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Flaeche" Property="Background" Value="{DynamicResource Surface2}"/></Trigger>
                </ControlTemplate.Triggers>
              </ControlTemplate>
            </Button.Template>
          </Button>
          <StackPanel Orientation="Horizontal">
            <TextBlock x:Name="OverlayIcon" Text="&#xE9D9;" FontFamily="{StaticResource Icons}" FontSize="20" Foreground="{DynamicResource Accent}" Margin="0,4,12,0"/>
            <TextBlock x:Name="OverlayTitel" Text="Prüfung läuft …" FontFamily="{DynamicResource Anzeige}" FontSize="21" FontWeight="SemiBold"/>
          </StackPanel>
          </DockPanel>
          <TextBlock x:Name="OverlayText" Foreground="{DynamicResource Muted}" Margin="0,8,0,18" TextWrapping="Wrap" MinHeight="20"/>
          <!-- Oben: wie weit (Prozent, gewichtet nach der Dauer jeder Quelle im letzten Lauf derselben Stufe).
               Unten: lebt er? Laeuft nur, solange der Pruefprozess oder seine Kindprozesse (DISM, SFC ...) wirklich CPU-Zeit verbrauchen -->
          <DockPanel>
            <TextBlock x:Name="OverlayProzent" DockPanel.Dock="Right" Text="" FontSize="12.5" FontWeight="SemiBold" Margin="12,-4,0,0" MinWidth="36" TextAlignment="Right"/>
            <ProgressBar x:Name="OverlayFortschritt" IsIndeterminate="True" Maximum="100"/>
          </DockPanel>
          <ProgressBar x:Name="OverlayAktiv" IsIndeterminate="True" Maximum="100" Height="3" Margin="0,10,48,0" Opacity="0.7"/>
          <DockPanel Margin="0,8,0,0">
            <TextBlock x:Name="OverlayZeit" DockPanel.Dock="Right" Text="00:00" Foreground="{DynamicResource Muted}" FontSize="12.5"/>
            <TextBlock x:Name="OverlayAktivText" Text="arbeitet" Foreground="{DynamicResource Muted}" FontSize="12.5"/>
          </DockPanel>
          <!-- Abbrechen (auch Esc): zwei gekreuzte rote Lichtschwerter, nur bei Pruefung/Analyse (die aendern nichts) -->
          <Button x:Name="OverlayAbbrechen" Width="60" Height="60" Margin="0,14,0,-10" HorizontalAlignment="Center" Cursor="Hand"
                  Focusable="False" ToolTip="Abbrechen (Esc)" Visibility="Collapsed">
            <Button.Template>
              <ControlTemplate TargetType="Button">
                <Grid Background="Transparent">
                  <!-- Wie auf den Hintergrundbildern: weisser Kern, scharfe rote Kante, Metallgriff; nur die Klingen leuchten -->
                  <Grid x:Name="Schwerter">
                    <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" RenderTransformOrigin="0.5,0.5">
                      <StackPanel.RenderTransform><RotateTransform Angle="40"/></StackPanel.RenderTransform>
                      <Border x:Name="Klinge1" Width="5" Height="40" CornerRadius="2.5,2.5,0.5,0.5">
                        <Border.Background>
                          <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                            <GradientStop Color="#D80E18" Offset="0"/><GradientStop Color="#FF4A4A" Offset="0.22"/><GradientStop Color="#FFF8F8" Offset="0.5"/>
                            <GradientStop Color="#FF4A4A" Offset="0.78"/><GradientStop Color="#D80E18" Offset="1"/>
                          </LinearGradientBrush>
                        </Border.Background>
                        <Border.Effect><DropShadowEffect Color="#FF1A1A" BlurRadius="4" ShadowDepth="0" Opacity="0.7"/></Border.Effect>
                      </Border>
                      <Border Width="8" Height="2.5" CornerRadius="0.5" Background="#24282D"/>
                      <Border Width="7" Height="13" CornerRadius="0,0,1.5,1.5" BorderBrush="#24282D" BorderThickness="0.6">
                        <Border.Background>
                          <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                            <GradientStop Color="#5A6067" Offset="0"/><GradientStop Color="#E6EAEE" Offset="0.38"/><GradientStop Color="#8C9299" Offset="0.7"/><GradientStop Color="#454A50" Offset="1"/>
                          </LinearGradientBrush>
                        </Border.Background>
                        <StackPanel VerticalAlignment="Center">
                          <Border Height="1" Background="#24282D" Margin="0,1.3"/><Border Height="1" Background="#24282D" Margin="0,1.3"/><Border Height="1" Background="#24282D" Margin="0,1.3"/>
                        </StackPanel>
                      </Border>
                    </StackPanel>
                    <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" RenderTransformOrigin="0.5,0.5">
                      <StackPanel.RenderTransform><RotateTransform Angle="-40"/></StackPanel.RenderTransform>
                      <Border x:Name="Klinge2" Width="5" Height="40" CornerRadius="2.5,2.5,0.5,0.5">
                        <Border.Background>
                          <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                            <GradientStop Color="#D80E18" Offset="0"/><GradientStop Color="#FF4A4A" Offset="0.22"/><GradientStop Color="#FFF8F8" Offset="0.5"/>
                            <GradientStop Color="#FF4A4A" Offset="0.78"/><GradientStop Color="#D80E18" Offset="1"/>
                          </LinearGradientBrush>
                        </Border.Background>
                        <Border.Effect><DropShadowEffect Color="#FF1A1A" BlurRadius="4" ShadowDepth="0" Opacity="0.7"/></Border.Effect>
                      </Border>
                      <Border Width="8" Height="2.5" CornerRadius="0.5" Background="#24282D"/>
                      <Border Width="7" Height="13" CornerRadius="0,0,1.5,1.5" BorderBrush="#24282D" BorderThickness="0.6">
                        <Border.Background>
                          <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                            <GradientStop Color="#5A6067" Offset="0"/><GradientStop Color="#E6EAEE" Offset="0.38"/><GradientStop Color="#8C9299" Offset="0.7"/><GradientStop Color="#454A50" Offset="1"/>
                          </LinearGradientBrush>
                        </Border.Background>
                        <StackPanel VerticalAlignment="Center">
                          <Border Height="1" Background="#24282D" Margin="0,1.3"/><Border Height="1" Background="#24282D" Margin="0,1.3"/><Border Height="1" Background="#24282D" Margin="0,1.3"/>
                        </StackPanel>
                      </Border>
                    </StackPanel>
                  </Grid>
                </Grid>
                <ControlTemplate.Triggers>
                  <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Klinge1" Property="Effect"><Setter.Value><DropShadowEffect Color="#FF2020" BlurRadius="16" ShadowDepth="0" Opacity="1"/></Setter.Value></Setter>
                    <Setter TargetName="Klinge2" Property="Effect"><Setter.Value><DropShadowEffect Color="#FF2020" BlurRadius="16" ShadowDepth="0" Opacity="1"/></Setter.Value></Setter>
                  </Trigger>
                  <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Schwerter" Property="Opacity" Value="0.35"/></Trigger>
                </ControlTemplate.Triggers>
              </ControlTemplate>
            </Button.Template>
          </Button>
        </StackPanel>
      </Border>
    </Grid>
  </Grid>
</Window>
'@

# ============================================================================
#  Dialoge
# ============================================================================
$Script:BestaetigungXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Bestätigen" Width="660" SizeToContent="Height" MaxHeight="820" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False" Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="False">
  <Grid x:Name="Wurzel" Background="{DynamicResource Bg}" Margin="0">
    <StackPanel Margin="30,26,30,24">
      <StackPanel Orientation="Horizontal">
        <Border Width="40" Height="40" CornerRadius="12" Background="{DynamicResource AccentSoft}" Margin="0,0,14,0">
          <TextBlock x:Name="Icon" Text="&#xE90F;" FontFamily="{StaticResource Icons}" FontSize="18" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <TextBlock x:Name="Titel" FontFamily="{DynamicResource Anzeige}" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </StackPanel>
      <TextBlock x:Name="Einleitung" Text="Nach „Jetzt ausführen“ werden die Maßnahmen mit Adminrechten ausgeführt. Jede Maßnahme wird unmittelbar vor der Ausführung noch einmal am System geprüft - ist das Problem nicht mehr da, wird sie übersprungen."
                 Foreground="{DynamicResource Muted}" TextWrapping="Wrap" Margin="0,14,0,14" LineHeight="20"/>
      <Border Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="14" Padding="18,8">
        <ScrollViewer MaxHeight="380" VerticalScrollBarVisibility="Auto">
          <ItemsControl x:Name="Liste">
            <ItemsControl.ItemTemplate>
              <DataTemplate>
                <Grid Margin="0,8">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="26"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock Text="&#xE73E;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" Margin="0,3,0,0"/>
                  <StackPanel Grid.Column="1">
                    <WrapPanel>
                      <TextBlock Text="{Binding Titel}" FontWeight="SemiBold" TextWrapping="Wrap"/>
                      <Border Style="{StaticResource Kennzeichen}"><TextBlock Text="{Binding Kennzeichen}" Style="{StaticResource KennzeichenText}"/></Border>
                      <Border Style="{StaticResource Kennzeichen}" Background="{DynamicResource InfoBg}" Visibility="{Binding HatDauer, Converter={StaticResource BoolSicht}}">
                        <TextBlock Text="{Binding Dauer}" Foreground="{DynamicResource Info}" FontSize="11.5" FontWeight="SemiBold"/>
                      </Border>
                    </WrapPanel>
                    <TextBlock Text="{Binding BefundTitel}" Foreground="{DynamicResource Muted}" FontSize="12.5" Margin="0,2,0,0"/>
                  </StackPanel>
                </Grid>
              </DataTemplate>
            </ItemsControl.ItemTemplate>
          </ItemsControl>
        </ScrollViewer>
      </Border>
      <Border Background="{DynamicResource Surface2}" CornerRadius="12" Padding="16,12" Margin="0,14,0,0">
        <TextBlock x:Name="Hinweise" TextWrapping="Wrap" Foreground="{DynamicResource Muted}" FontSize="13" LineHeight="20"/>
      </Border>
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
        <Button x:Name="Abbrechen" Style="{StaticResource Knopf}" Content="Abbrechen" Margin="0,0,10,0" IsCancel="True"/>
        <Button x:Name="Ausfuehren" Style="{StaticResource PrimaerKnopf}" IsDefault="True">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE7EF;" FontFamily="{StaticResource Icons}" Margin="0,2,9,0"/><TextBlock Text="Jetzt ausführen"/></StackPanel>
        </Button>
      </StackPanel>
    </StackPanel>
  </Grid>
</Window>
'@

$Script:ErgebnisXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Ergebnis" Width="680" SizeToContent="Height" MaxHeight="840" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False" Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="False">
  <Grid Background="{DynamicResource Bg}">
    <StackPanel Margin="30,26,30,24">
      <StackPanel Orientation="Horizontal">
        <Border Width="40" Height="40" CornerRadius="12" Background="{DynamicResource OkBg}" Margin="0,0,14,0">
          <TextBlock Text="&#xE73E;" FontFamily="{StaticResource Icons}" FontSize="18" Foreground="{DynamicResource Ok}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <StackPanel VerticalAlignment="Center">
          <TextBlock x:Name="Titel" FontFamily="{DynamicResource Anzeige}" FontSize="22" FontWeight="SemiBold"/>
          <TextBlock x:Name="Zusammenfassung" Foreground="{DynamicResource Muted}" FontSize="13"/>
        </StackPanel>
      </StackPanel>
      <TextBlock x:Name="Info" Foreground="{DynamicResource Muted}" TextWrapping="Wrap" Margin="0,14,0,12" FontSize="13"/>
      <Border Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="14" Padding="18,8">
        <ScrollViewer MaxHeight="420" VerticalScrollBarVisibility="Auto">
          <ItemsControl x:Name="Liste">
            <ItemsControl.ItemTemplate>
              <DataTemplate>
                <Grid Margin="0,8">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock Text="{Binding Glyph}" Style="{StaticResource ErgebnisGlyph}" Margin="0,2,0,0"/>
                  <StackPanel Grid.Column="1">
                    <TextBlock Text="{Binding Titel}" FontWeight="SemiBold" TextWrapping="Wrap"/>
                    <TextBlock Text="{Binding Meldung}" Foreground="{DynamicResource Muted}" FontSize="12.5" TextWrapping="Wrap" Margin="0,2,0,0"/>
                  </StackPanel>
                </Grid>
              </DataTemplate>
            </ItemsControl.ItemTemplate>
          </ItemsControl>
        </ScrollViewer>
      </Border>
      <DockPanel Margin="0,20,0,0">
        <Button x:Name="Rueckgaengig" DockPanel.Dock="Left" Style="{StaticResource Knopf}" Visibility="Collapsed">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE7A7;" FontFamily="{StaticResource Icons}" Margin="0,2,8,0"/><TextBlock Text="Rückgängig machen"/></StackPanel>
        </Button>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
          <!-- Kein „Erneut prüfen“ mehr: nach dem Schließen prüft bzw. misst die App selbst neu -->
          <Button x:Name="Schliessen" Style="{StaticResource PrimaerKnopf}" Content="Schließen" MinWidth="110" IsDefault="True" IsCancel="True"/>
        </StackPanel>
      </DockPanel>
    </StackPanel>
  </Grid>
</Window>
'@

$Script:HinweisXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Systemdiagnose" Width="520" SizeToContent="Height" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False" Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="False">
  <StackPanel Margin="28,24,28,22">
    <TextBlock x:Name="Titel" FontFamily="{DynamicResource Anzeige}" FontSize="20" FontWeight="SemiBold" TextWrapping="Wrap"/>
    <TextBox x:Name="Text" Style="{StaticResource Markierbar}" Foreground="{DynamicResource Muted}" Margin="0,10,0,0" TextBlock.LineHeight="20"/>
    <Button x:Name="Ok" Style="{StaticResource PrimaerKnopf}" Content="OK" HorizontalAlignment="Right" Margin="0,20,0,0" MinWidth="90" IsDefault="True" IsCancel="True"/>
  </StackPanel>
</Window>
'@

$Script:UpdateXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Update" Width="560" SizeToContent="Height" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False" Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="False">
  <Grid Background="{DynamicResource Bg}">
    <StackPanel Margin="30,26,30,24">
      <StackPanel Orientation="Horizontal">
        <Border Width="40" Height="40" CornerRadius="12" Background="{DynamicResource AccentSoft}" Margin="0,0,14,0">
          <TextBlock Text="&#xE896;" FontFamily="{StaticResource Icons}" FontSize="18" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <TextBlock x:Name="Titel" FontFamily="{DynamicResource Anzeige}" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </StackPanel>
      <TextBlock x:Name="Stand" Foreground="{DynamicResource Muted}" TextWrapping="Wrap" Margin="0,14,0,0"/>
      <TextBlock Text="Neu" Style="{StaticResource AbschnittKopf}" Margin="0,18,0,8"/>
      <Border Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="10" Padding="14,10">
        <ScrollViewer MaxHeight="200" VerticalScrollBarVisibility="Auto">
          <TextBox x:Name="Neu" Style="{StaticResource Markierbar}" TextBlock.LineHeight="20"/>
        </ScrollViewer>
      </Border>
      <TextBlock Foreground="{DynamicResource Muted}" FontSize="12.5" TextWrapping="Wrap" Margin="0,14,0,0"
                 Text="Lädt nur die Programmdateien (signiert und geprüft), Berichte, Archiv und Einstellungen bleiben. Danach startet die App neu."/>
      <DockPanel Margin="0,20,0,0" LastChildFill="False">
        <Button x:Name="Jetzt" DockPanel.Dock="Right" Style="{StaticResource PrimaerKnopf}" Content="Jetzt aktualisieren" IsDefault="True"/>
        <Button x:Name="Spaeter" DockPanel.Dock="Right" Style="{StaticResource Knopf}" Content="Später" Margin="0,0,10,0" IsCancel="True"/>
      </DockPanel>
    </StackPanel>
  </Grid>
</Window>
'@

$Script:FehlerberichtXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Fehler melden" Width="620" SizeToContent="Height" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False" Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="False">
  <Grid Background="{DynamicResource Bg}">
    <StackPanel Margin="30,26,30,24">
      <StackPanel Orientation="Horizontal">
        <Border Width="40" Height="40" CornerRadius="12" Background="{DynamicResource AccentSoft}" Margin="0,0,14,0">
          <TextBlock Text="&#xED15;" FontFamily="{StaticResource Icons}" FontSize="18" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <TextBlock Text="Fehler melden" FontFamily="{DynamicResource Anzeige}" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </StackPanel>
      <TextBlock Foreground="{DynamicResource Muted}" TextWrapping="Wrap" Margin="0,14,0,0" TextBlock.LineHeight="20"
                 Text="Fasst Fehlermeldungen und Fehlercodes dieser App auf diesem PC in einer Datei zusammen - für die Weiterentwicklung. Enthalten sind nur Programmfehler, Messwerte, PC-Name und Windows-Version, keine persönlichen Dateien."/>
      <TextBlock Text="Von der App selbst bemerkt" Style="{StaticResource AbschnittKopf}" Margin="0,20,0,8"/>
      <Border Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="10" Padding="14,10">
        <ScrollViewer MaxHeight="170" VerticalScrollBarVisibility="Auto">
          <TextBox x:Name="Gefunden" Style="{StaticResource Markierbar}" FontSize="13" TextBlock.LineHeight="19"/>
        </ScrollViewer>
      </Border>
      <TextBlock Text="Was ist passiert? (freiwillig)" Style="{StaticResource AbschnittKopf}" Margin="0,20,0,8"/>
      <TextBox x:Name="Beschreibung" Height="96" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
               Background="{DynamicResource Surface}" Foreground="{DynamicResource Text}" CaretBrush="{DynamicResource Text}" FontSize="13.5">
        <TextBox.Template>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="Rahmen" Background="{TemplateBinding Background}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="10" Padding="10,8">
              <Grid>
                <ScrollViewer x:Name="PART_ContentHost"/>
                <TextBlock x:Name="Platzhalter" Text="z. B. „Übermittlungsoptimierung zeigt 6,5 GB, Löschen schlägt immer fehl“" Foreground="{DynamicResource Muted}" Margin="2,0,0,0" IsHitTestVisible="False" Visibility="Collapsed" TextWrapping="Wrap"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="Text" Value=""><Setter TargetName="Platzhalter" Property="Visibility" Value="Visible"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Rahmen" Property="BorderBrush" Value="{DynamicResource Accent}"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </TextBox.Template>
      </TextBox>
      <DockPanel Margin="0,22,0,0" LastChildFill="False">
        <Button x:Name="Erstellen" DockPanel.Dock="Right" Style="{StaticResource PrimaerKnopf}" IsDefault="True">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE74E;" FontFamily="{StaticResource Icons}" Margin="0,2,8,0"/><TextBlock Text="Bericht speichern"/></StackPanel>
        </Button>
        <Button x:Name="Schliessen" DockPanel.Dock="Right" Style="{StaticResource Knopf}" Content="Schließen" Margin="0,0,10,0" IsCancel="True"/>
      </DockPanel>
    </StackPanel>
  </Grid>
</Window>
'@

$Script:EinstellungenXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Einstellungen" Width="640" SizeToContent="Height" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False" Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="False">
  <Grid Background="{DynamicResource Bg}">
    <StackPanel Margin="30,26,30,24">
      <StackPanel Orientation="Horizontal">
        <Border Width="40" Height="40" CornerRadius="12" Background="{DynamicResource AccentSoft}" Margin="0,0,14,0">
          <TextBlock Text="&#xE713;" FontFamily="{StaticResource Icons}" FontSize="18" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <TextBlock Text="Einstellungen" FontFamily="{DynamicResource Anzeige}" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </StackPanel>
      <TextBlock Text="Design" Style="{StaticResource AbschnittKopf}" Margin="0,22,0,8"/>
      <UniformGrid Columns="2">
        <RadioButton x:Name="DesignStandard" Tag="Standard" GroupName="Design" Style="{StaticResource ModusKarte}" Margin="0,0,8,8">
          <StackPanel>
            <StackPanel Orientation="Horizontal">
              <Ellipse Width="14" Height="14" Fill="#F3F4F7" Stroke="#808A99" StrokeThickness="1" Margin="0,2,0,0"/>
              <Ellipse Width="14" Height="14" Fill="#16181D" Stroke="#808A99" StrokeThickness="1" Margin="-4,2,0,0"/>
              <Ellipse Width="14" Height="14" Fill="#4C8DFF" Margin="-4,2,10,0"/>
              <TextBlock Text="Standard" FontWeight="SemiBold" FontSize="15"/>
            </StackPanel>
            <TextBlock Text="Hell oder dunkel - so wie Windows eingestellt ist" Foreground="{DynamicResource Muted}" FontSize="12.5" Margin="0,4,0,0" TextWrapping="Wrap"/>
          </StackPanel>
        </RadioButton>
        <RadioButton x:Name="DesignStarWars" Tag="StarWars" GroupName="Design" Style="{StaticResource ModusKarte}" Margin="0,0,0,8">
          <StackPanel>
            <StackPanel Orientation="Horizontal">
              <Ellipse Width="14" Height="14" Fill="#08090C" Stroke="#808A99" StrokeThickness="1" Margin="0,2,0,0"/>
              <Ellipse Width="14" Height="14" Fill="#5A616A" Margin="-4,2,0,0"/>
              <Ellipse Width="14" Height="14" Fill="#ECEEF1" Margin="-4,2,0,0"/>
              <Ellipse Width="14" Height="14" Fill="#C8102E" Margin="-4,2,10,0"/>
              <TextBlock Text="Star Wars · Imperium" FontWeight="SemiBold" FontSize="15"/>
            </StackPanel>
            <TextBlock Text="Schwarz, Stahlgrau und Rot - dahinter Weltraum mit Gaswolke, Sternzerstörern und einem Sith mit Lichtschwert" Foreground="{DynamicResource Muted}" FontSize="12.5" Margin="0,4,0,0" TextWrapping="Wrap"/>
          </StackPanel>
        </RadioButton>
        <RadioButton x:Name="DesignHell" Tag="Hell" GroupName="Design" Style="{StaticResource ModusKarte}" Margin="0,0,8,8">
          <StackPanel>
            <StackPanel Orientation="Horizontal">
              <Ellipse Width="14" Height="14" Fill="#FFFFFF" Stroke="#808A99" StrokeThickness="1" Margin="0,2,0,0"/>
              <Ellipse Width="14" Height="14" Fill="#2A68CC" Margin="-4,2,10,0"/>
              <TextBlock Text="Immer hell" FontWeight="SemiBold" FontSize="15"/>
            </StackPanel>
            <TextBlock Text="Unabhängig von Windows" Foreground="{DynamicResource Muted}" FontSize="12.5" Margin="0,4,0,0" TextWrapping="Wrap"/>
          </StackPanel>
        </RadioButton>
        <!-- Rechte Spalte = Star-Wars-Designs: Imperium, darunter Clone Wars -->
        <RadioButton x:Name="DesignCloneWars" Tag="CloneWars" GroupName="Design" Style="{StaticResource ModusKarte}" Margin="0,0,0,8">
          <StackPanel>
            <StackPanel Orientation="Horizontal">
              <Ellipse Width="14" Height="14" Fill="#08090C" Stroke="#808A99" StrokeThickness="1" Margin="0,2,0,0"/>
              <Ellipse Width="14" Height="14" Fill="#5A616A" Margin="-4,2,0,0"/>
              <Ellipse Width="14" Height="14" Fill="#ECEEF1" Margin="-4,2,0,0"/>
              <Ellipse Width="14" Height="14" Fill="#1566F1" Margin="-4,2,10,0"/>
              <TextBlock Text="Star Wars · The Clone Wars" FontWeight="SemiBold" FontSize="15"/>
            </StackPanel>
            <TextBlock Text="Schwarz, Stahlgrau und Lichtschwert-Blau - dahinter das eigene Clone-Wars-Bild" Foreground="{DynamicResource Muted}" FontSize="12.5" Margin="0,4,0,0" TextWrapping="Wrap"/>
          </StackPanel>
        </RadioButton>
        <RadioButton x:Name="DesignDunkel" Tag="Dunkel" GroupName="Design" Style="{StaticResource ModusKarte}" Margin="0,0,8,0">
          <StackPanel>
            <StackPanel Orientation="Horizontal">
              <Ellipse Width="14" Height="14" Fill="#1F2228" Stroke="#808A99" StrokeThickness="1" Margin="0,2,0,0"/>
              <Ellipse Width="14" Height="14" Fill="#4C8DFF" Margin="-4,2,10,0"/>
              <TextBlock Text="Immer dunkel" FontWeight="SemiBold" FontSize="15"/>
            </StackPanel>
            <TextBlock Text="Unabhängig von Windows" Foreground="{DynamicResource Muted}" FontSize="12.5" Margin="0,4,0,0" TextWrapping="Wrap"/>
          </StackPanel>
        </RadioButton>
      </UniformGrid>
      <TextBlock Text="Wird sofort übernommen und beim nächsten Start wieder verwendet." Foreground="{DynamicResource Muted}" FontSize="12.5" Margin="2,12,0,0"/>
      <TextBlock Text="Beim Start" Style="{StaticResource AbschnittKopf}" Margin="0,22,0,8"/>
      <CheckBox x:Name="AutoPruefen">
        <TextBlock TextWrapping="Wrap" Text="Automatisch prüfen, wenn der letzte Stand älter als 3 Tage ist (Modus Schnell, ohne Adminrechte, nur lesend)"/>
      </CheckBox>
      <TextBlock Text="Dieser PC" Style="{StaticResource AbschnittKopf}" Margin="0,22,0,6"/>
      <DockPanel>
        <Button x:Name="Einlernen" DockPanel.Dock="Right" Style="{StaticResource Knopf}" Margin="12,0,0,0" VerticalAlignment="Center"
                ToolTip="Den jetzigen Zustand (Monitore, Arbeitsspeicher, Energieplan) als Sollwert übernehmen - wirkt ab der nächsten Prüfung">
          <TextBlock Text="Sollwerte neu einlernen"/>
        </Button>
        <TextBlock x:Name="PcText" Foreground="{DynamicResource Muted}" FontSize="12.5" TextWrapping="Wrap" VerticalAlignment="Center"/>
      </DockPanel>
      <Button x:Name="Schliessen" Style="{StaticResource PrimaerKnopf}" Content="Fertig" HorizontalAlignment="Right" Margin="0,20,0,0" MinWidth="110" IsDefault="True" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
'@

$Script:ProgrammeXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Programme schließen" Width="640" SizeToContent="Height" MaxHeight="820" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False" Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="False">
  <Grid Background="{DynamicResource Bg}">
    <StackPanel Margin="30,26,30,24">
      <StackPanel Orientation="Horizontal">
        <Border Width="40" Height="40" CornerRadius="12" Background="{DynamicResource WarnBg}" Margin="0,0,14,0">
          <TextBlock Text="&#xE7BA;" FontFamily="{StaticResource Icons}" FontSize="18" Foreground="{DynamicResource Warn}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <TextBlock x:Name="Titel" Text="Erst Programme schließen" FontFamily="{DynamicResource Anzeige}" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </StackPanel>
      <TextBlock Text="Um diese Daten zu löschen, müssen die folgenden Programme ganz geschlossen sein. Solange sie laufen, sind ihre Dateien in Benutzung - die Kategorie würde übersprungen."
                 Foreground="{DynamicResource Muted}" TextWrapping="Wrap" Margin="0,14,0,14" LineHeight="20"/>
      <Border Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="14" Padding="18,4">
        <ScrollViewer MaxHeight="420" VerticalScrollBarVisibility="Auto">
          <ItemsControl x:Name="Liste">
            <ItemsControl.ItemTemplate>
              <DataTemplate>
                <Grid Margin="0,10">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock Text="&#xE7BA;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Warn}" Margin="0,3,0,0"/>
                  <StackPanel Grid.Column="1">
                    <TextBlock Text="{Binding Programm}" FontWeight="SemiBold"/>
                    <TextBlock Text="{Binding FuerText}" Foreground="{DynamicResource Muted}" FontSize="12.5" TextWrapping="Wrap" Margin="0,2,0,0"/>
                    <TextBlock Text="{Binding Anleitung}" FontSize="13" TextWrapping="Wrap" Margin="0,7,0,0" LineHeight="19"/>
                  </StackPanel>
                </Grid>
              </DataTemplate>
            </ItemsControl.ItemTemplate>
          </ItemsControl>
        </ScrollViewer>
      </Border>
      <Border Background="{DynamicResource Surface2}" CornerRadius="12" Padding="16,11" Margin="0,14,0,0">
        <StackPanel Orientation="Horizontal">
          <TextBlock Text="&#xE895;" FontFamily="{StaticResource Icons}" Foreground="{DynamicResource Accent}" Margin="0,2,10,0"/>
          <TextBlock x:Name="Status" Text="Sobald alles geschlossen ist, geht es von selbst weiter." Foreground="{DynamicResource Muted}" FontSize="13" TextWrapping="Wrap"/>
        </StackPanel>
      </Border>
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
        <Button x:Name="Abbrechen" Style="{StaticResource Knopf}" Content="Abbrechen" Margin="0,0,10,0" IsCancel="True"/>
        <Button x:Name="Ohne" Style="{StaticResource Knopf}" Content="Ohne diese reinigen" Margin="0,0,10,0"/>
        <Button x:Name="Pruefen" Style="{StaticResource PrimaerKnopf}" IsDefault="True">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE72C;" FontFamily="{StaticResource Icons}" Margin="0,2,9,0"/><TextBlock Text="Erneut prüfen"/></StackPanel>
        </Button>
      </StackPanel>
    </StackPanel>
  </Grid>
</Window>
'@

$Script:VergleichXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Aktualisiert" Width="680" SizeToContent="Height" MaxHeight="840" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False" Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="False">
  <Grid Background="{DynamicResource Bg}">
    <StackPanel Margin="30,26,30,24">
      <StackPanel Orientation="Horizontal">
        <Border Width="40" Height="40" CornerRadius="12" Background="{DynamicResource AccentSoft}" Margin="0,0,14,0">
          <TextBlock Text="&#xE72C;" FontFamily="{StaticResource Icons}" FontSize="18" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <StackPanel VerticalAlignment="Center">
          <TextBlock x:Name="Titel" FontFamily="{DynamicResource Anzeige}" FontSize="22" FontWeight="SemiBold"/>
          <TextBlock x:Name="Zusammenfassung" Foreground="{DynamicResource Muted}" FontSize="13"/>
        </StackPanel>
      </StackPanel>
      <TextBlock x:Name="Info" Foreground="{DynamicResource Muted}" TextWrapping="Wrap" Margin="0,14,0,12" FontSize="13"/>
      <Border x:Name="Rahmen" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Linie}" BorderThickness="1" CornerRadius="14" Padding="18,0,18,8">
        <ScrollViewer MaxHeight="540" VerticalScrollBarVisibility="Auto">
          <ItemsControl x:Name="Liste">
            <ItemsControl.ItemTemplate>
              <DataTemplate>
                <Grid Margin="0,7">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock Text="{Binding Glyph}" Style="{StaticResource ErgebnisGlyph}" Margin="0,2,0,0"/>
                  <StackPanel Grid.Column="1">
                    <TextBlock Text="{Binding Titel}" FontWeight="SemiBold" TextWrapping="Wrap"/>
                    <TextBlock Text="{Binding Meldung}" Foreground="{DynamicResource Muted}" FontSize="12.5" TextWrapping="Wrap" Margin="0,2,0,0"/>
                  </StackPanel>
                </Grid>
              </DataTemplate>
            </ItemsControl.ItemTemplate>
            <ItemsControl.GroupStyle>
              <GroupStyle>
                <GroupStyle.HeaderTemplate>
                  <DataTemplate>
                    <TextBlock Text="{Binding Name}" Foreground="{DynamicResource Muted}" FontSize="12" FontWeight="SemiBold" Margin="0,14,0,2"/>
                  </DataTemplate>
                </GroupStyle.HeaderTemplate>
              </GroupStyle>
            </ItemsControl.GroupStyle>
          </ItemsControl>
        </ScrollViewer>
      </Border>
      <Button x:Name="Schliessen" Style="{StaticResource PrimaerKnopf}" Content="Schließen" HorizontalAlignment="Right" Margin="0,20,0,0" MinWidth="110" IsDefault="True" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
'@

# ============================================================================
#  Hilfsfunktionen
# ============================================================================

function New-XamlObjekt {
    param([string] $Xaml)
    $leser = New-Object System.Xml.XmlNodeReader ([xml]$Xaml)
    return [Windows.Markup.XamlReader]::Load($leser)
}

function Get-Element { param([string] $Name) return $Script:Fenster.FindName($Name) }

function ConvertTo-Bgr {
    param([string] $Hex)
    $c = [System.Windows.Media.ColorConverter]::ConvertFromString($Hex)
    return ([int]$c.B -shl 16) -bor ([int]$c.G -shl 8) -bor [int]$c.R
}

# Design aus den Einstellungen: Standard (wie Windows), Hell, Dunkel, StarWars (Imperium) oder CloneWars
function Set-Design {
    $design = [string]$Script:Einstellungen.Design
    $sw = ($design -in 'StarWars', 'CloneWars')
    $hell = $false
    if ($design -eq 'Hell') { $hell = $true }
    elseif ($design -ne 'Dunkel' -and -not $sw) {
        try { $hell = ((Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name AppsUseLightTheme -ErrorAction Stop).AppsUseLightTheme -eq 1) } catch { }
    }
    $Script:Hell = $hell
    # Spalte = Hintergrund hinter den Listen (nur im Star-Wars-Design abgedunkelt, damit die Grafik den Text nicht stoert)
    $Script:Farben = if ($design -eq 'StarWars') {
        # Imperium: Stahlschwarz, Stahlgrau, Weiss wie die Sturmtruppen-Ruestung, Rot als Akzent; Flaechen durchscheinend
        @{ Bg = '#08090C'; Surface = '#B0111317'; Surface2 = '#A61A1D23'; SurfaceHover = '#D9262A31'; Linie = '#3A3F48'; Text = '#ECEEF1'; Muted = '#959CA6'
           Accent = '#C8102E'; AccentHover = '#E0243F'; AccentText = '#FFFFFF'; AccentSoft = '#47C8102E'; Crit = '#FF5A3C'; CritBg = '#40FF3B1F'
           Warn = '#F5B342'; WarnBg = '#40F5A300'; Info = '#9DB8D8'; InfoBg = '#33758FB0'; Ok = '#7FD99A'; OkBg = '#3342B86A'; Overlay = '#C8040405'; Spalte = '#5C060709' }
    } elseif ($design -eq 'CloneWars') {
        # Clone Wars: wie Imperium, nur Lichtschwert-Blau statt Rot (gemessen an der Klinge in Design\CloneWars.png: #1566F1)
        @{ Bg = '#08090C'; Surface = '#B0111317'; Surface2 = '#A61A1D23'; SurfaceHover = '#D9262A31'; Linie = '#3A3F48'; Text = '#ECEEF1'; Muted = '#959CA6'
           Accent = '#1566F1'; AccentHover = '#3A82FF'; AccentText = '#FFFFFF'; AccentSoft = '#471566F1'; Crit = '#FF5A3C'; CritBg = '#40FF3B1F'
           Warn = '#F5B342'; WarnBg = '#40F5A300'; Info = '#9DB8D8'; InfoBg = '#33758FB0'; Ok = '#7FD99A'; OkBg = '#3342B86A'; Overlay = '#C8040405'; Spalte = '#5C060709' }
    } elseif ($hell) {
        @{ Bg = '#F3F4F7'; Surface = '#FFFFFF'; Surface2 = '#F3F5F8'; SurfaceHover = '#EAEDF2'; Linie = '#DFE3E9'; Text = '#1C2028'; Muted = '#5D6876'
           Accent = '#2A68CC'; AccentHover = '#3B78DA'; AccentText = '#FFFFFF'; AccentSoft = '#E5EEFC'; Crit = '#C52A2A'; CritBg = '#FDE7E7'
           Warn = '#B06A00'; WarnBg = '#FFF2D9'; Info = '#2A68CC'; InfoBg = '#E5EEFC'; Ok = '#1D8A4B'; OkBg = '#E3F4E9'; Overlay = '#A6F3F4F7'; Spalte = '#00000000' }
    } else {
        @{ Bg = '#16181D'; Surface = '#1F2228'; Surface2 = '#262A31'; SurfaceHover = '#2D323A'; Linie = '#323741'; Text = '#E8EAED'; Muted = '#9AA3AE'
           Accent = '#4C8DFF'; AccentHover = '#6AA1FF'; AccentText = '#FFFFFF'; AccentSoft = '#1B2940'; Crit = '#FF6B6B'; CritBg = '#3B1F22'
           Warn = '#F2B33D'; WarnBg = '#3A2E16'; Info = '#6AA1FF'; InfoBg = '#1C2A42'; Ok = '#4CC27C'; OkBg = '#183022'; Overlay = '#B3101216'; Spalte = '#00000000' }
    }
    foreach ($k in @($Script:Farben.Keys)) {
        # ::new statt New-Object: New-Object liefert eine PowerShell-Huelle, die WPF nicht als Brush erkennt
        $farbe = [System.Windows.Media.Color][System.Windows.Media.ColorConverter]::ConvertFromString($Script:Farben[$k])
        $pinsel = [System.Windows.Media.SolidColorBrush]::new($farbe)
        $pinsel.Freeze()
        $Script:App.Resources[$k] = $pinsel
    }
    if ($sw) { $bild = Get-ImperiumHintergrund $design; if ($bild) { $Script:App.Resources['Bg'] = $bild } }
    $schrift = if ($sw) { 'Bahnschrift, Segoe UI' } else { 'Segoe UI Variable Display, Segoe UI' }
    $Script:App.Resources['Anzeige'] = [System.Windows.Media.FontFamily]::new($schrift)
    # Reiter: in den Star-Wars-Designs Lichtschwert-Rahmen statt Unterstrich (siehe TabItem-Vorlage), Klinge rot oder blau
    $klinge = if ($design -eq 'CloneWars') { @{ KlingeSchein = '#FF1566F1'; KlingeKern = '#FF4A8CFF'; KlingeKernHell = '#FFE4F2FF'; KlingeGlut = '#FF1A6CFF' } }
              else { @{ KlingeSchein = '#FFFF1A2E'; KlingeKern = '#FFFF4A5A'; KlingeKernHell = '#FFFFF0F2'; KlingeGlut = '#FFFF1E32' } }
    foreach ($k in @($klinge.Keys)) {
        $farbe = [System.Windows.Media.Color][System.Windows.Media.ColorConverter]::ConvertFromString($klinge[$k])
        if ($k -eq 'KlingeGlut') { $Script:App.Resources[$k] = $farbe; continue }   # DropShadowEffect will eine Farbe, keinen Pinsel
        $pinsel = [System.Windows.Media.SolidColorBrush]::new($farbe); $pinsel.Freeze()
        $Script:App.Resources[$k] = $pinsel
    }
    $Script:App.Resources['ReiterKlingeSichtbar'] = [double]$(if ($sw) { 1 } else { 0 })
    $Script:App.Resources['ReiterStrichSichtbar'] = [double]$(if ($sw) { 0 } else { 1 })
    # Reiter = oberste Ebene: groesser als die Filter-Chips darunter (Johannes, 28.10.)
    $Script:App.Resources['ReiterAbstand']  = if ($sw) { [System.Windows.Thickness]::new(0, 0, 14, 0) } else { [System.Windows.Thickness]::new(0, 0, 30, 0) }
    # Innenabstand der abgedunkelten, abgerundeten Listen-Flaeche (nur Star Wars; sonst ist sie unsichtbar)
    $Script:App.Resources['SpaltenRand'] = if ($sw) { [System.Windows.Thickness]::new(10, 10, 4, 2) } else { [System.Windows.Thickness]::new(0) }
    $Script:App.Resources['ReiterKopfRand'] = if ($sw) { [System.Windows.Thickness]::new(24, 15, 24, 15) } else { [System.Windows.Thickness]::new(4, 13, 4, 15) }
}

# Hintergrund fuer das Star-Wars-Design (Imperium): wird als Bild berechnet statt aus flachen Vektorformen -
# Gaswolke aus Rauschen, Sterne mit Gluehen, Planet mit Atmosphaere und Stadtlichtern, Sternzerstoerer mit
# beleuchteten Rumpfflaechen, im Vordergrund eine Sith-Gestalt im Gegenlicht mit rotem Lichtschwert und Glut.
# Erst beim ersten Gebrauch uebersetzt (C#, etwa 0,2 s) und berechnet (etwa 1,3 s), danach im Speicher.
$Script:WeltraumCode = @'
using System;

// Hintergrundbild fuer das Star-Wars-Design: alles wird pro Pixel berechnet (Rauschen, Gluehen, Gegenlicht),
// damit es wie ein Filmbild wirkt und nicht wie flache Vektorformen. Ergebnis: deckende BGRA-Pixel.
public static class SdWeltraum
{
    static int[] perm;
    static float[] R, G, B, staub;
    static int W, H, w2, h2;
    static double K;   // Massstab: 1 = 1440 Pixel Bildhoehe

    public static byte[] Zeichne(int breite, int hoehe)
    {
        W = breite; H = hoehe; K = hoehe / 1440.0;
        Mische(1977);
        R = new float[W * H]; G = new float[W * H]; B = new float[W * H];
        Grund();
        Nebel();
        Sterne();
        Planet();
        Flotte();
        Vordergrund();
        return Ausgabe();
    }

    // ------------------------------------------------------------------ Grundlagen
    static void Mische(int seed)
    {
        var r = new Random(seed);
        var p = new int[256];
        for (int i = 0; i < 256; i++) p[i] = i;
        for (int i = 255; i > 0; i--) { int j = r.Next(i + 1); int t = p[i]; p[i] = p[j]; p[j] = t; }
        perm = new int[512];
        for (int i = 0; i < 512; i++) perm[i] = p[i & 255];
    }

    static double Lerp(double a, double b, double t) { return a + (b - a) * t; }
    static double Weich(double a, double b, double x) { double t = Math.Max(0, Math.Min(1, (x - a) / (b - a))); return t * t * (3 - 2 * t); }
    static double Glatt(double t) { return t * t * t * (t * (t * 6 - 15) + 10); }

    static double Grad(int h, double x, double y)
    {
        switch (h & 7)
        {
            case 0: return x + y;
            case 1: return -x + y;
            case 2: return x - y;
            case 3: return -x - y;
            case 4: return x;
            case 5: return -x;
            case 6: return y;
            default: return -y;
        }
    }

    // Perlin-Rauschen, etwa -1..1
    static double Rauschen(double x, double y)
    {
        int xi = (int)Math.Floor(x), yi = (int)Math.Floor(y);
        double xf = x - xi, yf = y - yi;
        xi &= 255; yi &= 255;
        double u = Glatt(xf), v = Glatt(yf);
        int aa = perm[perm[xi] + yi], ab = perm[perm[xi] + yi + 1], ba = perm[perm[xi + 1] + yi], bb = perm[perm[xi + 1] + yi + 1];
        return Lerp(Lerp(Grad(aa, xf, yf), Grad(ba, xf - 1, yf), u), Lerp(Grad(ab, xf, yf - 1), Grad(bb, xf - 1, yf - 1), u), v);
    }

    static double Fbm(double x, double y, int oktaven)
    {
        double s = 0, a = 0.5, f = 1;
        for (int i = 0; i < oktaven; i++) { s += a * Rauschen(x * f, y * f); f *= 2.03; a *= 0.5; }
        return s;
    }

    static double Hash(int x, int y)
    {
        unchecked
        {
            uint h = (uint)(x * 374761393 + y * 668265263);
            h = (h ^ (h >> 13)) * 1274126177u;
            h ^= h >> 16;
            return (h & 0xFFFFFF) / 16777216.0;
        }
    }

    static void Addiere(int x, int y, double r, double g, double b)
    {
        if (x < 0 || y < 0 || x >= W || y >= H) return;
        int i = y * W + x;
        R[i] += (float)r; G[i] += (float)g; B[i] += (float)b;
    }

    static void Mische(int i, double r, double g, double b, double deck)
    {
        R[i] = (float)(R[i] * (1 - deck) + r * deck);
        G[i] = (float)(G[i] * (1 - deck) + g * deck);
        B[i] = (float)(B[i] * (1 - deck) + b * deck);
    }

    static double Abstand(double x, double y, double ax, double ay, double bx, double by)
    {
        double dx = bx - ax, dy = by - ay, l = dx * dx + dy * dy;
        double t = l > 0 ? ((x - ax) * dx + (y - ay) * dy) / l : 0;
        t = Math.Max(0, Math.Min(1, t));
        double qx = ax + t * dx - x, qy = ay + t * dy - y;
        return Math.Sqrt(qx * qx + qy * qy);
    }

    static bool Innen(double[] px, double[] py, double x, double y)
    {
        bool c = false;
        int n = px.Length;
        for (int i = 0, j = n - 1; i < n; j = i++)
            if (((py[i] > y) != (py[j] > y)) && (x < (px[j] - px[i]) * (y - py[i]) / (py[j] - py[i]) + px[i])) c = !c;
        return c;
    }

    // Kantenglaettung: Anteil eines Pixels innerhalb eines Polygons (4x4 Stichproben)
    static double Abdeckung(double[] px, double[] py, int x, int y)
    {
        int t = 0;
        for (int sy = 0; sy < 4; sy++)
            for (int sx = 0; sx < 4; sx++)
                if (Innen(px, py, x + (sx + 0.5) / 4, y + (sy + 0.5) / 4)) t++;
        return t / 16.0;
    }

    // ------------------------------------------------------------------ Weltraum
    static void Grund()
    {
        for (int y = 0; y < H; y++)
        {
            double v = (double)y / H;
            float r = (float)(0.008 + 0.004 * (1 - v)), g = (float)(0.010 + 0.005 * (1 - v)), b = (float)(0.015 + 0.008 * (1 - v));
            for (int x = 0; x < W; x++) { int i = y * W + x; R[i] = r; G[i] = g; B[i] = b; }
        }
    }

    // Rote Gaswolke: unregelmaessiges Band von links oben zur Bildmitte rechts
    static double MaskeRot(double u, double v)
    {
        double a = (double)W / H;
        double x0 = 0.00 * a, y0 = 0.06, x1 = 0.92 * a, y1 = 0.50;
        double px = u * a, py = v, dx = x1 - x0, dy = y1 - y0;
        double t = ((px - x0) * dx + (py - y0) * dy) / (dx * dx + dy * dy);
        double tc = Math.Max(0, Math.Min(1, t));
        double qx = x0 + tc * dx - px, qy = y0 + tc * dy - py;
        double d = Math.Sqrt(qx * qx + qy * qy);
        double breite = 0.15 + 0.09 * Math.Sin(t * 3.4 + 0.6);
        return Math.Exp(-(d * d) / (breite * breite)) * (1 - Weich(0.72, 1.08, t)) * Weich(-0.15, 0.10, t);
    }

    // Kuehler blauer Schleier rechts oben
    static double MaskeBlau(double u, double v)
    {
        double dx = (u - 0.84) / 0.34, dy = (v - 0.16) / 0.28;
        return Math.Exp(-(dx * dx + dy * dy));
    }

    static void Nebel()
    {
        w2 = W / 2; h2 = H / 2;
        float[] nr = new float[w2 * h2], ng = new float[w2 * h2], nb = new float[w2 * h2];
        staub = new float[w2 * h2];
        double s = 2.3 / h2;
        for (int y = 0; y < h2; y++)
            for (int x = 0; x < w2; x++)
            {
                double px = x * s, py = y * s, u = (double)x / w2, v = (double)y / h2;
                double qx = Fbm(px + 1.7, py + 9.2, 4), qy = Fbm(px + 8.3, py + 2.8, 4);
                double n = Fbm(px + 2.6 * qx, py + 2.6 * qy, 6);
                double n2 = Fbm(px * 1.4 + 21.1 + 1.8 * qy, py * 1.4 + 4.4 + 1.8 * qx, 5);
                double d = Fbm(px * 2.2 + 40.3 + 1.2 * qx, py * 2.2 + 13.7 + 1.2 * qy, 5);
                double mr = MaskeRot(u, v), mb = MaskeBlau(u, v);
                double rot = Weich(-0.12, 0.55, n) * mr;
                double faden = Math.Pow(Weich(0.18, 0.72, n), 2.4) * mr;
                double blau = Weich(-0.10, 0.62, n2) * (mb + 0.22 * mr);
                double dunst = 0.5 * Weich(-0.6, 0.6, n) + 0.5 * Weich(-0.5, 0.7, n2);
                double st = Weich(0.02, 0.42, d) * (0.30 + 0.70 * Math.Max(mr, mb));
                double dimm = 1 - 0.80 * st;
                int i = y * w2 + x;
                nr[i] = (float)((0.40 * rot + 0.52 * faden + 0.035 * blau + 0.008 * dunst) * dimm);
                ng[i] = (float)((0.026 * rot + 0.16 * faden + 0.105 * blau + 0.010 * dunst) * dimm);
                nb[i] = (float)((0.060 * rot + 0.12 * faden + 0.210 * blau + 0.017 * dunst) * dimm);
                staub[i] = (float)st;
            }
        for (int y = 0; y < H; y++)
        {
            double fy = Math.Min(h2 - 1.001, Math.Max(0, (y + 0.5) / 2.0 - 0.5));
            int y0 = (int)fy; double ty = fy - y0;
            for (int x = 0; x < W; x++)
            {
                double fx = Math.Min(w2 - 1.001, Math.Max(0, (x + 0.5) / 2.0 - 0.5));
                int x0 = (int)fx; double tx = fx - x0;
                int a = y0 * w2 + x0, b = a + 1, c = a + w2, e = c + 1;
                double wa = (1 - tx) * (1 - ty), wb = tx * (1 - ty), wc = (1 - tx) * ty, we = tx * ty;
                int i = y * W + x;
                R[i] += (float)(nr[a] * wa + nr[b] * wb + nr[c] * wc + nr[e] * we);
                G[i] += (float)(ng[a] * wa + ng[b] * wb + ng[c] * wc + ng[e] * we);
                B[i] += (float)(nb[a] * wa + nb[b] * wb + nb[c] * wc + nb[e] * we);
            }
        }
    }

    static double StaubBei(double x, double y)
    {
        int sx = Math.Max(0, Math.Min(w2 - 1, (int)(x / 2))), sy = Math.Max(0, Math.Min(h2 - 1, (int)(y / 2)));
        return staub[sy * w2 + sx];
    }

    static void Sterne()
    {
        var r = new Random(77);
        int anzahl = (int)(W * H / 480.0);
        for (int n = 0; n < anzahl; n++)
        {
            double x = r.NextDouble() * W, y = r.NextDouble() * H;
            double dichte = 0.40 + 0.95 * MaskeRot(x / W, y / H) + 0.30 * MaskeBlau(x / W, y / H);
            double zufall = r.NextDouble(), hellZufall = r.NextDouble(), farbe = r.NextDouble();
            if (zufall > dichte / 1.65) continue;
            double hell = (0.04 + 0.40 * Math.Pow(hellZufall, 6)) * (1 - 0.72 * StaubBei(x, y));
            Stern(x, y, hell, farbe);
        }
        // nur wenige helle Sterne mit Lichthof
        for (int n = 0; n < 16; n++)
        {
            double x = r.NextDouble() * W, y = r.NextDouble() * H * 0.85, hell = 0.55 + 0.40 * Math.Pow(r.NextDouble(), 2.5);
            Stern(x, y, hell * (1 - 0.6 * StaubBei(x, y)), r.NextDouble());
        }
    }

    static void Stern(double x, double y, double hell, double farbe)
    {
        if (farbe < 0.18) Stern(x, y, hell, 0.72, 0.84, 1.0);
        else if (farbe < 0.70) Stern(x, y, hell, 1.0, 0.98, 0.96);
        else if (farbe < 0.90) Stern(x, y, hell, 1.0, 0.90, 0.78);
        else Stern(x, y, hell, 1.0, 0.76, 0.58);
    }

    static void Stern(double x, double y, double hell, double cr, double cg, double cb)
    {
        double sigma = (0.62 + 1.15 * hell) * K;
        int rad = (int)Math.Ceiling(sigma * 3.2) + 1;
        bool halo = hell > 0.5;
        if (halo) rad = Math.Max(rad, (int)(40 * K));
        int xi = (int)x, yi = (int)y;
        for (int yy = yi - rad; yy <= yi + rad; yy++)
        {
            if (yy < 0 || yy >= H) continue;
            for (int xx = xi - rad; xx <= xi + rad; xx++)
            {
                if (xx < 0 || xx >= W) continue;
                double dx = xx + 0.5 - x, dy = yy + 0.5 - y, d2 = dx * dx + dy * dy;
                double I = 2.0 * hell * Math.Exp(-d2 / (2 * sigma * sigma));
                if (halo)
                {
                    double d = Math.Sqrt(d2);
                    I += 0.045 * hell * Math.Exp(-d / ((4 + 9 * hell) * K));
                    if (hell > 0.82)
                    {
                        // zarte Lichtstrahlen wie durch ein Kameraobjektiv
                        double sp = Math.Exp(-Math.Abs(dy) / (0.55 * K)) * Math.Exp(-Math.Abs(dx) / (11 * hell * K))
                                  + Math.Exp(-Math.Abs(dx) / (0.55 * K)) * Math.Exp(-Math.Abs(dy) / (11 * hell * K));
                        I += 0.11 * hell * sp;
                    }
                }
                if (I < 0.0008) continue;
                int i = yy * W + xx;
                R[i] += (float)(I * cr); G[i] += (float)(I * cg); B[i] += (float)(I * cb);
            }
        }
    }

    // ------------------------------------------------------------------ Planet unten links
    static double pcx, pcy, prad;
    static readonly double atR = 0.20, atG = 0.44, atB = 0.80;

    static void Planet()
    {
        pcx = 0.13 * W; pcy = 1.62 * H; prad = 0.80 * H;
        double lx = 0.62, ly = -0.60, lz = 0.50;
        double ll = Math.Sqrt(lx * lx + ly * ly + lz * lz); lx /= ll; ly /= ll; lz /= ll;
        double l2 = Math.Sqrt(lx * lx + ly * ly), l2x = lx / l2, l2y = ly / l2;
        int yStart = (int)Math.Max(0, pcy - prad - 260 * K);
        for (int y = yStart; y < H; y++)
            for (int x = 0; x < W; x++)
            {
                double dx = x + 0.5 - pcx, dy = y + 0.5 - pcy, d = Math.Sqrt(dx * dx + dy * dy);
                int i = y * W + x;
                double richt = Math.Max(0, (dx * l2x + dy * l2y) / d);
                if (d < prad + 1.0)
                {
                    double nx = dx / prad, ny = dy / prad, nz2 = 1 - nx * nx - ny * ny;
                    double nz = nz2 > 0 ? Math.Sqrt(nz2) : 0;
                    double lam = Math.Max(0, nx * lx + ny * ly + nz * lz);
                    double band = 0.72 + 0.28 * Fbm(nx * 2.4 + 30, ny * 8.5 + nz * 3, 4);
                    double pr = 0.005 + lam * 0.12 * band, pg = 0.006 + lam * 0.14 * band, pb = 0.009 + lam * 0.18 * band;
                    double saum = Math.Pow(1 - nz, 5) * (0.10 + 0.90 * Math.Pow(richt, 1.4));
                    pr += atR * saum; pg += atG * saum; pb += atB * saum;
                    if (lam < 0.04)
                    {
                        // Stadtlichter auf der Nachtseite
                        double stadt = Fbm(nx * 38 + 7, ny * 38 + 3, 3);
                        if (stadt > 0.18 && Hash(x, y) > 0.975)
                        {
                            double c = Math.Min(1, (stadt - 0.18) * 3) * (1 - lam / 0.04) * nz;
                            pr += 0.32 * c; pg += 0.15 * c; pb += 0.04 * c;
                        }
                    }
                    double cov = Math.Max(0, Math.Min(1, prad + 0.5 - d));
                    Mische(i, pr, pg, pb, cov);
                    if (cov >= 1) continue;
                }
                double h = Math.Max(0, d - prad) / K;
                double glanz = (Math.Exp(-h / 8) * 0.50 + Math.Exp(-h / 38) * 0.11 + Math.Exp(-h / 150) * 0.03) * (0.06 + 0.94 * Math.Pow(richt, 1.3));
                if (glanz < 0.0004) continue;
                R[i] += (float)(atR * glanz); G[i] += (float)(atG * glanz); B[i] += (float)(atB * glanz);
            }
    }

    static double PlanetRandY(double x)
    {
        double dx = x - pcx;
        if (Math.Abs(dx) >= prad) return H * 2;
        return pcy - Math.Sqrt(prad * prad - dx * dx);
    }

    // ------------------------------------------------------------------ Sternzerstoerer, schraeg von oben, Bug links
    // Flaechen in Schiffslaengen (Bug bei 0,0; y nach unten); Licht kommt von links oben aus der Gaswolke
    static readonly double[] RueckenFern = { 0.000, 0.000, 1.000, -0.150, 1.000, -0.082 };
    static readonly double[] RueckenNah = { 0.000, 0.000, 1.000, -0.082, 1.000, -0.022 };
    static readonly double[] Seite = { 0.000, 0.000, 1.000, -0.022, 1.000, 0.052, 0.640, 0.036 };
    static readonly double[] Stufe1 = { 0.600, -0.072, 0.650, -0.102, 0.975, -0.122, 0.975, -0.070 };
    static readonly double[] Stufe1Vorne = { 0.600, -0.072, 0.975, -0.070, 0.975, -0.058, 0.610, -0.062 };
    static readonly double[] Stufe2 = { 0.735, -0.104, 0.765, -0.126, 0.955, -0.136, 0.955, -0.110 };
    static readonly double[] Turm = { 0.858, -0.178, 0.884, -0.178, 0.888, -0.130, 0.854, -0.130 };
    static readonly double[] Bruecke = { 0.822, -0.200, 0.948, -0.200, 0.950, -0.180, 0.820, -0.180 };

    static void Flotte()
    {
        Schiff(0.545 * W, 0.425 * H, 0.30 * H, 1.00);
        Schiff(0.735 * W, 0.300 * H, 0.17 * H, 0.90);
        Schiff(0.475 * W, 0.275 * H, 0.105 * H, 0.78);
    }

    static void Schiff(double x0, double y0, double L, double deck)
    {
        Flaeche(RueckenFern, x0, y0, L, 0.085, 0.092, 0.106, deck, 1);
        Flaeche(RueckenNah, x0, y0, L, 0.150, 0.158, 0.175, deck, 2);
        Flaeche(Seite, x0, y0, L, 0.026, 0.029, 0.036, deck, 3);
        Flaeche(Stufe1, x0, y0, L, 0.165, 0.172, 0.190, deck, 4);
        Flaeche(Stufe1Vorne, x0, y0, L, 0.040, 0.044, 0.052, deck, 5);
        Flaeche(Stufe2, x0, y0, L, 0.185, 0.192, 0.210, deck, 6);
        Flaeche(Turm, x0, y0, L, 0.060, 0.064, 0.074, deck, 7);
        Flaeche(Bruecke, x0, y0, L, 0.120, 0.126, 0.140, deck, 8);
        Kuppel(x0 + 0.838 * L, y0 - 0.206 * L, 0.011 * L, deck);
        Kuppel(x0 + 0.932 * L, y0 - 0.206 * L, 0.011 * L, deck);
        // Lichtkanten: Grat und obere Kanten hell, Unterkante mit blauem Widerschein vom Planeten
        Linie(x0, y0, x0 + L, y0 - 0.082 * L, 0.55 * K, 0.30 * deck, 0.32 * deck, 0.37 * deck);
        Linie(x0, y0, x0 + L, y0 - 0.150 * L, 0.45 * K, 0.16 * deck, 0.17 * deck, 0.20 * deck);
        Linie(x0 + 0.650 * L, y0 - 0.102 * L, x0 + 0.975 * L, y0 - 0.122 * L, 0.45 * K, 0.24 * deck, 0.25 * deck, 0.29 * deck);
        Linie(x0 + 0.822 * L, y0 - 0.200 * L, x0 + 0.948 * L, y0 - 0.200 * L, 0.45 * K, 0.26 * deck, 0.27 * deck, 0.31 * deck);
        Linie(x0, y0, x0 + 0.640 * L, y0 + 0.036 * L, 0.5 * K, 0.04 * deck, 0.09 * deck, 0.17 * deck);
        Linie(x0 + 0.640 * L, y0 + 0.036 * L, x0 + L, y0 + 0.052 * L, 0.5 * K, 0.04 * deck, 0.09 * deck, 0.17 * deck);
        // Positionslichter an der Seite
        var r = new Random((int)(x0 + y0));
        int lichter = (int)(L / (9 * K));
        for (int k = 0; k < lichter; k++)
        {
            double t = 0.25 + 0.72 * r.NextDouble();
            double lx = x0 + t * L, ly = y0 + Lerp(-0.010, 0.030, r.NextDouble()) * L * t;
            double I = (0.25 + 0.5 * r.NextDouble()) * deck;
            Addiere((int)lx, (int)ly, 0.75 * I, 0.85 * I, 1.0 * I);
        }
        // Triebwerke am Heck
        foreach (double ty in new[] { -0.070, -0.030, 0.012 })
        {
            double ex = x0 + 1.006 * L, ey = y0 + ty * L, sg = 0.7 * K + 0.0035 * L;
            int rad = (int)(L * 0.07) + 3;
            for (int y = (int)ey - rad; y <= (int)ey + rad; y++)
                for (int x = (int)ex - rad; x <= (int)ex + rad; x++)
                {
                    double ddx = x + 0.5 - ex, ddy = y + 0.5 - ey, d2 = ddx * ddx + ddy * ddy, d = Math.Sqrt(d2);
                    double I = (1.0 * Math.Exp(-d2 / (2 * sg * sg)) + 0.20 * Math.Exp(-d / (0.02 * L))) * deck;
                    Addiere(x, y, 0.62 * I, 0.80 * I, 1.0 * I);
                }
        }
    }

    // Eine Rumpfflaeche mit feiner Plattenstruktur (blockiges Rauschen statt glatter Farbe)
    static void Flaeche(double[] umriss, double x0, double y0, double L, double r, double g, double b, double deck, int saat)
    {
        int n = umriss.Length / 2;
        double[] px = new double[n], py = new double[n];
        double minx = 1e9, maxx = -1e9, miny = 1e9, maxy = -1e9;
        for (int k = 0; k < n; k++)
        {
            px[k] = x0 + umriss[2 * k] * L; py[k] = y0 + umriss[2 * k + 1] * L;
            minx = Math.Min(minx, px[k]); maxx = Math.Max(maxx, px[k]); miny = Math.Min(miny, py[k]); maxy = Math.Max(maxy, py[k]);
        }
        double plX = Math.Max(2, 0.012 * L), plY = Math.Max(1.5, 0.006 * L);
        for (int y = (int)miny - 1; y <= (int)maxy + 1; y++)
            for (int x = (int)minx - 1; x <= (int)maxx + 1; x++)
            {
                if (x < 0 || y < 0 || x >= W || y >= H) continue;
                double cov = Abdeckung(px, py, x, y);
                if (cov <= 0) continue;
                double platte = 0.80 + 0.40 * Hash((int)((x - x0) / plX) + saat * 977, (int)((y - y0) / plY) + saat * 131);
                double vorn = 0.75 + 0.25 * Math.Min(1, (x - x0) / L * 1.6);   // zum Bug hin etwas dunkler
                Mische(y * W + x, r * platte * vorn, g * platte * vorn, b * platte * vorn, cov * deck);
            }
    }

    static void Kuppel(double cx, double cy, double rad, double deck)
    {
        for (int y = (int)(cy - rad) - 1; y <= (int)(cy + rad) + 1; y++)
            for (int x = (int)(cx - rad) - 1; x <= (int)(cx + rad) + 1; x++)
            {
                if (x < 0 || y < 0 || x >= W || y >= H) continue;
                double d = Math.Sqrt(Math.Pow(x + 0.5 - cx, 2) + Math.Pow(y + 0.5 - cy, 2));
                double cov = Math.Max(0, Math.Min(1, rad + 0.5 - d));
                if (cov <= 0) continue;
                double l = 0.08 + 0.14 * Math.Max(0, (cx - x) / rad * 0.6 + (cy - y) / rad * 0.6);
                Mische(y * W + x, l, l * 1.05, l * 1.18, cov * deck);
            }
    }

    // Weiche Linie (Kantenlicht)
    static void Linie(double ax, double ay, double bx, double by, double dicke, double r, double g, double b)
    {
        int x0 = (int)(Math.Min(ax, bx) - 3), x1 = (int)(Math.Max(ax, bx) + 3), y0 = (int)(Math.Min(ay, by) - 3), y1 = (int)(Math.Max(ay, by) + 3);
        for (int y = y0; y <= y1; y++)
            for (int x = x0; x <= x1; x++)
            {
                double d = Abstand(x + 0.5, y + 0.5, ax, ay, bx, by);
                double c = Math.Max(0, 1 - d / (dicke + 0.6));
                if (c > 0) Addiere(x, y, r * c, g * c, b * c);
            }
    }

    // ------------------------------------------------------------------ Vordergrund: Felsvorsprung, Sith, Lichtschwert, Glut
    // Gestalt von hinten (Kapuze, Umhang), Masse in einer 0..120 x 0..260-Box, Fuesse unten
    static readonly double[] Gestalt = {
        52, 0,   60, 4,   66, 14,  68, 28,  70, 40,  80, 46,  88, 54,  92, 64,  95, 82,  100, 104,
        107, 127, 114, 145, 109, 152, 101, 143, 95, 124, 91, 140, 94, 176, 98, 212, 102, 241, 106, 257,
        97, 259,  90, 255, 82, 259, 72, 256, 62, 260, 50, 257, 40, 260, 30, 256, 20, 259, 10, 256,
        1, 258,   5, 234,  9, 198, 12, 160, 14, 120, 13, 90,  14, 66,  20, 54,  32, 45,  36, 38,
        34, 26,   37, 12,  43, 4 };

    static void Vordergrund()
    {
        // Felsvorsprung unten links, Oberkante als Hoehenlinie
        int xEnde = (int)(0.47 * W);
        double[] oben = new double[xEnde + 1];
        for (int x = 0; x <= xEnde; x++)
        {
            double t = (double)x / xEnde;
            double y = Lerp(0.930 * H, 0.968 * H, t) + Fbm(x * 0.006 / K, 3.3, 4) * 0.018 * H + Fbm(x * 0.045 / K, 7.7, 3) * 0.004 * H;
            if (t > 0.88) y += Math.Pow((t - 0.88) / 0.12, 2) * (H - y + 4);
            oben[x] = y;
        }
        double sx = 0.285 * W, sy = oben[(int)sx] + 2 * K;   // Fuesse der Gestalt
        double s = 0.33 * H / 260.0;
        double hx = sx + (118 - 50) * s, hy = sy - (260 - 164) * s;   // Klingenansatz
        double tx = sx + (190 - 50) * s, ty = sy - (260 - 259) * s;   // Klingenspitze am Boden
        double handX = sx + (111 - 50) * s, handY = sy - (260 - 150) * s;

        for (int x = 0; x <= xEnde; x++)
        {
            double yTop = oben[x];
            double rand = PlanetRandY(x);
            double gegen = 0.25 + 0.75 * Math.Exp(-Math.Max(0, yTop - rand) / (120 * K)) * Weich(0.0, 0.25, (double)x / W + 0.05);
            for (int y = (int)yTop; y < H; y++)
            {
                double cov = Math.Max(0, Math.Min(1, y + 1 - yTop));
                double tief = (y - yTop) / K;
                double r = 0.006, g = 0.007, b = 0.009;
                double kante = Math.Exp(-tief / 1.6) * gegen;
                r += 0.10 * kante; g += 0.20 * kante; b += 0.36 * kante;
                double dS = Abstand(x + 0.5, y + 0.5, hx, hy, tx, ty) / K;
                double rotLicht = 0.55 / (1 + Math.Pow(dS / 70, 2)) * Math.Exp(-tief / 45);
                r += 0.95 * rotLicht; g += 0.06 * rotLicht; b += 0.07 * rotLicht;
                Mische(y * W + x, r, g, b, cov);
            }
        }

        // Gestalt: Abdeckung in eine lokale Maske, dann weichgezeichnet fuer das Randlicht
        int n = Gestalt.Length / 2;
        double[] px = new double[n], py = new double[n];
        double minx = 1e9, maxx = -1e9, miny = 1e9, maxy = -1e9;
        for (int k = 0; k < n; k++)
        {
            px[k] = sx + (Gestalt[2 * k] - 50) * s; py[k] = sy - (260 - Gestalt[2 * k + 1]) * s;
            minx = Math.Min(minx, px[k]); maxx = Math.Max(maxx, px[k]); miny = Math.Min(miny, py[k]); maxy = Math.Max(maxy, py[k]);
        }
        int bx0 = (int)minx - 8, by0 = (int)miny - 8, bw = (int)(maxx - minx) + 17, bh = (int)(maxy - miny) + 17;
        double[] maske = new double[bw * bh];
        for (int y = 0; y < bh; y++)
            for (int x = 0; x < bw; x++)
                maske[y * bw + x] = Abdeckung(px, py, bx0 + x, by0 + y);
        double[] weich = Weichzeichnen(Weichzeichnen(maske, bw, bh, (int)Math.Max(2, 3 * K)), bw, bh, (int)Math.Max(2, 3 * K));
        for (int y = 0; y < bh; y++)
            for (int x = 0; x < bw; x++)
            {
                double cov = maske[y * bw + x];
                if (cov <= 0) continue;
                int gx = bx0 + x, gy = by0 + y;
                if (gx < 0 || gy < 0 || gx >= W || gy >= H) continue;
                double rand = Math.Max(0, Math.Min(1, (1 - weich[y * bw + x]) * 2.4)) * cov;
                // Normale aus dem weichgezeichneten Umriss (zeigt nach aussen) - Randlicht nur dort, wo Licht herkommt
                double nx = (x > 0 && x < bw - 1) ? weich[y * bw + x - 1] - weich[y * bw + x + 1] : 0;
                double ny = (y > 0 && y < bh - 1) ? weich[(y - 1) * bw + x] - weich[(y + 1) * bw + x] : 0;
                double nl = Math.Sqrt(nx * nx + ny * ny) + 1e-9; nx /= nl; ny /= nl;
                // Gegenlicht vom Planetensaum hinter der Gestalt, schwach von der Gaswolke links oben, Klinge rechts
                double unterhalb = Weich(PlanetRandY(gx) - 200 * K, PlanetRandY(gx) + 30 * K, gy);
                double planet = Math.Max(0, -0.55 * nx + 0.30 * ny + 0.35) * (0.25 + 0.75 * unterhalb);
                double wolke = Math.Max(0, -0.6 * nx - 0.8 * ny) * 0.35;
                double dS = Abstand(gx + 0.5, gy + 0.5, hx, hy, tx, ty) / K;
                double nah = 1 / (1 + Math.Pow(dS / 40, 2));
                double rot = nah * (0.20 + 0.80 * Math.Max(0, 0.75 * nx + 0.40 * ny) * rand) * 0.85;
                double falten = 0.004 * (1 + Fbm(gx * 0.03 / K, gy * 0.006 / K, 3));
                double r = 0.005 + falten + rand * (0.13 * planet + 0.30 * wolke) + 0.90 * rot;
                double g = 0.005 + falten + rand * (0.26 * planet + 0.05 * wolke) + 0.05 * rot;
                double b = 0.008 + falten * 1.3 + rand * (0.48 * planet + 0.07 * wolke) + 0.06 * rot;
                Mische(gy * W + gx, r, g, b, cov);
            }

        // Griff, dann die Klinge mit Gluehen
        Griff(handX, handY, hx, hy, 3.4 * K);
        Klinge(hx, hy, tx, ty);
        Glut(tx, ty);
    }

    static double[] Weichzeichnen(double[] a, int w, int h, int r)
    {
        double[] t = new double[w * h], o = new double[w * h];
        for (int y = 0; y < h; y++)
        {
            double sum = 0; int cnt = 0;
            for (int x = -r; x < w + r; x++)
            {
                if (x + r < w) { sum += a[y * w + x + r]; cnt++; }
                if (x - r - 1 >= 0) { sum -= a[y * w + x - r - 1]; cnt--; }
                if (x >= 0 && x < w) t[y * w + x] = sum / (2 * r + 1);
            }
        }
        for (int x = 0; x < w; x++)
        {
            double sum = 0;
            for (int y = -r; y < h + r; y++)
            {
                if (y + r < h) sum += t[(y + r) * w + x];
                if (y - r - 1 >= 0) sum -= t[(y - r - 1) * w + x];
                if (y >= 0 && y < h) o[y * w + x] = sum / (2 * r + 1);
            }
        }
        return o;
    }

    static void Griff(double ax, double ay, double bx, double by, double dicke)
    {
        int x0 = (int)(Math.Min(ax, bx) - dicke - 3), x1 = (int)(Math.Max(ax, bx) + dicke + 3);
        int y0 = (int)(Math.Min(ay, by) - dicke - 3), y1 = (int)(Math.Max(ay, by) + dicke + 3);
        for (int y = y0; y <= y1; y++)
            for (int x = x0; x <= x1; x++)
            {
                if (x < 0 || y < 0 || x >= W || y >= H) continue;
                double d = Abstand(x + 0.5, y + 0.5, ax, ay, bx, by);
                double cov = Math.Max(0, Math.Min(1, dicke - d + 0.5));
                if (cov <= 0) continue;
                double glanz = Math.Max(0, 1 - d / dicke) * 0.5;
                Mische(y * W + x, 0.10 + 0.20 * glanz, 0.10 + 0.12 * glanz, 0.11 + 0.12 * glanz, cov);
            }
    }

    static void Klinge(double ax, double ay, double bx, double by)
    {
        int rad = (int)(120 * K);
        int x0 = (int)Math.Min(ax, bx) - rad, x1 = (int)Math.Max(ax, bx) + rad, y0 = (int)Math.Min(ay, by) - rad, y1 = (int)Math.Max(ay, by) + rad;
        for (int y = Math.Max(0, y0); y <= Math.Min(H - 1, y1); y++)
            for (int x = Math.Max(0, x0); x <= Math.Min(W - 1, x1); x++)
            {
                double d = Abstand(x + 0.5, y + 0.5, ax, ay, bx, by) / K;
                double kern = Math.Max(0, Math.Min(1, 2.4 - d));
                double innen = Math.Exp(-d * d / (2 * 3.6 * 3.6));
                double mitte = Math.Exp(-d / 11);
                double weit = Math.Exp(-d / 58);
                double r = 1.7 * kern + 1.25 * innen + 0.42 * mitte + 0.085 * weit;
                double g = 1.45 * kern + 0.10 * innen + 0.030 * mitte + 0.004 * weit;
                double b = 1.45 * kern + 0.12 * innen + 0.040 * mitte + 0.008 * weit;
                Addiere(x, y, r, g, b);
            }
    }

    // Funken an der Klingenspitze und aufsteigende Glut
    static void Glut(double tx, double ty)
    {
        var r = new Random(19);
        for (int n = 0; n < 150; n++)
        {
            double nah = n < 30 ? 1 : 0;
            double ax = tx + (r.NextDouble() - 0.5) * (nah > 0 ? 60 : 700) * K;
            double ay = ty - r.NextDouble() * (nah > 0 ? 40 : 620) * K;
            double dist = Math.Sqrt(Math.Pow(ax - tx, 2) + Math.Pow(ay - ty, 2)) / K;
            double hell = (0.25 + 0.75 * r.NextDouble()) * (nah > 0 ? 1.0 : Math.Exp(-dist / 260)) * 0.9;
            double sg = (0.55 + 0.8 * r.NextDouble()) * K;
            double streif = (2 + r.NextDouble() * 6) * K;
            double wink = -Math.PI / 2 + (r.NextDouble() - 0.5) * 0.9;
            double cg = 0.30 + 0.20 * r.NextDouble();
            for (int st = 0; st < 5; st++)
            {
                double f = 1 - st * 0.2;
                double cx = ax - Math.Cos(wink) * streif * st * 0.5, cy = ay - Math.Sin(wink) * streif * st * 0.5;
                int rad = (int)(sg * 3) + 1;
                for (int y = (int)cy - rad; y <= (int)cy + rad; y++)
                    for (int x = (int)cx - rad; x <= (int)cx + rad; x++)
                    {
                        double ddx = x + 0.5 - cx, ddy = y + 0.5 - cy;
                        double I = hell * f * Math.Exp(-(ddx * ddx + ddy * ddy) / (2 * sg * sg));
                        Addiere(x, y, 1.0 * I, cg * I, 0.08 * I);
                    }
            }
        }
    }

    // ------------------------------------------------------------------ Ausgabe: Vignette, Tonwerte, Dithering
    static byte[] Ausgabe()
    {
        byte[] o = new byte[W * H * 4];
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
            {
                int i = y * W + x;
                double u = (x + 0.5) / W - 0.5, v = (y + 0.5) / H - 0.5;
                double vig = 1 - 0.45 * Math.Pow(Math.Min(1, u * u * 2.2 + v * v * 2.8), 1.25);
                double r = 1 - Math.Exp(-R[i] * vig * 1.12), g = 1 - Math.Exp(-G[i] * vig * 1.12), b = 1 - Math.Exp(-B[i] * vig * 1.12);
                double z = Hash(x, y) - 0.5;
                o[4 * i] = Kanal(b, z); o[4 * i + 1] = Kanal(g, z); o[4 * i + 2] = Kanal(r, z); o[4 * i + 3] = 255;
            }
        return o;
    }

    static byte Kanal(double c, double z)
    {
        int v = (int)Math.Floor(c * 255 + 0.5 + z);
        return (byte)(v < 0 ? 0 : v > 255 ? 255 : v);
    }
}
'@

# Eigenes Bild hat Vorrang: Design\<Design>.png oder .jpg im Werkzeugordner (StarWars.png, CloneWars.png).
# Es wird um 20 % abgedunkelt (30 % war Johannes zu dunkel), damit die Schrift lesbar bleibt. Ohne Datei: StarWars berechnet das Bild mit
# SdWeltraum (roter Sith, passt nicht zu Clone Wars), CloneWars liefert $null -> einfarbiger Hintergrund.
$Script:ImperiumHintergrund = @{}
function Get-ImperiumHintergrund {
    param([string] $Design = 'StarWars')
    if ($Script:ImperiumHintergrund[$Design]) { return $Script:ImperiumHintergrund[$Design] }
    # Liegen mehrere Formate da (z. B. altes PNG und komprimiertes JPG), gewinnt die neueste Datei
    $datei = @(foreach ($e in 'png', 'jpg', 'jpeg') { Join-Path $Script:Ordner "Design\$Design.$e" }) | Where-Object { Test-Path -LiteralPath $_ } |
        Sort-Object { (Get-Item -LiteralPath $_).LastWriteTime } -Descending | Select-Object -First 1
    $bild = $null
    if ($datei) {
        try {
            $bild = New-Object System.Windows.Media.Imaging.BitmapImage
            $bild.BeginInit()
            $bild.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad   # Datei nicht sperren
            $bild.UriSource = New-Object System.Uri($datei)
            $bild.EndInit()
            $bild.Freeze()
        } catch { $bild = $null }
    }
    if ($bild) {
        $flaeche = [System.Windows.Rect]::new(0, 0, $bild.PixelWidth, $bild.PixelHeight)
        $gruppe = [System.Windows.Media.DrawingGroup]::new()
        $gruppe.Children.Add([System.Windows.Media.ImageDrawing]::new($bild, $flaeche))
        $schleier = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(51, 0, 0, 0))
        $gruppe.Children.Add([System.Windows.Media.GeometryDrawing]::new($schleier, $null, [System.Windows.Media.RectangleGeometry]::new($flaeche)))
        [System.Windows.Media.RenderOptions]::SetBitmapScalingMode($gruppe, [System.Windows.Media.BitmapScalingMode]::HighQuality)
        $pinsel = [System.Windows.Media.DrawingBrush]::new($gruppe)
        $pinsel.Stretch = [System.Windows.Media.Stretch]::UniformToFill
        $pinsel.Freeze()
        $Script:ImperiumHintergrund[$Design] = $pinsel
        return $pinsel
    }
    if ($Design -ne 'StarWars') { return $null }
    if (-not ('SdWeltraum' -as [type])) { Add-Type -TypeDefinition $Script:WeltraumCode }
    $breite = 2560; $hoehe = 1440
    $pixel = [SdWeltraum]::Zeichne($breite, $hoehe)
    $bild = [System.Windows.Media.Imaging.BitmapSource]::Create($breite, $hoehe, 96, 96, [System.Windows.Media.PixelFormats]::Bgra32, $null, $pixel, $breite * 4)
    $bild.Freeze()
    $pinsel = [System.Windows.Media.ImageBrush]::new($bild)
    $pinsel.Stretch = [System.Windows.Media.Stretch]::UniformToFill
    [System.Windows.Media.RenderOptions]::SetBitmapScalingMode($pinsel, [System.Windows.Media.BitmapScalingMode]::HighQuality)
    $pinsel.Freeze()
    $Script:ImperiumHintergrund[$Design] = $pinsel
    return $pinsel
}

# Einstellungen.json im Werkzeugordner: Design und automatische Pruefung beim Start
function Import-Einstellungen {
    try {
        if (Test-Path -LiteralPath $Script:EinstellungenDatei) {
            $e = Get-Content -LiteralPath $Script:EinstellungenDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            if ([string]$e.Design -in 'Standard', 'Hell', 'Dunkel', 'StarWars', 'CloneWars') { $Script:Einstellungen.Design = [string]$e.Design }
            if ($null -ne $e.AutoPruefen) { $Script:Einstellungen.AutoPruefen = [bool]$e.AutoPruefen }
            if ([string]$e.LetzterFehlerbericht -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$') { $Script:Einstellungen.LetzterFehlerbericht = [string]$e.LetzterFehlerbericht }
        }
    } catch { }
}

function Save-Einstellungen {
    if ($Script:NurLesen) { return }
    $obj = [ordered]@{ Format = 'Systemdiagnose-Einstellungen/1'; Design = $Script:Einstellungen.Design; AutoPruefen = [bool]$Script:Einstellungen.AutoPruefen; LetzterFehlerbericht = [string]$Script:Einstellungen.LetzterFehlerbericht }
    [IO.File]::WriteAllText($Script:EinstellungenDatei, ($obj | ConvertTo-Json), (New-Object Text.UTF8Encoding($true)))
}

# Sofort uebernehmen: Farben, Schrift und Titelleisten aller offenen Fenster
function Set-DesignWahl {
    param([string] $Design)
    if ($Script:Einstellungen.Design -eq $Design) { return }
    $Script:Einstellungen.Design = $Design
    Set-Design
    foreach ($w in $Script:App.Windows) { Update-Titelleiste $w }
    if (-not $Vorschau) { Save-Einstellungen }
}

function Show-Einstellungen {
    $d = New-Dialog $Script:EinstellungenXaml
    # Erst den Stand setzen, dann die Ereignisse anhaengen - sonst wuerde das Setzen schon umschalten
    foreach ($n in 'Standard', 'Hell', 'Dunkel', 'StarWars', 'CloneWars') {
        $r = $d.FindName("Design$n")
        $r.IsChecked = ($Script:Einstellungen.Design -eq $n)
        $r.Add_Checked({ param($s, $e) Set-DesignWahl ([string]$s.Tag) })
    }
    $auto = $d.FindName('AutoPruefen')
    $auto.IsChecked = [bool]$Script:Einstellungen.AutoPruefen
    $auto.Add_Click({ param($s, $e) $Script:Einstellungen.AutoPruefen = [bool]$s.IsChecked; if (-not $Vorschau) { Save-Einstellungen } })
    $d.FindName('PcText').Text = Get-PcText
    $d.FindName('Einlernen').IsEnabled = [bool]($Script:Daten -and $Script:Daten.Maschine -and (Test-EigeneDaten $Script:Daten))
    $d.FindName('Einlernen').Add_Click({
        param($s, $e)
        $w = [System.Windows.Window]::GetWindow($s)
        if (Save-Profil $Script:Daten) { $w.FindName('PcText').Text = (Get-PcText) + "`nGilt ab der nächsten Prüfung." }
    })
    $d.FindName('Schliessen').Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).Close() })
    if ($Vorschau) { return $d }
    [void]$d.ShowDialog()
}

# ---- Fehlerbericht (02.11.): Fehler anderer PCs als Datei zurück an PC 1, statt dort eine Sitzung zu starten ----
# Quellen: Fehler.log (Programmfehler + leere Läufe), Bereinigung\Ergebnis-*.json (fehlgeschlagen/abgelehnt + Fehlercode),
# letzte Messung/Prüfung. Nur lesen und eine neue Datei auf dem Desktop schreiben.
function Write-Fehlerlog {
    param([string] $Text)
    if ($Script:NurLesen) { return }
    try { [IO.File]::AppendAllText((Join-Path $Script:DatenOrdner 'Fehler.log'), "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Text`r`n`r`n", (New-Object Text.UTF8Encoding($true))) } catch { }
}

function Get-FehlerSeit {
    $grenze = (Get-Date).AddDays(-30)
    $t = [datetime]::MinValue
    if ([datetime]::TryParseExact([string]$Script:Einstellungen.LetzterFehlerbericht, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$t) -and $t -gt $grenze) { return $t }
    return $grenze
}

function Get-FehlerFunde {
    param([datetime] $Seit = (Get-FehlerSeit))
    $funde = New-Object System.Collections.Generic.List[object]
    $log = Join-Path $Script:DatenOrdner 'Fehler.log'
    if (Test-Path -LiteralPath $log) {
        $text = Read-GeteilteDatei $log 65001
        foreach ($m in [regex]::Matches($text, '(?m)^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})  (.*?)(?=^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}  |\z)', 'Singleline')) {
            $z = [datetime]::ParseExact($m.Groups[1].Value, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
            if ($z -lt $Seit) { continue }
            $voll = $m.Groups[2].Value.Trim()
            $zeilen = @($voll -split "`r?`n")
            $code = if ($zeilen.Count -gt 1 -and $zeilen[1] -match '^0x[0-9A-F]{8} ') { $zeilen[1] } else { '' }
            $funde.Add([pscustomobject]@{ Zeit = $m.Groups[1].Value; Quelle = 'App'; Text = $zeilen[0]; Code = $code; Details = $voll })
        }
    }
    foreach ($f in @(Get-ChildItem -LiteralPath $Script:BereinigungsOrdner -Filter 'Ergebnis-*.json' -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $Seit } | Sort-Object Name)) {
        try { $e = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }
        if ($e.Simulation) { continue }
        if ($e.Fehler) { $funde.Add([pscustomobject]@{ Zeit = [string]$e.Zeit; Quelle = "Bereinigung ($($f.Name))"; Text = [string]$e.Fehler; Code = [string]$e.Fehlercode; Details = '' }) }
        foreach ($a in @($e.Aktionen | Where-Object { $_ -and $_.Status -in 'fehlgeschlagen', 'abgelehnt' })) {
            $funde.Add([pscustomobject]@{ Zeit = [string]$e.Zeit; Quelle = "Bereinigung $($a.Id)"; Text = "$($a.Titel) - $($a.Status): $($a.Meldung)"; Code = [string]$a.Fehlercode; Details = '' })
        }
    }
    return $funde.ToArray()
}

function Update-FehlerPunkt {
    try { (Get-Element 'FehlerPunkt').Visibility = if (@(Get-FehlerFunde).Count) { 'Visible' } else { 'Collapsed' } } catch { }
}

function Get-JsonDatei {
    param($Datei)
    if (-not $Datei) { return $null }
    try { return (Get-Content -LiteralPath $Datei.FullName -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}

function New-Fehlerbericht {
    param([string] $Beschreibung)
    $seit = Get-FehlerSeit
    $win = try { Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop } catch { $null }
    $version = Get-AppVersion
    $logText = Read-GeteilteDatei (Join-Path $Script:DatenOrdner 'Fehler.log') 65001
    $protokoll = Read-GeteilteDatei (Join-Path $Script:BereinigungsOrdner 'Protokoll.txt') 65001
    $ergebnisse = @(Get-ChildItem -LiteralPath $Script:BereinigungsOrdner -Filter 'Ergebnis-*.json' -File -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending | Select-Object -First 8 | ForEach-Object { Get-JsonDatei $_ } | Where-Object { $_ })
    $messung = Get-JsonDatei (Get-NeuesteEigeneDatei 'Reinigung-*.json')
    $pruefung = Get-JsonDatei (Get-NeuesteEigeneDatei 'Befunde-*.json')
    # Übermittlungsoptimierung roh (Laptop 02.11.: App 6,5 GB, Datenträgerbereinigung ~80 MB, Löschen schlug fehl)
    $liefer = [ordered]@{}
    try { $s = Get-DeliveryOptimizationPerfSnap -ErrorAction Stop -WarningAction SilentlyContinue; $liefer.CacheSizeBytes = $s.CacheSizeBytes; $liefer.Files = $s.Files; $liefer.DownloadMode = [string]$s.DownloadMode } catch { $liefer.PerfSnapFehler = $_.Exception.Message }
    try { $liefer.Jobs = @(Get-DeliveryOptimizationStatus -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -First 30 Status, FileSize, FileSizeInCache, Priority, IsPinned, PredefinedCallerApplication, ExpireOn | ForEach-Object { [ordered]@{ Status = [string]$_.Status; FileSize = $_.FileSize; FileSizeInCache = $_.FileSizeInCache; Priority = [string]$_.Priority; IsPinned = $_.IsPinned; Aufrufer = [string]$_.PredefinedCallerApplication; Ablauf = [string]$_.ExpireOn } }) } catch { $liefer.StatusFehler = $_.Exception.Message }
    $istAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    return [ordered]@{
        Format = 'Systemdiagnose-Fehlerbericht/1'; Erstellt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Seit = $seit.ToString('yyyy-MM-dd HH:mm:ss')
        Version = $version; PcName = $env:COMPUTERNAME; Maschine = $Script:MaschinenId; IstPc1 = ($Script:MaschinenId -eq $Script:Pc1Id)
        # ProductName sagt auch unter Windows 11 „Windows 10“ - die Build-Nummer entscheidet
        Windows = $(if ($win) { "$(if ([int]$win.CurrentBuild -ge 22000) { $win.ProductName -replace 'Windows 10', 'Windows 11' } else { $win.ProductName }) $($win.DisplayVersion) Build $($win.CurrentBuild).$($win.UBR)" } else { [Environment]::OSVersion.VersionString })
        PowerShell = [string]$PSVersionTable.PSVersion; Sprache = (Get-Culture).Name; Admin = $istAdmin; Ordner = $Script:Ordner; DatenOrdner = $Script:DatenOrdner
        Beschreibung = $Beschreibung
        Funde = @(Get-FehlerFunde $seit)
        FehlerLog = $(if ($logText.Length -gt 30000) { $logText.Substring($logText.Length - 30000) } else { $logText })
        BereinigungProtokoll = (($protokoll -split "`r?`n" | Where-Object { $_ } | Select-Object -Last 150) -join "`n")
        Bereinigungen = $ergebnisse
        Reinigungsmessung = $(if ($messung) { [ordered]@{ Zeit = [string]$messung.ZeitText; Kategorien = @($messung.Kategorien | Where-Object { $_ } | ForEach-Object { [ordered]@{ Id = $_.Id; Bytes = $_.Bytes; Dateien = $_.Dateien; Messbar = $_.Messbar; Laeuft = $_.Laeuft; Sperre = $_.Sperre; Fehler = $_.Fehler } }) } })
        Pruefung = $(if ($pruefung) { [ordered]@{ Zeit = [string]$pruefung.ZeitText; Modus = [string]$pruefung.Modus; Profil = [string]$pruefung.Profil
            Befunde = @($pruefung.Befunde | Where-Object { $_ -and $_.Status -ne 'OK' } | ForEach-Object { [ordered]@{ Typ = $_.Typ; Status = $_.Status; Pruefung = $_.Pruefung; Ergebnis = $_.Ergebnis; Details = @($_.Details | Select-Object -First 8) } }) } })
        Lieferoptimierung = $liefer
    }
}

function Show-Fehlerbericht {
    $d = New-Dialog $Script:FehlerberichtXaml
    $funde = @(Get-FehlerFunde)
    $d.FindName('Gefunden').Text = if ($funde.Count) {
        (@($funde | Select-Object -Last 12 | ForEach-Object { "$($_.Zeit)  ·  $($_.Text)$(if ($_.Code) { "`n      Code: $($_.Code)" })" }) -join "`n") +
        $(if ($funde.Count -gt 12) { "`n… und $($funde.Count - 12) ältere" })
    } else { 'Seit dem letzten Bericht keine Fehler bemerkt. Fällt dir etwas auf (falsche Größe, falscher Befund), beschreibe es unten - Messwerte und Prüfergebnisse kommen trotzdem mit.' }
    $d.FindName('Schliessen').Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).Close() })
    $d.FindName('Erstellen').Add_Click({
        param($s, $e)
        $w = [System.Windows.Window]::GetWindow($s)
        $bericht = New-Fehlerbericht ([string]$w.FindName('Beschreibung').Text.Trim())
        $name = 'Systemdiagnose-Fehlerbericht-{0}-{1}.json' -f ($env:COMPUTERNAME -replace '[^A-Za-z0-9\-]', '_'), (Get-Date -Format 'yyyy-MM-dd_HHmmss')
        $ziel = Join-Path ([Environment]::GetFolderPath('Desktop')) $name
        if ($Script:NurLesen) { $ziel = Join-Path $env:TEMP $name }
        [IO.File]::WriteAllText($ziel, ($bericht | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($true)))
        $Script:Einstellungen.LetzterFehlerbericht = $bericht.Erstellt
        Save-Einstellungen
        Update-FehlerPunkt
        $w.Close()
        if ($Script:NurLesen) { return }
        Start-Process -FilePath 'explorer.exe' -ArgumentList "/select,`"$ziel`""
        Show-Hinweis 'Bericht gespeichert' ("Was? $name liegt auf dem Desktop ($([math]::Max(1, [math]::Round((Get-Item -LiteralPath $ziel).Length / 1KB))) KB).`n`nWas tun? Die Datei an PC 1 schicken (Google Drive, Mail oder USB-Stick). Dort in den Ordner Tools\Systemdiagnose\Rueckmeldungen legen - Claude liest sie dort und baut die Korrektur ein.")
    })
    if ($Vorschau) { return $d }
    [void]$d.ShowDialog()
}

# ---- Update (02.11.): neue Programmdateien von GitHub statt Neuinstallation ----
# Ablauf: beim Start im Hintergrund manifest.json + manifest.sig laden -> Signatur (RSA, $Script:UpdateSchluessel) ->
# neuere Version? Knopf „Update“. Klick: Dateien laden, SHA-256 gegen das signierte Manifest, alte sichern, ersetzen, neu starten.
function Get-AppVersion {
    try { return ([regex]::Match((Get-Content -LiteralPath $Script:Engine -TotalCount 120 -Encoding UTF8) -join "`n", "Version\s*=\s*'([0-9.]+)")).Groups[1].Value } catch { return '' }
}

function Test-UpdateErlaubt {
    if ($Script:NurLesen -or -not $Script:UpdateQuelle) { return $false }
    # Entwicklungsordner auf PC 1 (Git) und Store-Fassung (schreibgeschuetzt) aktualisieren sich nie selbst
    if (Test-Path -LiteralPath (Join-Path $Script:Ordner '.git')) { return $false }
    if ($Script:Ordner -match '\\WindowsApps\\') { return $false }
    return $true
}

function Test-UpdatePfad {
    param([string] $Pfad)
    return ($Pfad -match '^(Design\\)?[A-Za-z0-9][A-Za-z0-9 \-]{0,60}\.(ps1|psd1|cmd|md|jpg|png|txt)$')
}

function Start-UpdatePruefung {
    if (-not (Test-UpdateErlaubt)) { return }
    $ps = [PowerShell]::Create()
    [void]$ps.AddScript({
        param($Quelle)
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $wc = New-Object Net.WebClient
        $wc.Headers['User-Agent'] = 'Systemdiagnose'
        $t = [DateTime]::UtcNow.Ticks
        $q = if ($Quelle -like 'http*') { "?t=$t" } else { '' }   # Zwischenspeicher umgehen; file:// (Netzwerkordner, Test) kennt keine Abfrage
        @{ Manifest = $wc.DownloadData("$Quelle/manifest.json$q"); Signatur = $wc.DownloadString("$Quelle/manifest.sig$q") }
    }).AddArgument($Script:UpdateQuelle)
    $Script:UpdateSuche = @{ Ps = $ps; Handle = $ps.BeginInvoke(); Start = Get-Date }
    # Nur solange die Abfrage laeuft (max. 20 s), danach kein Takt mehr
    $Script:UpdateUhr = New-Object System.Windows.Threading.DispatcherTimer
    $Script:UpdateUhr.Interval = [TimeSpan]::FromMilliseconds(500)
    $Script:UpdateUhr.Add_Tick({
        $s = $Script:UpdateSuche
        if (-not $s.Handle.IsCompleted -and ((Get-Date) - $s.Start).TotalSeconds -lt 20) { return }
        $Script:UpdateUhr.Stop()
        try {
            if (-not $s.Handle.IsCompleted) { $s.Ps.Stop(); return }
            $r = @($s.Ps.EndInvoke($s.Handle))[0]
            if ($s.Ps.HadErrors -or -not $r) { return }
            $u = Read-UpdateManifest $r.Manifest ([string]$r.Signatur)
            if ($u -and [version]$u.Version -gt [version](Get-AppVersion)) {
                $Script:Update = $u
                (Get-Element 'UpdateText').Text = "Update $($u.Version)"
                (Get-Element 'UpdateKnopf').Visibility = 'Visible'
            }
        } catch { } finally { $s.Ps.Dispose() }
    })
    $Script:UpdateUhr.Start()
}

# Nur ein Manifest mit gueltiger Signatur, bekanntem Format und harmlosen Pfaden wird angenommen
function Read-UpdateManifest {
    param([byte[]] $Bytes, [string] $Signatur)
    $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider
    $rsa.PersistKeyInCsp = $false
    $rsa.FromXmlString($Script:UpdateSchluessel)
    if (-not $rsa.VerifyData($Bytes, 'SHA256', [Convert]::FromBase64String($Signatur.Trim()))) { Write-Fehlerlog 'Update: Signatur ungültig - Manifest verworfen'; return $null }
    $m = [Text.Encoding]::UTF8.GetString($Bytes) | ConvertFrom-Json
    if ($m.Format -ne 'Systemdiagnose-Update/1' -or [string]$m.Version -notmatch '^\d+(\.\d+){1,3}$') { return $null }
    foreach ($d in @($m.Dateien)) { if (-not (Test-UpdatePfad ([string]$d.Pfad)) -or [string]$d.Sha256 -notmatch '^[0-9A-F]{64}$') { Write-Fehlerlog "Update: unzulässiger Eintrag $($d.Pfad)"; return $null } }
    return $m
}

function Install-Update {
    param($U)
    if ($Script:Lauf) { Show-Hinweis 'Später aktualisieren' 'Gerade läuft eine Prüfung oder Bereinigung. Danach erneut auf „Update“ klicken.'; return }
    $temp = Join-Path $env:TEMP ('Systemdiagnose-Update-' + $U.Version)
    if (Test-Path -LiteralPath $temp) { Get-ChildItem -LiteralPath $temp -Recurse -File | ForEach-Object { [IO.File]::Delete($_.FullName) } }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $wc = New-Object Net.WebClient
    $wc.Headers['User-Agent'] = 'Systemdiagnose'
    $q = if ($Script:UpdateQuelle -like 'http*') { "?t=$([DateTime]::UtcNow.Ticks)" } else { '' }
    # 1. alles laden und pruefen - erst wenn jede Datei stimmt, wird etwas ersetzt
    foreach ($d in @($U.Dateien)) {
        $ziel = Join-Path $temp $d.Pfad
        [void](New-Item -ItemType Directory -Force -Path (Split-Path $ziel))
        $wc.DownloadFile("$Script:UpdateQuelle/app/$(([string]$d.Pfad -replace '\\', '/') -replace ' ', '%20')$q", $ziel)
        if ((Get-FileHash -LiteralPath $ziel -Algorithm SHA256).Hash -ne [string]$d.Sha256) { throw "Prüfsumme von $($d.Pfad) stimmt nicht - das Update wird gerade hochgeladen oder wurde verändert. Nichts ersetzt; später erneut versuchen." }
    }
    # 2. alte Fassung sichern (nur die letzte), 3. ersetzen
    $alt = Get-AppVersion
    $sicherung = Join-Path $Script:DatenOrdner 'Update-Sicherung'
    if (Test-Path -LiteralPath $sicherung) { Get-ChildItem -LiteralPath $sicherung -Recurse -File | ForEach-Object { [IO.File]::Delete($_.FullName) } }
    foreach ($d in @($U.Dateien)) {
        $jetzt = Join-Path $Script:Ordner $d.Pfad
        if (Test-Path -LiteralPath $jetzt) { $s = Join-Path $sicherung $d.Pfad; [void](New-Item -ItemType Directory -Force -Path (Split-Path $s)); [IO.File]::Copy($jetzt, $s, $true) }
    }
    foreach ($d in @($U.Dateien)) {
        $jetzt = Join-Path $Script:Ordner $d.Pfad
        [void](New-Item -ItemType Directory -Force -Path (Split-Path $jetzt))
        [IO.File]::Copy((Join-Path $temp $d.Pfad), $jetzt, $true)
    }
    try { [IO.File]::WriteAllText((Join-Path $Script:Ordner 'Version.txt'), [string]$U.Version) } catch { }
    # Eigenes Protokoll statt Fehler.log - ein Update ist kein Fehler (roter Punkt)
    try { [IO.File]::AppendAllText((Join-Path $Script:DatenOrdner 'Update.log'), "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $alt -> $($U.Version)`r`n", (New-Object Text.UTF8Encoding($true))) } catch { }
    Restart-App
}

function Restart-App {
    # Einzelinstanz freigeben, sonst holt die neue Instanz nur dieses Fenster nach vorn und beendet sich
    try { $Script:Einzelinstanz.ReleaseMutex() } catch { }
    $starter = Join-Path $Script:Ordner 'Systemdiagnose.exe'
    if (Test-Path -LiteralPath $starter) { Start-Process -FilePath $starter -WorkingDirectory $Script:Ordner }
    else { Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Join-Path $Script:Ordner 'Systemdiagnose-App.ps1')`"" -WorkingDirectory $Script:Ordner -WindowStyle Hidden }
    $Script:Fenster.Close()
}

function Show-Update {
    $u = $Script:Update
    if (-not $u) { return }
    $d = New-Dialog $Script:UpdateXaml
    $d.FindName('Titel').Text = "Version $($u.Version) ist da"
    $d.FindName('Stand').Text = "Installiert: $(Get-AppVersion)  ·  Neu: $($u.Version) vom $($u.Datum)"
    $d.FindName('Neu').Text = if (@($u.Neu).Count) { (@($u.Neu) | ForEach-Object { "•  $_" }) -join "`n" } else { 'Verbesserungen und Fehlerkorrekturen' }
    $d.FindName('Spaeter').Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).Close() })
    $d.FindName('Jetzt').Add_Click({
        param($s, $e)
        $w = [System.Windows.Window]::GetWindow($s)
        $w.Close()
        try { Install-Update $Script:Update }
        catch { Write-Fehlerlog "Update fehlgeschlagen: $($_.Exception.Message)"; Update-FehlerPunkt; Show-Hinweis 'Update nicht installiert' "Was? $($_.Exception.Message)`n`nWarum? Download oder Prüfung fehlgeschlagen - die bisherige Version ist unverändert.`n`nWas tun? Später erneut versuchen. Bleibt es dabei: Fehler melden." }
    })
    if ($Vorschau) { return $d }
    [void]$d.ShowDialog()
}

function Set-Titelleiste {
    param($Fenster)
    $Fenster.Add_SourceInitialized({ param($s, $e) Update-Titelleiste $s })
    if ($Script:Symbol) { $Fenster.Icon = $Script:Symbol }
}

function Update-Titelleiste {
    param($Fenster)
    try {
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($Fenster)).Handle
        if ($hwnd -ne [IntPtr]::Zero) { [SdFenster]::Titelleiste($hwnd, (-not $Script:Hell), (ConvertTo-Bgr $Script:Farben['Bg']), (ConvertTo-Bgr $Script:Farben['Text'])) }
    } catch { }
}

function New-Dialog {
    param([string] $Xaml)
    $d = New-XamlObjekt $Xaml
    if ($Script:Fenster -and $Script:Fenster.IsLoaded) { $d.Owner = $Script:Fenster }
    Set-Titelleiste $d
    return $d
}

function Show-Hinweis {
    param([string] $Titel, [string] $Text)
    $d = New-Dialog $Script:HinweisXaml
    $d.FindName('Titel').Text = $Titel
    $d.FindName('Text').Text = $Text
    $d.FindName('Ok').Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).Close() })
    [void]$d.ShowDialog()
}

function Read-GeteilteDatei {
    param([string] $Pfad, [int] $Codepage = 65001)
    try {
        $fs = [IO.File]::Open($Pfad, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
        try {
            $leser = New-Object IO.StreamReader($fs, [Text.Encoding]::GetEncoding($Codepage))
            return $leser.ReadToEnd()
        } finally { $fs.Close() }
    } catch { return '' }
}

# Englische Zustandswerte aus Windows (Starttyp von Diensten, Zustand geplanter Aufgaben) fuer die Anzeige
$Script:Werte = @{ Unknown = 'unbekannt'; Manual = 'manuell'; Auto = 'automatisch'; Automatic = 'automatisch'; Disabled = 'deaktiviert'
    Boot = 'beim Booten'; Ready = 'bereit'; Running = 'läuft'; Stopped = 'gestoppt'; Queued = 'wartet'; True = 'ja'; False = 'nein' }

function ConvertTo-Tabelle {
    param([object[]] $Zeilen, [switch] $Deutsch)
    $zeilen = @($Zeilen | Where-Object { $null -ne $_ })
    if ($zeilen.Count -eq 0) { return $null }
    $dt = New-Object System.Data.DataTable
    $spalten = New-Object System.Collections.Generic.List[string]
    foreach ($z in $zeilen) { foreach ($p in $z.PSObject.Properties) { if (-not $spalten.Contains($p.Name)) { $spalten.Add($p.Name) } } }
    foreach ($s in $spalten) { [void]$dt.Columns.Add($s, [string]) }
    foreach ($z in $zeilen) {
        $r = $dt.NewRow()
        foreach ($s in $spalten) {
            $v = $z.$s
            $r[$s] = if ($null -eq $v) { '' } elseif ($v -is [array]) { (@($v) -join ', ') } else { [string]$v }
            if ($Deutsch -and $Script:Werte.ContainsKey([string]$r[$s])) { $r[$s] = $Script:Werte[[string]$r[$s]] }
        }
        [void]$dt.Rows.Add($r)
    }
    return , $dt
}

# ============================================================================
#  Befunde laden und anzeigen
# ============================================================================

function Import-Ergebnis {
    # -Behalten (beim Aktualisieren): Filter und ausgewaehlten Befund beibehalten
    param([string] $Pfad, [switch] $Behalten)
    $d = Get-Content -LiteralPath $Pfad -Raw -Encoding UTF8 | ConvertFrom-Json
    $alt = (Get-Element 'BefundListe').SelectedItem
    $merken = if ($Behalten -and $alt) { $alt.Schluessel } else { $null }
    $liste = New-Object 'System.Collections.ObjectModel.ObservableCollection[SdBefund]'
    foreach ($b in @($d.Befunde)) {
        $o = New-Object SdBefund
        $o.Nr = [int]$b.Nr; $o.Bereich = [string]$b.Bereich; $o.BereichTitel = [string]$b.BereichTitel; $o.BereichRang = [int]$b.BereichRang
        $o.Pruefung = [string]$b.Pruefung; $o.Status = [string]$b.Status; $o.Rang = [int]$b.Rang; $o.Ergebnis = [string]$b.Ergebnis
        $o.Was = [string]$b.Was; $o.Ursache = [string]$b.Ursache; $o.Massnahme = [string]$b.Massnahme; $o.Empfehlung = [string]$b.Empfehlung
        $o.Hinweis = [string]$b.Hinweis
        $o.Schluessel = "$($b.Typ)|$($b.Pruefung)"
        $o.Zuletzt = [string]$b.Zuletzt
        $o.Details = ConvertTo-Tabelle @($b.Details)
        foreach ($l in @($b.Loesungen)) {
            if (-not $l) { continue }
            $x = New-Object SdLoesung
            $x.Id = [string]$l.Id; $x.Titel = [string]$l.Titel; $x.Art = [string]$l.Art; $x.Umkehrbar = [bool]$l.Umkehrbar
            $x.Dauer = [string]$l.Dauer; $x.Beschreibung = [string]$l.Beschreibung
            $x.ParameterJson = if ($null -ne $l.Parameter) { ConvertTo-Json -InputObject $l.Parameter -Depth 6 -Compress } else { '{}' }
            $x.BefundNr = $o.Nr; $x.BefundTitel = $o.Pruefung; $x.Empfehlung = $o.Empfehlung
            $o.Loesungen.Add($x)
        }
        $liste.Add($o)
    }
    Update-ArchivStand $liste
    Update-AutoArchiv $liste $d
    $Script:Daten = $d
    $Script:ErgebnisPfad = $Pfad
    $Script:Liste = $liste
    Test-NeuerPc $d
    $Script:LetzterModus = [string]$d.Modus

    Update-Zaehler
    $rechte = if ($d.Admin) { 'mit Adminrechten' } else { 'ohne Adminrechte' }
    $dauer = if ([int]$d.DauerSekunden -ge 60) { '{0}:{1:00} Min.' -f [Math]::Floor([int]$d.DauerSekunden / 60), ([int]$d.DauerSekunden % 60) } else { "$($d.DauerSekunden) s" }
    (Get-Element 'Untertitel').Text = "Letzte Prüfung: $($d.ZeitText)  ·  $($d.Modus -replace 'Gruendlich|Tief', 'Gründlich'), $rechte  ·  Dauer $dauer  ·  $($d.Profil)"

    $Script:Ansicht = [System.Windows.Data.CollectionViewSource]::GetDefaultView($liste)
    $Script:Ansicht.GroupDescriptions.Clear()
    [void]$Script:Ansicht.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription 'BereichTitel'))
    $Script:Ansicht.SortDescriptions.Clear()
    foreach ($s in 'BereichRang', 'Rang', 'Nr') { $Script:Ansicht.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $s, ([System.ComponentModel.ListSortDirection]::Ascending))) }
    $Script:Ansicht.Filter = [Predicate[object]] { param($o) return (Test-Filter $o) }
    (Get-Element 'BefundListe').ItemsSource = $Script:Ansicht

    $offen = @($liste | Where-Object { $_.IstOffen }).Count
    if ($Behalten -and -not ($Script:Filter -eq 'Archiv' -and -not @($liste | Where-Object { $_.IstArchiviert }).Count)) { Set-Filter $Script:Filter } elseif ($offen -gt 0) { Set-Filter 'Offen' } else { Set-Filter 'Alle' }
    if ($merken) {
        $wieder = @($liste | Where-Object { $_.Schluessel -eq $merken })[0]
        if ($wieder -and (Test-Filter $wieder)) { (Get-Element 'BefundListe').SelectedItem = $wieder; Update-Detail }
    }

    $aend = @($d.Aenderungen)
    # Direkt zuweisen: "if (...) { $x.DefaultView }" wuerde die Ansicht in einzelne Zeilen auspacken
    $tab = ConvertTo-Tabelle $aend -Deutsch
    (Get-Element 'AenderungenGrid').ItemsSource = $null
    if ($tab) { (Get-Element 'AenderungenGrid').ItemsSource = $tab.DefaultView }
    (Get-Element 'AenderungenGrid').Visibility = if ($tab) { 'Visible' } else { 'Collapsed' }   # sonst leerer Tabellenkopf
    (Get-Element 'VerlaufAenderungen').Content = if ($aend.Count) { "Systemänderungen ($($aend.Count))" } else { 'Systemänderungen' }
    (Get-Element 'AenderungenInfo').Text = if (-not $d.AltZeit) { 'Erster Lauf - ab dem nächsten Lauf steht hier, was sich verändert hat: Treiber, Dienste, Autostart, Programme, Einstellungen, Monitore und Speicherplatz.' } elseif ($aend.Count -eq 0) { "Keine Änderungen seit dem Lauf vom $($d.AltZeit)." } else { "$($aend.Count) Änderung(en) seit der Prüfung vom $($d.AltZeit)" }
    $q = ConvertTo-Tabelle @($d.Quellen)
    (Get-Element 'QuellenGrid').ItemsSource = $null
    if ($q) { (Get-Element 'QuellenGrid').ItemsSource = $q.DefaultView }
    Update-Auswahl
    if (Test-EntwicklungOffen) { Update-Entwicklung }
}

# Verlauf: eine der drei Ansichten zeigen; Entwicklung erst beim Oeffnen laden
function Set-VerlaufAnsicht {
    param([ValidateSet('Massnahmen', 'Aenderungen', 'Entwicklung')] [string] $Ansicht)
    foreach ($a in 'Massnahmen', 'Aenderungen', 'Entwicklung') {
        (Get-Element "Ansicht$a").Visibility = if ($a -eq $Ansicht) { 'Visible' } else { 'Collapsed' }
        (Get-Element "Verlauf$a").IsChecked = ($a -eq $Ansicht)
    }
    if (Test-EntwicklungOffen) { Update-Entwicklung }
}

function Test-EntwicklungOffen {
    return ((Get-Element 'Reiter').SelectedItem -eq (Get-Element 'ReiterVerlauf') -and (Get-Element 'VerlaufEntwicklung').IsChecked)
}

# Datenquellen & Sicherheit (Reiter ohne Kopf): oeffnen bzw. zurueck zum vorigen Reiter
function Show-Quellen {
    $r = Get-Element 'Reiter'
    if ($r.SelectedItem -eq (Get-Element 'ReiterQuellen')) { $r.SelectedItem = $Script:ReiterVorQuellen; return }
    $Script:ReiterVorQuellen = $r.SelectedItem
    $r.SelectedItem = Get-Element 'ReiterQuellen'
}

function Test-Filter {
    param($o)
    # Archivierte stehen nur unter „Archiviert“ und (abgeblendet) unter „Alle Prüfungen“
    $ok = switch ($Script:Filter) {
        'Offen'   { $o.IstOffen }
        'Loesbar' { $o.HatLoesungen -and -not $o.IstArchiviert }
        'Archiv'  { $o.IstArchiviert }
        'Alle'    { $true }
        default   { $o.Status -eq $Script:Filter -and -not $o.IstArchiviert }
    }
    if (-not $ok) { return $false }
    if ($Script:SuchText) { return $o.Suchtext.Contains($Script:SuchText) }
    return $true
}

function Set-Filter {
    param([string] $Filter)
    $Script:Filter = $Filter
    $Script:FilterWirdGesetzt = $true
    (Get-Element 'FilterOffen').IsChecked = ($Filter -eq 'Offen')
    (Get-Element 'FilterInfo').IsChecked = ($Filter -eq 'Info')
    (Get-Element 'FilterAlle').IsChecked = ($Filter -eq 'Alle')
    (Get-Element 'FilterArchiv').IsChecked = ($Filter -eq 'Archiv')
    $Script:FilterWirdGesetzt = $false
    Update-Liste
}

function Update-Liste {
    if (-not $Script:Ansicht) { return }
    $Script:Ansicht.Refresh()
    $lb = Get-Element 'BefundListe'
    $anzahl = @($Script:Ansicht).Count
    $leer = Get-Element 'ListeLeer'
    if ($anzahl -eq 0) {
        $leer.Text = if ($Script:Filter -eq 'Offen') { 'Kein Handlungsbedarf - nichts Kritisches und keine Warnungen. Unter „Alle Prüfungen“ steht jede einzelne Prüfung.' } elseif ($Script:Filter -eq 'Archiv') { 'Nichts archiviert.' } else { 'Keine passenden Befunde.' }
        $leer.Visibility = 'Visible'
    } else {
        $leer.Visibility = 'Collapsed'
        if ($null -eq $lb.SelectedItem -or -not (Test-Filter $lb.SelectedItem)) { $lb.SelectedIndex = 0 }
    }
    Update-Detail
}

function Update-Detail {
    $lb = Get-Element 'BefundListe'
    $hat = $null -ne $lb.SelectedItem
    (Get-Element 'Detail').Visibility = if ($hat) { 'Visible' } else { 'Collapsed' }
    (Get-Element 'DetailLeer').Visibility = if ($hat) { 'Collapsed' } else { 'Visible' }
    if ($hat) { (Get-Element 'Detail').ScrollToTop() }
}

function Get-AlleLoesungen {
    if (-not $Script:Liste) { return }
    foreach ($b in $Script:Liste) { foreach ($l in $b.Loesungen) { $l } }
}

function Get-AusgewaehlteLoesungen {
    Get-AlleLoesungen | Where-Object { $_.IstBereinigung -and $_.Ausgewaehlt }
}

function Update-Auswahl {
    $wahl = @(Get-AusgewaehlteLoesungen)
    $knopf = Get-Element 'BereinigenKnopf'
    # Auswaehlbar ist nur, was die App selbst beheben kann - bei aktiven (nicht archivierten) Befunden
    $behebbar = @($Script:Liste | Where-Object { $_ -and -not $_.IstArchiviert -and $_.HatBereinigung })
    (Get-Element 'EmpfohleneKnopf').IsEnabled = [bool]$behebbar.Count
    (Get-Element 'AuswahlLeerenKnopf').IsEnabled = [bool]$wahl.Count
    if ($wahl.Count -eq 0) {
        (Get-Element 'AuswahlText').Text = if ($behebbar.Count) { [string]$behebbar.Count + ' Befund(e) in der App behebbar - unter „Lösung“ auswählen oder „Empfohlene auswählen“.' }
            else { 'Nichts in der App behebbar. Hinweise, bei denen die App nichts tun kann, liegen automatisch im Archiv.' }
        (Get-Element 'BereinigenText').Text = 'Beheben'
        $knopf.IsEnabled = $false
    } else {
        $nu = @($wahl | Where-Object { -not $_.Umkehrbar }).Count
        (Get-Element 'AuswahlText').Text = ('{0} Maßnahme(n) ausgewählt  ·  {1} umkehrbar  ·  {2} nicht umkehrbar' -f $wahl.Count, ($wahl.Count - $nu), $nu)
        (Get-Element 'BereinigenText').Text = "Beheben ($($wahl.Count))"
        $knopf.IsEnabled = ($null -eq $Script:Lauf)
    }
}

# ============================================================================
#  PC-Erkennung: die App kann auf jeden PC kopiert werden und arbeitet nur mit dessen Daten
# ============================================================================
# Gleiche Kennung wie in der Engine (Rechnername + Mainboard-UUID). Das eingebaute Profil der Engine gilt nur fuer PC 1.
$Script:Pc1Id = 'DESKTOP-620U4TE|17566DA4-794F-1A1F-84B3-B082E2460564'
$Script:MaschinenId = ''
$Script:ProfilHinweis = ''

function Get-MaschinenId {
    $uuid = ''
    try { $uuid = [string](Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop).UUID } catch { }
    if ($uuid -match '^(0{8}|F{8})-') { $uuid = '' }
    return "$env:COMPUTERNAME|$uuid"
}

# Gehoert eine Befund-/Messdatei zu diesem PC? Alte Dateien ohne Kennung stammen von PC 1.
function Test-EigeneDaten {
    param($Daten)
    $id = if ($Daten.Maschine -is [string]) { [string]$Daten.Maschine } elseif ($Daten.Maschine) { [string]$Daten.Maschine.Id } else { '' }
    if ($id) { return ($id -eq $Script:MaschinenId) }
    return ($Script:MaschinenId -eq $Script:Pc1Id)
}

function Get-NeuesteEigeneDatei {
    param([string] $Muster)
    foreach ($f in @(Get-ChildItem -LiteralPath $Script:BerichtOrdner -Filter $Muster -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 40)) {
        try { if (Test-EigeneDaten (Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json)) { return $f } } catch { }
    }
}

function Get-ProfilDatei { return (Join-Path $Script:DatenOrdner ('Profile\' + ($Script:MaschinenId -replace '[^A-Za-z0-9\-]', '_') + '.json')) }

# Sollwerte einlernen: der aktuelle Zustand (Monitore, Arbeitsspeicher, Energieplan) wird zum Soll. Die Engine liest
# die Datei bei jedem Lauf; auf PC 1 ergaenzt sie das eingebaute Profil (Dienste usw. bleiben).
function Save-Profil {
    param($Daten)
    if ($Script:NurLesen) { return $null }
    $m = $Daten.Maschine
    if (-not $m -or -not $m.Id -or -not (Test-EigeneDaten $Daten)) { return $null }
    $datei = Get-ProfilDatei
    $ordner = Split-Path -Parent $datei
    if (-not (Test-Path -LiteralPath $ordner)) { [void](New-Item -ItemType Directory -Path $ordner) }
    $name = if ([string]$m.ProfilQuelle -eq 'neu') {
        $cpu = ([string]$m.Cpu -replace '\s+\d+-Core Processor|\(R\)|\(TM\)|CPU\s+@.*$', '').Trim()
        $gpu = (@(([string]$m.Gpu) -split ',\s*') | Where-Object { $_ -notmatch 'Basic|Virtual|Parsec|Meta' } | Select-Object -First 1)
        (@([string]$m.Name, $cpu, $gpu) | Where-Object { $_ }) -join ' - '
    } else { [string]$Daten.Profil }
    $obj = [ordered]@{
        Format = 'Systemdiagnose-Profil/1'; Maschine = [string]$m.Id; Name = $name; Eingelernt = (Get-Date).ToString('dd.MM.yyyy HH:mm')
        Monitore = @(@($m.Monitore) | Where-Object { $_ } | ForEach-Object { [ordered]@{ Edid = [string]$_.Edid; Name = [string]$_.Name; Aufloesung = [string]$_.Aufloesung; Hz = [double]$_.Hz; Hdr = [bool]$_.HdrAktiv } })
        RamModule = [int]$m.RamModule; RamMhz = [int]$m.RamMhz; Energieplan = [string]$m.Energieplan
    }
    [IO.File]::WriteAllText($datei, ($obj | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding($true)))
    return $obj
}

# Erster Lauf auf einem neuen PC: Sollwerte still einlernen und einmal erklaeren, was passiert ist
function Test-NeuerPc {
    param($Daten)
    if ($Vorschau -or -not $Daten.Maschine -or [string]$Daten.Maschine.ProfilQuelle -ne 'neu' -or (Test-Path -LiteralPath (Get-ProfilDatei))) { return }
    $p = Save-Profil $Daten
    if (-not $p) { return }
    $mon = @($p.Monitore | ForEach-Object { "$($_.Aufloesung) @ $($_.Hz) Hz" + $(if ($_.Hdr) { ' HDR' } else { '' }) }) -join ', '
    $Script:ProfilHinweis = "Neuer PC erkannt: $($p.Name).`n`nDer aktuelle Zustand gilt ab der nächsten Prüfung als Sollwert:`n• Monitore: $(if ($mon) { $mon } else { 'keine erkannt' })`n• Arbeitsspeicher: $($p.RamModule) Riegel mit $($p.RamMhz) MHz`n• Energieplan: $($p.Energieplan)`n`nWeicht später etwas davon ab (z. B. ein Monitor nur noch mit 60 Hz oder ein Riegel fehlt), meldet die App es. " + 'Stimmt der Zustand gerade nicht, in den Einstellungen (Zahnrad) „Sollwerte neu einlernen“.'
}

function Show-ProfilHinweis {
    if (-not $Script:ProfilHinweis) { return }
    $t = $Script:ProfilHinweis; $Script:ProfilHinweis = ''
    Show-Hinweis 'Neuer PC erkannt' $t
}

# Kurzbeschreibung fuer die Einstellungen
function Get-PcText {
    $datei = Get-ProfilDatei
    $name = if ($Script:Daten) { [string]$Script:Daten.Profil } else { $env:COMPUTERNAME }
    if (Test-Path -LiteralPath $datei) {
        try { $p = Get-Content -LiteralPath $datei -Raw -Encoding UTF8 | ConvertFrom-Json; return "$($p.Name)`nSollwerte eingelernt am $($p.Eingelernt)" } catch { }
    }
    if ($Script:MaschinenId -eq $Script:Pc1Id) { return "$name`nSollwerte: eingebaut (PC 1)" }
    return "$name`nNoch keine Sollwerte - werden bei der ersten Prüfung eingelernt"
}

# ============================================================================
#  Archiv: Befunde ausblenden, die erledigt oder nicht mehr relevant sind
# ============================================================================
# Archiv.json im Werkzeugordner haelt je Pruefung (Typ|Pruefung) den Status, das neueste Ereignis und den
# Diagnose-Stand beim Archivieren. Der Befund kommt zurueck, sobald er schlimmer wird oder ein neueres
# Ereignis meldet; ist er wieder OK, faellt der Eintrag weg.

# Eintraege anderer PCs (Ordner kopiert) bleiben in der Datei, gelten hier aber nicht; alte ohne Kennung stammen von PC 1
function Import-Archiv {
    $Script:Archiv = @{}
    $Script:ArchivFremd = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $Script:ArchivDatei)) { return }
    try {
        foreach ($e in @((Get-Content -LiteralPath $Script:ArchivDatei -Raw -Encoding UTF8 | ConvertFrom-Json).Eintraege)) {
            if (-not $e -or -not $e.Schluessel) { continue }
            $eigen = if ($e.Maschine) { [string]$e.Maschine -eq $Script:MaschinenId } else { $Script:MaschinenId -eq $Script:Pc1Id }
            if ($eigen) { $Script:Archiv[[string]$e.Schluessel] = $e } else { $Script:ArchivFremd.Add($e) }
        }
    } catch { }
}

function Save-Archiv {
    if ($Script:NurLesen) { return }
    # .ToArray() statt @($Script:ArchivFremd): @() auf eine leere List[object] wirft in PS 5.1 „Die Argumenttypen stimmen nicht ueberein“
    $obj = [ordered]@{ Format = 'Systemdiagnose-Archiv/1'; Eintraege = @(@($Script:Archiv.Values) + $Script:ArchivFremd.ToArray() | Where-Object { $_ } | Sort-Object Archiviert) }
    [IO.File]::WriteAllText($Script:ArchivDatei, ($obj | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding($true)))
}

function Format-ArchivZeit {
    param([string] $Zeit)
    try { return [datetime]::ParseExact($Zeit, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture).ToString('dd.MM.yyyy HH:mm') }
    catch { return $Zeit }
}

# Beim Laden jedes Ergebnisses: archiviert lassen, zurueckholen oder austragen
function Update-ArchivStand {
    param($Liste)
    $geaendert = $false
    foreach ($o in $Liste) {
        if (-not $Script:Archiv.ContainsKey($o.Schluessel)) { continue }
        $e = $Script:Archiv[$o.Schluessel]
        if ($o.Status -eq 'OK') { $Script:Archiv.Remove($o.Schluessel); $geaendert = $true; continue }
        if ($e.Angeheftet) { continue }
        # Ohne Ereigniszeit beim Archivieren gilt der damalige Diagnose-Stand als Grenze
        $grenze = if ($e.Zuletzt) { [string]$e.Zuletzt } else { [string]$e.Stand }
        $grund = $null
        if ((Get-StatusStufe $o.Status) -gt (Get-StatusStufe ([string]$e.Status))) { $grund = "jetzt $($o.Status) statt $($e.Status)" }
        elseif ($o.Zuletzt -and $grenze -and [string]$o.Zuletzt -gt $grenze) { $grund = 'neues Ereignis am ' + (Format-ArchivZeit $o.Zuletzt) }
        if ($grund) {
            $o.ArchivHinweis = "Wieder aufgetaucht ($grund) - archiviert hattest du den Befund am $(Format-ArchivZeit $e.Archiviert)."
            $Script:Archiv.Remove($o.Schluessel); $geaendert = $true
        } else {
            $o.ArchiviertAm = Format-ArchivZeit $e.Archiviert
            $o.AutoArchiviert = [bool]$e.Automatisch
            $o.IstArchiviert = $true
        }
    }
    if ($geaendert) { Save-Archiv }
}

# Kann die App bei diesem Befund nichts tun? Kritisches bleibt immer sichtbar (Hardwarefehler, Bluescreens),
# alles mit Loesung (Beheben-Auswahl, Einstellungsseite, Befehl) ebenfalls
function Test-AutoArchivierbar {
    param($Gleiche)
    if (-not @($Gleiche).Count) { return $false }
    return -not @($Gleiche | Where-Object { $_.Status -notin 'Warnung', 'Info' -or $_.HatLoesungen }).Count
}

# Nach Update-ArchivStand: nicht behebbare Warnungen/Hinweise gleich archivieren. Wieder aufgetauchte (neues Ereignis,
# schlimmer) bleiben fuer diesen Stand sichtbar, von Hand wiederhergestellte (Angeheftet) dauerhaft.
function Update-AutoArchiv {
    param($Liste, $Daten)
    $geaendert = $false
    foreach ($gruppe in @($Liste | Where-Object { -not $_.IstArchiviert } | Group-Object Schluessel)) {
        $gleiche = @($gruppe.Group)
        if (-not (Test-AutoArchivierbar $gleiche) -or @($gleiche | Where-Object { $_.HatArchivHinweis }).Count) { continue }
        $e = $Script:Archiv[$gruppe.Name]
        if ($e -and $e.Angeheftet) { continue }
        $jetzt = Get-Date
        $schlimmster = @($gleiche | Sort-Object { Get-StatusStufe $_.Status } -Descending)[0]
        $Script:Archiv[$gruppe.Name] = [pscustomobject]@{
            Schluessel = $gruppe.Name; Pruefung = $schlimmster.Pruefung; Status = $schlimmster.Status; Ergebnis = $schlimmster.Ergebnis
            Zuletzt = [string]$schlimmster.Zuletzt; Stand = [string]$Daten.Zeit; Archiviert = $jetzt.ToString('yyyy-MM-dd HH:mm:ss')
            Maschine = $Script:MaschinenId; Automatisch = $true
        }
        foreach ($o in $gleiche) { $o.ArchiviertAm = $jetzt.ToString('dd.MM.yyyy HH:mm'); $o.AutoArchiviert = $true; $o.IstArchiviert = $true }
        $geaendert = $true
    }
    if ($geaendert) { Save-Archiv }
}

function Set-Archiviert {
    param($Befund, [bool] $An)
    if (-not $Befund -or -not $Script:Daten) { return }
    $gleiche = @($Script:Liste | Where-Object { $_.Schluessel -eq $Befund.Schluessel })
    if ($An) {
        $jetzt = Get-Date
        $schlimmster = @($gleiche | Sort-Object { Get-StatusStufe $_.Status } -Descending)[0]
        $Script:Archiv[$Befund.Schluessel] = [pscustomobject]@{
            Schluessel = $Befund.Schluessel; Pruefung = $Befund.Pruefung; Status = $schlimmster.Status; Ergebnis = $schlimmster.Ergebnis
            Zuletzt = [string]$schlimmster.Zuletzt; Stand = [string]$Script:Daten.Zeit; Archiviert = $jetzt.ToString('yyyy-MM-dd HH:mm:ss')
            Maschine = $Script:MaschinenId
        }
        foreach ($o in $gleiche) {
            foreach ($l in $o.Loesungen) { $l.Ausgewaehlt = $false }
            $o.ArchiviertAm = $jetzt.ToString('dd.MM.yyyy HH:mm'); $o.ArchivHinweis = $null; $o.AutoArchiviert = $false; $o.IstArchiviert = $true
        }
    } else {
        # Waere er automatisch archiviert worden, merkt sich das Archiv „Angeheftet“ - sonst landet er beim naechsten Laden wieder dort
        if (Test-AutoArchivierbar $gleiche) {
            $Script:Archiv[$Befund.Schluessel] = [pscustomobject]@{ Schluessel = $Befund.Schluessel; Pruefung = $Befund.Pruefung; Status = $Befund.Status; Angeheftet = $true
                Archiviert = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Maschine = $Script:MaschinenId }
        } else { $Script:Archiv.Remove($Befund.Schluessel) }
        foreach ($o in $gleiche) { $o.AutoArchiviert = $false; $o.IstArchiviert = $false }
    }
    Save-Archiv
    Update-Zaehler
    Update-Auswahl
    if ($Script:Filter -eq 'Archiv' -and -not @($Script:Liste | Where-Object { $_.IstArchiviert }).Count) { Set-Filter 'Offen' } else { Update-Liste }
}

# Zaehler oben und Chip „Archiviert“ - archivierte Befunde zaehlen nicht mit
function Update-Zaehler {
    $aktiv = @($Script:Liste | Where-Object { -not $_.IstArchiviert })
    foreach ($s in 'Kritisch', 'Warnung', 'Info') { (Get-Element "Zahl$s").Text = [string]@($aktiv | Where-Object { $_.Status -eq $s }).Count }
    (Get-Element 'ZahlOffen').Text = [string]@($aktiv | Where-Object { $_.Status -in 'Kritisch', 'Warnung' }).Count; (Get-Element 'ZahlAlle').Text = [string]$aktiv.Count
    $n = @($Script:Liste | Where-Object { $_.IstArchiviert }).Count
    $chip = Get-Element 'FilterArchiv'
    $chip.Content = "Archiv ($n)"
    $chip.Visibility = if ($n) { 'Visible' } else { 'Collapsed' }
    Update-Note
}

# Alter des Stands in Worten
function Get-StandAlter {
    if (-not $Script:Daten -or -not $Script:Daten.Zeit) { return $null }
    try { return (Get-Date) - [datetime]::ParseExact([string]$Script:Daten.Zeit, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) } catch { return $null }
}

# Gesamtnote neben dem Titel: Zustand in einem Wort, offene Befunde, wie alt der Stand ist
function Update-Note {
    $note = Get-Element 'Note'
    if (-not $Script:Liste) { $note.Visibility = 'Collapsed'; return }
    $aktiv = @($Script:Liste | Where-Object { -not $_.IstArchiviert })
    $krit = @($aktiv | Where-Object { $_.Status -eq 'Kritisch' }).Count
    $warn = @($aktiv | Where-Object { $_.Status -eq 'Warnung' }).Count
    if ($krit) { $farbe = 'Crit'; $text = "Handlungsbedarf · $krit kritisch" + $(if ($warn) { ", $warn Warnung(en)" } else { '' }) }
    elseif ($warn) { $farbe = 'Warn'; $text = "$warn Warnung(en)" }
    else { $farbe = 'Ok'; $text = 'System in Ordnung' }
    $alter = Get-StandAlter
    if ($alter) {
        $text += '  ·  geprüft ' + $(if ($alter.TotalMinutes -lt 2) { 'gerade eben' } elseif ($alter.TotalMinutes -lt 60) { 'vor {0:N0} Min.' -f $alter.TotalMinutes } elseif ($alter.TotalHours -lt 24) { 'vor {0:N0} Std.' -f $alter.TotalHours } else { 'vor {0:N0} Tagen' -f [Math]::Floor($alter.TotalDays) })
        if ($alter.TotalDays -ge 3) { $text += ' - bitte neu prüfen' }
    }
    $note.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, $farbe + 'Bg')
    (Get-Element 'NotePunkt').SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, $farbe)
    $t = Get-Element 'NoteText'
    $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $farbe)
    $t.Text = $text
    $note.Visibility = 'Visible'
}

# Beim Start: ist der letzte Stand aelter als 3 Tage (oder fehlt), still eine Schnellpruefung starten
function Start-AutoPruefung {
    if (-not $Script:Einstellungen.AutoPruefen -or $Script:Lauf) { return }
    $alter = Get-StandAlter
    if ($alter -and $alter.TotalDays -lt 3) { return }
    Start-Diagnose 'Schnell'
    if ($Script:Lauf) {
        (Get-Element 'OverlayTitel').Text = 'Automatische Prüfung …'
        (Get-Element 'OverlayText').Text = $(if ($alter) { 'Der letzte Stand ist {0:N0} Tage alt - die Datenquellen werden neu gelesen.' -f [Math]::Floor($alter.TotalDays) } else { 'Noch kein Stand vorhanden - die Datenquellen werden gelesen.' })
    }
}

# ============================================================================
#  Reinigung (Messung laden, Auswahl)
# ============================================================================

function ConvertTo-Kategorien {
    param($Daten)
    $kats = @($Daten.Kategorien | Where-Object { $_ })
    $max = 0.0
    foreach ($k in $kats) { if ($k.Messbar -and [double]$k.Bytes -gt $max) { $max = [double]$k.Bytes } }
    $liste = New-Object 'System.Collections.ObjectModel.ObservableCollection[SdKategorie]'
    foreach ($k in $kats) {
        $o = New-Object SdKategorie
        $o.Id = [string]$k.Id; $o.Gruppe = [string]$k.Gruppe; $o.Name = [string]$k.Name; $o.Beschreibung = [string]$k.Beschreibung
        $o.Hinweis = [string]$k.Hinweis; $o.Vorauswahl = [bool]$k.Vorauswahl; $o.Bytes = [int64]$k.Bytes; $o.Dateien = [int]$k.Dateien
        $o.Messbar = [bool]$k.Messbar; $o.Laeuft = [string]$k.Laeuft; $o.Sperre = [string]$k.Sperre; $o.Mindestens = [bool]$k.Mindestens
        $o.Anteil = if ($o.Messbar -and $max -gt 0) { [double]$o.Bytes / $max } else { 0.0 }
        $o.Ausgewaehlt = ($o.Vorauswahl -and -not $o.IstLeer)
        $liste.Add($o)
    }
    # Gruppensumme: nicht messbare Kategorien oder Mindestwerte machen die Summe zum Mindestwert
    foreach ($g in ($liste | Group-Object Gruppe)) {
        $mess = @($g.Group | Where-Object { $_.Messbar })
        $bytes = [int64](($mess | Measure-Object -Property Bytes -Sum).Sum)
        $text = if ($mess.Count -eq 0) { '?' }
                elseif (@($g.Group | Where-Object { -not $_.IstLeer }).Count -eq 0) { 'leer' }
                else { $(if (@($g.Group | Where-Object { $_.Mindestens -or -not $_.Messbar }).Count) { 'mind. ' } else { '' }) + [SdFormat]::Groesse($bytes) }
        foreach ($k in $g.Group) { $k.GruppeSumme = $text }
    }
    return , $liste
}

function Import-Reinigung {
    param([string] $Pfad)
    $d = Get-Content -LiteralPath $Pfad -Raw -Encoding UTF8 | ConvertFrom-Json
    $liste = ConvertTo-Kategorien $d
    $Script:Kategorien = $liste
    $Script:ReinigungStand = [string]$d.ZeitText
    $Script:ErwarteteKategorien = [Math]::Max(1, $liste.Count)
    $ansicht = [System.Windows.Data.CollectionViewSource]::GetDefaultView($liste)
    $ansicht.GroupDescriptions.Clear()
    [void]$ansicht.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription 'Gruppe'))
    $ansicht.Filter = [Predicate[object]] { param($k) Test-ReinigungSichtbar $k }
    (Get-Element 'ReinigungListe').ItemsSource = $ansicht

    $summe = [int64](($liste | Where-Object { $_.Messbar } | Measure-Object -Property Bytes -Sum).Sum)
    $offen = @($liste | Where-Object { -not $_.Messbar }).Count
    (Get-Element 'ReinigungSumme').Text = [SdFormat]::Groesse($summe)
    $mind = @($liste | Where-Object { $_.Mindestens }).Count
    (Get-Element 'ReinigungSummeText').Text = "in $($liste.Count) Kategorien" + $(if ($offen) { "  ·  $offen davon erst mit Adminrechten messbar (werden trotzdem gereinigt)" } else { '' }) +
        $(if ($mind) { "  ·  bei $mind ist die Größe ein Mindestwert (ohne Adminrechte nur teilweise lesbar)" } else { '' })
    (Get-Element 'ReinigungStand').Text = "Stand $($d.ZeitText)" + $(if ($d.Admin) { ', mit Adminrechten' } else { '' })
    (Get-Element 'ReiterReinigungText').Text = 'Reinigung  ·  ' + [SdFormat]::Groesse($summe)
    $Script:ReinigungFehler = [string]$d.Fehler
    Update-ReinigungsFilter
    Update-ReinigungsAuswahl
}

# Welche Kategorie zeigt die gewaehlte Ansicht? Leere nie (es gibt nichts zu tun).
function Test-ReinigungSichtbar {
    param($K)
    if (-not $K -or $K.IstLeer) { return $false }
    return ([bool]$K.Vorauswahl -eq ($Script:ReinigungFilter -ne 'Ansehen'))
}

function Update-ReinigungsFilter {
    $liste = @($Script:Kategorien | Where-Object { $_ -and -not $_.IstLeer })
    (Get-Element 'ZahlReinigungImmer').Text = [string]@($liste | Where-Object { $_.Vorauswahl }).Count
    (Get-Element 'ZahlReinigungAnsehen').Text = [string]@($liste | Where-Object { -not $_.Vorauswahl }).Count
    $quelle = (Get-Element 'ReinigungListe').ItemsSource
    if ($quelle -is [System.ComponentModel.ICollectionView]) { $quelle.Refresh() }
    $sichtbar = @($liste | Where-Object { Test-ReinigungSichtbar $_ }).Count
    $leer = ($sichtbar -eq 0)
    (Get-Element 'ReinigungLeer').Visibility = if ($leer) { 'Visible' } else { 'Collapsed' }
    (Get-Element 'ReinigungScroll').Visibility = if ($leer) { 'Collapsed' } else { 'Visible' }
    (Get-Element 'ReinigungEmpfohlenKnopf').IsEnabled = ($Script:ReinigungFilter -ne 'Ansehen')
    if ($leer) {
        (Get-Element 'ReinigungLeerText').Text = if ($Script:ReinigungFehler) { "Die Messung ist fehlgeschlagen: $Script:ReinigungFehler" }
            elseif (-not @($Script:Kategorien).Count) { 'Keine Kategorie gefunden.' }
            elseif ($Script:ReinigungFilter -eq 'Ansehen') { 'Hier liegt gerade nichts.' }
            else { 'Alles sauber - nichts mehr, was man bedenkenlos löschen kann.' }
    }
}

function Import-NeuesteReinigung {
    $f = Get-NeuesteEigeneDatei 'Reinigung-*.json'
    if ($f) { Import-Reinigung $f.FullName; return $true }
    return $false
}

function Update-ReinigungsAuswahl {
    $wahl = @($Script:Kategorien | Where-Object { $_ -and $_.Ausgewaehlt })
    $bytes = [int64](($wahl | Where-Object { $_.Messbar } | Measure-Object -Property Bytes -Sum).Sum)
    $ohne = @($wahl | Where-Object { -not $_.Messbar }).Count
    (Get-Element 'ReinigungAuswahlGross').Text = if ($wahl.Count) { [SdFormat]::Groesse($bytes) + $(if ($ohne) { ' +' } else { '' }) } else { '-' }
    (Get-Element 'ReinigungAuswahlText').Text = if ($wahl.Count) { "$($wahl.Count) Kategorie(n)" + $(if ($ohne) { ", $ohne ohne Messwert" } else { '' }) } else { 'Nichts ausgewählt' }
    (Get-Element 'ReinigenText').Text = if ($wahl.Count) { "Reinigen ($($wahl.Count))" } else { 'Reinigen' }
    (Get-Element 'ReinigenKnopf').IsEnabled = ($wahl.Count -gt 0) -and ($null -eq $Script:Lauf)
    $schliessen = @($wahl | Where-Object { $_.Laeuft } | ForEach-Object { $_.Laeuft -split ',\s*' } | ForEach-Object { (Get-ProgrammInfo $_)[0] } | Sort-Object -Unique)
    $t = Get-Element 'ReinigungSchliessenText'
    $t.Text = 'Vorher schließen: ' + ($schliessen -join ', ') + ' - sonst werden diese Kategorien übersprungen.'
    $t.Visibility = if ($schliessen.Count) { 'Visible' } else { 'Collapsed' }
}

# Ausgewaehlte Kategorien als Massnahmen fuer den Bereinigungsteil - nur die Kennung, nie ein Pfad
function Get-ReinigungsWahl {
    foreach ($k in @($Script:Kategorien | Where-Object { $_ -and $_.Ausgewaehlt })) {
        $x = New-Object SdLoesung
        $x.Id = 'Reinigen'; $x.Art = 'Bereinigung'; $x.Umkehrbar = $false
        $x.Titel = if ($k.Messbar) { "$($k.Name) ($($k.GroesseText))" } else { $k.Name }
        $x.ParameterJson = ConvertTo-Json -InputObject @{ Kategorie = $k.Id } -Compress
        $x.BefundNr = -1; $x.BefundTitel = "Reinigung  ·  $($k.Gruppe)"
        $x
    }
}

# ============================================================================
#  Laufende Programme - ihre Daten lassen sich erst nach dem Schliessen loeschen
# ============================================================================

# Prozessnamen je Kategorie aus Reinigung.psd1 - dieselbe Quelle wie in Bereinigung.ps1
function Get-KategorieProzesse {
    if (-not $Script:KategorieProzesse) {
        $Script:KategorieProzesse = @{}
        try {
            $def = Import-PowerShellDataFile -LiteralPath (Join-Path $Script:Ordner 'Reinigung.psd1')
            foreach ($k in @($def.Kategorien)) { $Script:KategorieProzesse[[string]$k.Id] = @(@($k.Prozesse) | Where-Object { [string]$_ -match '^[\w\-\.]{1,60}$' }) }
        } catch { }
    }
    return $Script:KategorieProzesse
}

# Anzeigename und Anleitung zum Schliessen; unbekannte Prozesse bekommen den allgemeinen Hinweis
function Get-ProgrammInfo {
    param([string] $Prozess)
    switch ($Prozess) {
        'chrome'  { return 'Google Chrome', 'Alle Chrome-Fenster schließen und im Infobereich (Pfeil unten rechts in der Taskleiste) Chrome beenden. Läuft es danach weiter: in Chrome unter Einstellungen > System „Apps im Hintergrund weiter ausführen“ ausschalten.' }
        'msedge'  { return 'Microsoft Edge', 'Alle Edge-Fenster schließen. Edge läuft oft unsichtbar weiter: in Edge unter Einstellungen > System und Leistung „Startboost“ und „Erweiterungen und Apps im Hintergrund weiter ausführen“ ausschalten.' }
        'firefox' { return 'Mozilla Firefox', 'Alle Firefox-Fenster schließen.' }
        'brave'   { return 'Brave', 'Alle Brave-Fenster schließen und im Infobereich Brave beenden.' }
        'discord' { return 'Discord', 'Rechtsklick auf das Discord-Symbol im Infobereich > „Discord beenden“ - das X am Fenster versteckt Discord nur.' }
        'blitz'   { return 'Blitz', 'Im Infobereich Blitz beenden. Normalerweise schließt es sich zusammen mit League of Legends von selbst.' }
        { $_ -in 'steam', 'steamwebhelper' } { return 'Steam', 'In Steam oben links „Steam“ > „Beenden“ - das X am Fenster schickt Steam nur in den Infobereich. Laufende Spiele und Downloads vorher beenden.' }
        default   { return $Prozess, 'Programm ganz schließen, auch im Infobereich (Pfeil unten rechts in der Taskleiste).' }
    }
}

# Prueft jetzt, welche Programme laufen: frischt die Hinweise „läuft gerade“ in der Liste auf und liefert
# fuer die Kategorien -Ids je laufendem Programm eine Zeile (gleiche Pruefung wie Bereinigung.ps1: Prozessname)
function Update-LaufendeProgramme {
    param([string[]] $Ids)
    if (-not $Script:Kategorien) { return }
    $prozesse = Get-KategorieProzesse
    $laufend = @{}
    foreach ($p in Get-Process) { $laufend[$p.ProcessName] = $true }
    $geaendert = $false
    foreach ($k in @($Script:Kategorien | Where-Object { $_ })) {
        $text = (@($prozesse[$k.Id]) | Where-Object { $_ -and $laufend.ContainsKey($_) }) -join ', '
        if ([string]$k.Laeuft -ne $text) { $k.Laeuft = $text; $geaendert = $true }
    }
    if ($geaendert) { [System.Windows.Data.CollectionViewSource]::GetDefaultView($Script:Kategorien).Refresh(); Update-ReinigungsAuswahl }
    $programme = [ordered]@{}
    foreach ($k in @($Script:Kategorien | Where-Object { $_ -and $_.Laeuft -and $Ids -contains $_.Id })) {
        $titel = if ($k.Messbar -and -not $k.IstLeer) { "$($k.Name) ($($k.GroesseText))" } else { $k.Name }
        foreach ($p in ($k.Laeuft -split ',\s*')) {
            $name, $anleitung = Get-ProgrammInfo $p
            if (-not $programme.Contains($name)) { $x = New-Object SdProgramm; $x.Programm = $name; $x.Anleitung = $anleitung; $programme[$name] = $x }
            if (-not $programme[$name].Kategorien.Contains($titel)) { $programme[$name].Kategorien.Add($titel) }
        }
    }
    return @($programme.Values)
}

# Vor dem Reinigen: laufen Programme, deren Daten geloescht werden sollen, kommt erst der Hinweis „Programme schließen“.
# Gibt die Auswahl zurueck - bei „Ohne diese reinigen“ ohne die blockierten Kategorien, bei Abbruch nichts.
function Confirm-ProgrammeGeschlossen {
    param([object[]] $Wahl)
    $ids = @($Wahl | Where-Object { $_.Id -eq 'Reinigen' } | ForEach-Object { [string]($_.ParameterJson | ConvertFrom-Json).Kategorie })
    $laufend = @()
    if ($ids.Count) { $laufend = @(Update-LaufendeProgramme $ids) }
    if (-not $laufend.Count) { return $Wahl }
    $Script:ProgrammIds = $ids
    $blockiert = @($Script:Kategorien | Where-Object { $_ -and $_.Laeuft -and $ids -contains $_.Id })
    $antwort = Show-ProgrammeSchliessen $laufend ($blockiert.Count -lt $Wahl.Count)
    if ($antwort -eq 'Weiter') { return $Wahl }
    if ($antwort -ne 'Ohne') { return }
    $null = Update-LaufendeProgramme $ids
    $blockiert = @($Script:Kategorien | Where-Object { $_ -and $_.Laeuft -and $ids -contains $_.Id } | ForEach-Object { $_.Id })
    return @($Wahl | Where-Object { -not ($_.Id -eq 'Reinigen' -and $blockiert -contains [string]($_.ParameterJson | ConvertFrom-Json).Kategorie) })
}

# Auf Knopfdruck und alle 1,5 s von selbst: ist alles geschlossen, geht es weiter zur Bestaetigung
function Invoke-ProgrammPruefung {
    param([switch] $Still)
    $f = $Script:ProgrammFenster
    $noch = @(Update-LaufendeProgramme $Script:ProgrammIds)
    if (-not $noch.Count) { $Script:DialogAntwort = 'Weiter'; $f.Close(); return }
    $signatur = (@($noch | ForEach-Object { $_.Programm + ':' + ($_.Kategorien -join '|') }) -join ';')
    if ($signatur -ne $Script:ProgrammSignatur) { $f.FindName('Liste').ItemsSource = $noch; $Script:ProgrammSignatur = $signatur }
    if (-not $Still) { $f.FindName('Status').Text = 'Läuft noch (geprüft ' + (Get-Date).ToString('HH:mm:ss') + ') - sobald alles geschlossen ist, geht es von selbst weiter.' }
}

function Show-ProgrammeSchliessen {
    param([object[]] $Programme, [bool] $KannOhne)
    $d = New-Dialog $Script:ProgrammeXaml
    $d.FindName('Liste').ItemsSource = $Programme
    $d.FindName('Ohne').Visibility = if ($KannOhne) { 'Visible' } else { 'Collapsed' }
    $Script:DialogAntwort = 'Abbrechen'
    $Script:ProgrammFenster = $d
    $Script:ProgrammSignatur = ''
    $Script:ProgrammUhr = New-Object System.Windows.Threading.DispatcherTimer
    $Script:ProgrammUhr.Interval = [TimeSpan]::FromMilliseconds(1500)
    $Script:ProgrammUhr.Add_Tick({ try { Invoke-ProgrammPruefung -Still } catch { $Script:ProgrammUhr.Stop() } })
    $d.Add_Closed({ $Script:ProgrammUhr.Stop() })
    $d.FindName('Pruefen').Add_Click({ Invoke-ProgrammPruefung })
    $d.FindName('Ohne').Add_Click({ param($s, $e) $Script:DialogAntwort = 'Ohne'; [System.Windows.Window]::GetWindow($s).Close() })
    $d.FindName('Abbrechen').Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).Close() })
    if ($Vorschau) { return $d }
    $Script:ProgrammUhr.Start()
    [void]$d.ShowDialog()
    return $Script:DialogAntwort
}

# ============================================================================
#  Laeufe (Diagnose, Bereinigung, Rueckgaengig) - ohne das Fenster zu blockieren
# ============================================================================

function Show-Overlay {
    param([string] $Titel, [string] $Text, [bool] $Unbestimmt, [string] $Icon = [string][char]0xE9D9)
    (Get-Element 'OverlayTitel').Text = $Titel
    (Get-Element 'OverlayText').Text = $Text
    (Get-Element 'OverlayIcon').Text = $Icon
    $f = Get-Element 'OverlayFortschritt'
    $f.IsIndeterminate = $Unbestimmt
    $f.Value = 0
    (Get-Element 'OverlayProzent').Text = $(if ($Unbestimmt) { '' } else { '0 %' })
    (Get-Element 'OverlayAktiv').IsIndeterminate = $true
    (Get-Element 'OverlayAktivText').Text = 'arbeitet'
    (Get-Element 'OverlayZeit').Text = '00:00'
    $abbrechen = Get-Element 'OverlayAbbrechen'
    $abbrechen.Visibility = $(if ($Script:Lauf -and $Script:Lauf.Art -in 'Diagnose', 'Analyse') { 'Visible' } else { 'Collapsed' })
    $abbrechen.IsEnabled = $true
    (Get-Element 'LaufMini').Visibility = 'Collapsed'
    (Get-Element 'Overlay').Visibility = 'Visible'
    foreach ($n in 'StartKnopf', 'BereinigenKnopf', 'ReinigenKnopf', 'AnalysierenKnopf', 'AnalysierenLeerKnopf') { (Get-Element $n).IsEnabled = $false }
}

# Minimieren / wiederherstellen: nur die Anzeige, der Lauf und die gesperrten Start-Knoepfe bleiben, wie sie sind
function Set-OverlayKlein {
    param([bool] $Klein)
    if ($Klein -and -not $Script:Lauf) { return }
    (Get-Element 'Overlay').Visibility = $(if ($Klein) { 'Collapsed' } else { 'Visible' })
    (Get-Element 'LaufMini').Visibility = $(if ($Klein) { 'Visible' } else { 'Collapsed' })
    Set-StartKnopfVerdeckt $Klein
    if ($Klein) { Update-LaufMini }
}

# Hidden statt Collapsed: Breite bleibt fuer Chip und Schalter. ClearValue gibt die Sichtbarkeit wieder an den Stil (nur im Reiter Befunde)
function Set-StartKnopfVerdeckt {
    param([bool] $Verdeckt)
    $k = Get-Element 'StartKnopf'
    if ($Verdeckt) { $k.Visibility = 'Hidden' } else { $k.ClearValue([System.Windows.UIElement]::VisibilityProperty) }
}

function Update-LaufMini {
    if ((Get-Element 'LaufMini').Visibility -ne 'Visible') { return }
    $f = Get-Element 'OverlayFortschritt'
    $m = Get-Element 'LaufMiniBalken'
    $m.IsIndeterminate = $f.IsIndeterminate
    $m.Value = $f.Value
    $titel = ((Get-Element 'OverlayTitel').Text -replace '\s*…$', '' -replace '^Prüfung läuft - ', '' -replace ' läuft$', '')
    $prozent = (Get-Element 'OverlayProzent').Text
    (Get-Element 'LaufMiniText').Text = $titel + $(if ($prozent) { ' · ' + $prozent } else { '' })
}

function Hide-Overlay {
    (Get-Element 'LaufMini').Visibility = 'Collapsed'
    Set-StartKnopfVerdeckt $false
    (Get-Element 'Overlay').Visibility = 'Collapsed'
    foreach ($n in 'StartKnopf', 'AnalysierenKnopf', 'AnalysierenLeerKnopf') { (Get-Element $n).IsEnabled = $true }
    Update-Auswahl
    Update-ReinigungsAuswahl
}

function Get-GewaehlterModus {
    if ((Get-Element 'ModusSchalter').IsChecked) { return 'Gruendlich' }
    return 'Schnell'
}

function Start-Diagnose {
    # -Vorher (aus Get-Vorher): danach mit diesem Stand vergleichen
    param([string] $Modus, $Vorher)
    if ($Script:Lauf -or $Script:NurLesen) { return }
    $argumente = "-NoProfile -ExecutionPolicy Bypass -File `"$Script:Engine`" -Modus $Modus -KeinBrowser"
    $lauf = @{ Art = 'Diagnose'; Modus = $Modus; Start = (Get-Date); Log = $null; StatusDatei = $null; Vorher = $Vorher; Gewichte = (Get-QuellenGewichte $Modus) }
    $lauf.Log = Join-Path $env:TEMP ('Systemdiagnose-' + [guid]::NewGuid().ToString('N') + '.log')
    try {
        if ($Modus -eq 'Schnell') {
            $lauf.Prozess = Start-Process -FilePath 'powershell.exe' -ArgumentList $argumente -WindowStyle Hidden -RedirectStandardOutput $lauf.Log -PassThru
        } else {
            # Mit Adminrechten laesst sich die Ausgabe nicht umleiten - ein Protokoll (Transcript) des erhoehten Prozesses
            # liefert dieselben Zeilen fuer den Fortschritt. Die Engine selbst schreibt weiterhin nur in Berichte\.
            # Abbrechen: die App darf den erhoehten Prozess nicht beenden - ein Waechter-Thread darin beendet ihn samt DISM/SFC,
            # sobald die App die Datei <Log>.abbruch anlegt. Als -EncodedCommand, damit keine Anfuehrungszeichen verschachtelt werden.
            $befehl = @"
`$w = [powershell]::Create().AddScript({ param(`$f, `$p) while (-not (Test-Path -LiteralPath `$f)) { Start-Sleep -Milliseconds 400 }; taskkill.exe /T /F /PID `$p | Out-Null }).AddArgument('$($lauf.Log).abbruch').AddArgument(`$PID); [void]`$w.BeginInvoke()
Start-Transcript -LiteralPath '$($lauf.Log)' | Out-Null; & '$Script:Engine' -Modus $Modus -KeinBrowser; Stop-Transcript | Out-Null
"@
            $kodiert = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($befehl))
            $lauf.Prozess = Start-Process -FilePath 'powershell.exe' -ArgumentList "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $kodiert" -Verb RunAs -WindowStyle Hidden -PassThru
        }
    } catch {
        Show-Hinweis 'Prüfung nicht gestartet' "Windows hat die Prüfung nicht gestartet - vermutlich wurden die Adminrechte nicht erteilt.`n`n$($_.Exception.Message)"
        return
    }
    $Script:Lauf = $lauf
    $text = if ($Modus -eq 'Schnell') { 'Die Datenquellen werden gelesen …' } else { 'Läuft mit Adminrechten im Hintergrund: Systemdateien, Komponentenspeicher, Dateisystem jedes Laufwerks und Signaturen (nur lesen) - ca. 5-10 Minuten. Der PC bleibt normal benutzbar.' }
    $titel = 'Prüfung läuft - ' + $(if ($Modus -eq 'Gruendlich') { 'Gründlich' } else { $Modus })
    Show-Overlay $titel $text ($Modus -ne 'Schnell')
    $Script:Uhr.Start()
}

# Pruefung/Analyse abbrechen (Knopf oder Esc). Bereinigung/Rueckgaengig bewusst nicht: halb ausgefuehrte Aenderungen waeren schlimmer.
function Stop-Lauf {
    $l = $Script:Lauf
    if (-not $l -or $l.Abgebrochen -or $l.Art -notin 'Diagnose', 'Analyse') { return }
    $l.Abgebrochen = $true
    (Get-Element 'OverlayAbbrechen').IsEnabled = $false
    (Get-Element 'OverlayText').Text = 'Wird abgebrochen …'
    if ($l.Art -eq 'Diagnose' -and $l.Modus -ne 'Schnell') { [IO.File]::WriteAllText($l.Log + '.abbruch', '') }
    else { & taskkill.exe /T /F /PID $l.Prozess.Id 2>&1 | Out-Null }
}

function Start-Analyse {
    param($Vorher)
    if ($Script:Lauf -or $Script:NurLesen) { return }
    $log = Join-Path $env:TEMP ('Systemdiagnose-' + [guid]::NewGuid().ToString('N') + '.log')
    $argumente = "-NoProfile -ExecutionPolicy Bypass -File `"$Script:Engine`" -Modus Reinigung -KeinBrowser"
    try { $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $argumente -WindowStyle Hidden -RedirectStandardOutput $log -PassThru }
    catch { Show-Hinweis 'Analyse nicht gestartet' $_.Exception.Message; return }
    $Script:Lauf = @{ Art = 'Analyse'; Start = (Get-Date); Prozess = $p; Log = $log; StatusDatei = $null; Vorher = $Vorher }
    Show-Overlay 'Analyse läuft …' 'Die Reinigungskategorien werden gemessen …' $true ([string][char]0xE721)
    $Script:Uhr.Start()
}

function Start-Bereinigung {
    param([object[]] $Wahl, [string] $Quelle = 'Befunde')
    $wahl = @($Wahl | Where-Object { $_ })
    if ($wahl.Count -eq 0 -or $Script:Lauf -or $Script:NurLesen) { return }
    $wahl = @(Confirm-ProgrammeGeschlossen $wahl | Where-Object { $_ })
    if ($wahl.Count -eq 0) { return }
    if (-not (Show-Bestaetigung $wahl)) { return }
    if (-not (Test-Path -LiteralPath $Script:BereinigungsOrdner)) { [void](New-Item -ItemType Directory -Path $Script:BereinigungsOrdner) }
    $stempel = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $plan = [ordered]@{
        Format = 'Systemdiagnose-Plan/1'; Erstellt = (Get-Date).ToString('s')
        BenutzerSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value; Diagnose = $Script:ErgebnisPfad
        Aktionen = @($wahl | ForEach-Object { [ordered]@{ Id = $_.Id; Titel = $_.Titel; BefundNr = $_.BefundNr; BefundTitel = $_.BefundTitel; Parameter = ($_.ParameterJson | ConvertFrom-Json) } })
    }
    $planPfad = Join-Path $Script:BereinigungsOrdner "Plan-$stempel.json"
    [IO.File]::WriteAllText($planPfad, ($plan | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($true)))
    $argumente = "-NoProfile -ExecutionPolicy Bypass -File `"$Script:Bereiniger`" -Plan `"$planPfad`""
    try { $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $argumente -Verb RunAs -WindowStyle Hidden -PassThru }
    catch { Show-Hinweis 'Nicht gestartet' 'Die Adminrechte wurden nicht erteilt - es wurde nichts verändert.'; return }
    $Script:Lauf = @{ Art = 'Bereinigung'; Quelle = $Quelle; Start = (Get-Date); Prozess = $p; Log = $null; Wahl = $wahl
        StatusDatei = (Join-Path $Script:BereinigungsOrdner "Status-$stempel.txt"); ErgebnisDatei = (Join-Path $Script:BereinigungsOrdner "Ergebnis-$stempel.json") }
    if ($Quelle -eq 'Reinigung') { Show-Overlay 'Reinigung läuft …' 'Die ausgewählten Kategorien werden gereinigt.' $true ([string][char]0xE74D) }
    else { Show-Overlay 'Wird behoben …' 'Jede Maßnahme wird vor der Ausführung erneut geprüft.' $true ([string][char]0xE90F) }
    $Script:Uhr.Start()
}

function Start-Rueckgaengig {
    param([string] $ErgebnisDatei)
    if ($Script:Lauf -or $Script:NurLesen -or -not (Test-Path -LiteralPath $ErgebnisDatei)) { return }
    $stempel = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $argumente = "-NoProfile -ExecutionPolicy Bypass -File `"$Script:Bereiniger`" -Rueckgaengig `"$ErgebnisDatei`" -Stempel $stempel"
    try { $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $argumente -Verb RunAs -WindowStyle Hidden -PassThru }
    catch { Show-Hinweis 'Nicht gestartet' 'Die Adminrechte wurden nicht erteilt - es wurde nichts verändert.'; return }
    $Script:Lauf = @{ Art = 'Rueckgaengig'; Start = (Get-Date); Prozess = $p; Log = $null
        StatusDatei = (Join-Path $Script:BereinigungsOrdner "Status-$stempel.txt"); ErgebnisDatei = (Join-Path $Script:BereinigungsOrdner "Rueckgaengig-$stempel.json") }
    Show-Overlay 'Wird rückgängig gemacht …' 'Die gespeicherten alten Zustände werden wiederhergestellt.' $true ([string][char]0xE7A7)
    $Script:Uhr.Start()
}

function Set-OverlayProzent {
    param([double] $Wert)
    $w = [Math]::Max(0, [Math]::Min(99, $Wert))
    $f = Get-Element 'OverlayFortschritt'
    $f.IsIndeterminate = $false
    if ($w -gt $f.Value) { $f.Value = $w }
    (Get-Element 'OverlayProzent').Text = ('{0:N0} %' -f $f.Value)
}

# Gewichte fuer den Prozentbalken: Dauer jeder Datenquelle im letzten Lauf derselben Stufe (Tief zaehlt als Gruendlich)
function Get-QuellenGewichte {
    param([string] $Modus)
    foreach ($f in @(Get-ChildItem -LiteralPath $Script:BerichtOrdner -Filter 'Befunde-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 15)) {
        try { $j = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }
        $m = if ($j.Modus -in 'Tief', 'Gruendlich') { 'Gruendlich' } else { [string]$j.Modus }
        if ($m -ne $Modus) { continue }
        $sek = @{}; $reihe = New-Object System.Collections.Generic.List[string]; $summe = 0.0
        foreach ($q in @($j.Quellen | Where-Object { [string]$_.Status -notmatch '^übersprungen' })) {
            $n = [string]$q.Datenquelle; $s = [Math]::Max(0.2, [double]$q.Sekunden)
            $sek[$n] = $s; $reihe.Add($n); $summe += $s
        }
        if ($reihe.Count) { return @{ Sekunden = $sek; Reihenfolge = $reihe.ToArray(); Summe = $summe } }
    }
    return $null
}

# Lebenszeichen: CPU-Zeit des Pruefprozesses samt Kindprozessen (DISM, DismHost, SFC ...). Win32_Process liefert sie auch
# fuer Prozesse mit Adminrechten. Nur alle 2 s abgefragt (~140 ms), und nur, waehrend ein Lauf aktiv ist.
function Update-Aktivitaet {
    param($Lauf)
    $jetzt = Get-Date
    if (-not $Lauf.LetzteAktivitaet) { $Lauf.LetzteAktivitaet = $jetzt; $Lauf.CpuStand = -1; $Lauf.LogStand = -1; $Lauf.Abfrage = [datetime]::MinValue }
    if (($jetzt - $Lauf.Abfrage).TotalSeconds -ge 2) {
        $Lauf.Abfrage = $jetzt
        $cpu = [int64]0
        try {
            $alle = @(Get-CimInstance -ClassName Win32_Process -Property Name, ProcessId, ParentProcessId, KernelModeTime, UserModeTime -ErrorAction Stop)
            $baum = @{ ([int]$Lauf.Prozess.Id) = $true }
            # SFC und DISM lassen die eigentliche Arbeit vom Windows-Moduldienst erledigen (kein Kindprozess) - der zaehlt mit.
            # WmiPrvSE bewusst nicht: den stoesst diese Abfrage selbst an, er waere immer „aktiv“.
            foreach ($p in $alle) { if ([string]$p.Name -in 'TiWorker.exe', 'TrustedInstaller.exe') { $baum[[int]$p.ProcessId] = $true } }
            do {
                $neu = 0
                foreach ($p in $alle) { if (-not $baum.ContainsKey([int]$p.ProcessId) -and $baum.ContainsKey([int]$p.ParentProcessId)) { $baum[[int]$p.ProcessId] = $true; $neu++ } }
            } while ($neu)
            foreach ($p in $alle) { if ($baum.ContainsKey([int]$p.ProcessId)) { $cpu += [int64]$p.KernelModeTime + [int64]$p.UserModeTime } }
        } catch { $cpu = -2 }
        $log = if ($Lauf.Log -and (Test-Path -LiteralPath $Lauf.Log)) { (Get-Item -LiteralPath $Lauf.Log).Length } else { 0 }
        if ($cpu -ne $Lauf.CpuStand -or $log -ne $Lauf.LogStand -or $cpu -eq -2) { $Lauf.LetzteAktivitaet = $jetzt }
        $Lauf.CpuStand = $cpu; $Lauf.LogStand = $log
    }
    $still = ($jetzt - $Lauf.LetzteAktivitaet).TotalSeconds
    $balken = Get-Element 'OverlayAktiv'
    if ($still -lt 6) {
        if (-not $balken.IsIndeterminate) { $balken.IsIndeterminate = $true }
        (Get-Element 'OverlayAktivText').Text = 'arbeitet'
    } else {
        if ($balken.IsIndeterminate) { $balken.IsIndeterminate = $false; $balken.Value = 0 }
        (Get-Element 'OverlayAktivText').Text = ('keine Aktivität seit {0:N0} s' -f $still) + $(if ($still -ge 120) { ' - evtl. hängt der Prozess' } else { '' })
    }
}

function Update-Lauf {
    $l = $Script:Lauf
    if (-not $l) { $Script:Uhr.Stop(); return }
    $dauer = (Get-Date) - $l.Start
    (Get-Element 'OverlayZeit').Text = ('{0:00}:{1:00}' -f [int][Math]::Floor($dauer.TotalMinutes), $dauer.Seconds)
    if ($l.Log -and (Test-Path -LiteralPath $l.Log) -and $l.Art -eq 'Analyse') {
        $kat = [regex]::Matches((Read-GeteilteDatei $l.Log 65001), '(?m)^\s{6}> (.+?)\s*$')
        if ($kat.Count) {
            (Get-Element 'OverlayText').Text = 'Gemessen: ' + $kat[$kat.Count - 1].Groups[1].Value
            Set-OverlayProzent ($kat.Count / $(if ($Script:ErwarteteKategorien) { $Script:ErwarteteKategorien } else { 18 }) * 100)
        }
    } elseif ($l.Log -and (Test-Path -LiteralPath $l.Log)) {
        $text = Read-GeteilteDatei $l.Log 65001
        $treffer = [regex]::Matches($text, '(?m)^\s{4}(\S.*?)\s{2,}[\d\.,]+ s')
        if ($treffer.Count) {
            $fertigListe = @($treffer | ForEach-Object { $_.Groups[1].Value.Trim() })
            $g = $l.Gewichte
            if ($g -and $g.Summe -gt 0) {
                # Prozent nach der Dauer jeder Quelle im letzten Lauf derselben Stufe - DISM/SFC zaehlen so viel, wie sie dauern
                $erledigt = 0.0
                foreach ($n in $fertigListe) { if ($g.Sekunden.ContainsKey($n)) { $erledigt += $g.Sekunden[$n] } }
                Set-OverlayProzent ($erledigt / $g.Summe * 100)
                $naechste = @($g.Reihenfolge | Where-Object { $fertigListe -notcontains $_ } | Select-Object -First 1)
                (Get-Element 'OverlayText').Text = if ($naechste.Count) { 'Liest: ' + $naechste[0] } else { 'Gelesen: ' + $fertigListe[-1] }
            } else {
                Set-OverlayProzent ($treffer.Count / $(if ($l.Modus -eq 'Gruendlich') { 75 } else { 55 }) * 100)
                (Get-Element 'OverlayText').Text = 'Gelesen: ' + $fertigListe[-1]
            }
        }
    }
    if ($l.StatusDatei -and (Test-Path -LiteralPath $l.StatusDatei)) {
        $s = (Read-GeteilteDatei $l.StatusDatei).Trim()
        if ($s) {
            (Get-Element 'OverlayText').Text = $s
            if ($s -match '^\((\d+)/(\d+)\)' -and [int]$Matches[2] -gt 0) { Set-OverlayProzent (([int]$Matches[1] - 1) / [int]$Matches[2] * 100) }
        }
    }
    Update-Aktivitaet $l
    $fertig = $true
    try { $fertig = $l.Prozess.HasExited } catch { $fertig = $true }
    if (-not $fertig) { return }
    $Script:Uhr.Stop()
    $Script:Lauf = $null
    Hide-Overlay
    if ($l.Abgebrochen) {
        # Abgebrochen: nichts einlesen, bisherige Ergebnisse bleiben stehen
        foreach ($f in $l.Log, ($l.Log + '.abbruch')) { if ($f) { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue } }
        return
    }
    if ($l.Art -eq 'Diagnose') { Complete-Diagnose $l }
    elseif ($l.Art -eq 'Analyse') { Complete-Analyse $l }
    else { Complete-Bereinigung $l }
    Update-FehlerPunkt
}

function Complete-Analyse {
    param($Lauf)
    $text = Read-GeteilteDatei $Lauf.Log 65001
    Remove-Item -LiteralPath $Lauf.Log -Force -ErrorAction SilentlyContinue
    $m = [regex]::Match($text, 'Reinigungsdatei:\s*(.+?\.json)')
    if ($m.Success -and (Test-Path -LiteralPath $m.Groups[1].Value.Trim())) {
        Import-Reinigung $m.Groups[1].Value.Trim()
        (Get-Element 'Reiter').SelectedItem = Get-Element 'ReiterReinigung'
        if ($Lauf.Vorher) { Show-Vergleich $Lauf.Vorher -NurReinigung }
        # Die Kategorien sind jetzt frisch gemessen - Befund-Massnahmen warten weiter auf eine Pruefung
        foreach ($k in @($Script:Ausgefuehrt.Keys | Where-Object { $_ -like 'K:*' })) { $Script:Ausgefuehrt.Remove($k) }
    } else {
        $ende = (($text -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 6) -join "`n")
        Write-Fehlerlog ("Analyse ohne Messdatei`r`n" + (($text -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 40) -join "`r`n"))
        Show-Hinweis 'Keine Messung' ("Die Analyse hat keine Messdatei geschrieben (Abbruch oder Sicherheitsprüfung fehlgeschlagen).`n`n" + $ende)
    }
}

function Complete-Diagnose {
    param($Lauf)
    $json = $null
    $logText = ''
    if ($Lauf.Log -and (Test-Path -LiteralPath $Lauf.Log)) {
        $logText = Read-GeteilteDatei $Lauf.Log 65001
        $m = [regex]::Match($logText, 'Befunddatei:\s*(.+?\.json)')
        if ($m.Success) { $json = $m.Groups[1].Value.Trim() }
        Remove-Item -LiteralPath $Lauf.Log -Force -ErrorAction SilentlyContinue
    }
    if (-not $json) {
        $neu = Get-ChildItem -LiteralPath $Script:BerichtOrdner -Filter 'Befunde-*.json' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -ge $Lauf.Start.AddSeconds(-5) } | Sort-Object Name -Descending | Select-Object -First 1
        if ($neu) { $json = $neu.FullName }
    }
    if ($json -and (Test-Path -LiteralPath $json)) {
        Import-Ergebnis $json -Behalten:([bool]$Lauf.Vorher)
        try { [void](Import-NeuesteReinigung) } catch { }
        # Beim Aktualisieren bleibt der Reiter, wo er war
        if ($Lauf.Vorher) { Show-Vergleich $Lauf.Vorher } else { (Get-Element 'Reiter').SelectedIndex = 0 }
        Show-ProfilHinweis
        $Script:Ausgefuehrt = @{}
    } else {
        $ende = if ($logText) { "`n`nLetzte Ausgabe:`n" + (($logText -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 6) -join "`n") } else { '' }
        Write-Fehlerlog ("Prüfung ($($Lauf.Modus)) ohne Befunddatei`r`n" + (($logText -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 40) -join "`r`n"))
        Show-Hinweis 'Keine Ergebnisse' ("Die Prüfung hat keine Befunddatei geschrieben (Abbruch oder Sicherheitsprüfung fehlgeschlagen)." + $ende)
    }
}

function Complete-Bereinigung {
    param($Lauf)
    $datei = $Lauf.ErgebnisDatei
    if (-not (Test-Path -LiteralPath $datei)) {
        Write-Fehlerlog "Bereinigung ($($Lauf.Art)) ohne Ergebnisdatei: $datei"
        Show-Hinweis 'Kein Ergebnis' 'Es wurde kein Ergebnis geschrieben. Details stehen in Bereinigung\Protokoll.txt.'
        return
    }
    $e = Get-Content -LiteralPath $datei -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($Lauf.Quelle -ne 'Reinigung') { foreach ($l in @(Get-AlleLoesungen)) { $l.Ausgewaehlt = $false } }
    Add-Ausgefuehrt $Lauf.Wahl $e
    Update-Verlauf
    # Lief alles glatt, braucht es keinen eigenen Ergebnisdialog: der Vergleich nach dem automatischen Neu-Pruefen
    # nennt jede Massnahme mit ihrer Meldung. Nur bei Fehlern, Abgelehntem oder Simulation erst das Ergebnis zeigen
    # („übersprungen“ = leer, gesperrt oder Programm laeuft - das zeigt der Vergleich ohnehin).
    $zeilen = @($e.Aktionen | Where-Object { $_ })
    $erledigt = @($zeilen | Where-Object { $_.Status -eq 'erledigt' }).Count
    $glatt = -not $e.Fehler -and -not $e.Simulation -and $erledigt -gt 0 -and -not @($zeilen | Where-Object { $_.Status -in 'fehlgeschlagen', 'abgelehnt' }).Count
    if (-not $glatt) {
        if ((Show-ErgebnisDialog $e $datei) -eq 'Rueckgaengig') { Start-Rueckgaengig $datei; return }
    }
    # Automatisch neu messen bzw. pruefen, sobald sich etwas geaendert hat - damit nichts Geloeschtes stehen bleibt
    if ($erledigt -and -not $e.Simulation) {
        if (Test-IstReinigung $e) {
            Start-Analyse (Get-Vorher)
            if ($Script:Lauf) { (Get-Element 'OverlayTitel').Text = 'Wird neu gemessen …' }
        } else {
            # Immer Schnell: keine zweite Adminabfrage direkt nach der Behebung (Johannes, 28.10.)
            Start-Diagnose 'Schnell' (Get-Vorher)
            if ($Script:Lauf) { (Get-Element 'OverlayTitel').Text = 'Wird neu geprüft …' }
        }
    }
}

# Merkt sich, was ausgefuehrt wurde - der naechste Vergleich nennt dazu die Meldung des Bereinigungsteils
function Add-Ausgefuehrt {
    param([object[]] $Wahl, $Ergebnis)
    if ($Ergebnis.Simulation) { return }
    foreach ($a in @($Ergebnis.Aktionen | Where-Object { $_ })) {
        $w = @($Wahl | Where-Object { $_ -and $_.Titel -eq [string]$a.Titel -and $_.BefundNr -eq [int]$a.BefundNr })[0]
        $k = $null
        if (-not $w) {
            # Rueckgaengig: keine Auswahl dazu, aber das System hat sich veraendert
            if ([string]$a.Status -eq 'erledigt') { $k = 'R:' + [string]$a.Titel }
        } elseif ($w.Id -eq 'Reinigen') {
            $k = 'K:' + [string]($w.ParameterJson | ConvertFrom-Json).Kategorie
        } else {
            $b = @($Script:Daten.Befunde | Where-Object { $_ -and [int]$_.Nr -eq $w.BefundNr })[0]
            if ($b) { $k = "B:$($b.Typ)|$($b.Pruefung)" }
        }
        if (-not $k) { continue }
        # Mehrere Massnahmen zu einem Befund: eine nicht erledigte sagt mehr als eine erledigte
        $bisher = $Script:Ausgefuehrt[$k]
        if ($bisher -and $bisher.Status -ne 'erledigt' -and [string]$a.Status -eq 'erledigt') { continue }
        $Script:Ausgefuehrt[$k] = [pscustomobject]@{ Status = [string]$a.Status; Meldung = [string]$a.Meldung }
    }
}

# Ein Lauf, der nur aus Reinigungs-Kategorien besteht
function Test-IstReinigung {
    param($Ergebnis)
    $a = @($Ergebnis.Aktionen | Where-Object { $_ })
    return ($a.Count -gt 0) -and (@($a | Where-Object { $_.Id -ne 'Reinigen' }).Count -eq 0)
}

function Get-FreigegebenBytes {
    param($Ergebnis)
    return [int64]((@($Ergebnis.Aktionen | Where-Object { $_ -and $_.FreigegebenBytes -gt 0 }) | Measure-Object -Property FreigegebenBytes -Sum).Sum)
}

# ============================================================================
#  Dialoge: Bestaetigung und Ergebnis
# ============================================================================

function Show-Bestaetigung {
    param([object[]] $Wahl)
    $d = New-Dialog $Script:BestaetigungXaml
    $nurReinigung = (@($Wahl | Where-Object { $_.Id -ne 'Reinigen' }).Count -eq 0)
    $d.FindName('Titel').Text = if ($nurReinigung) { "$($Wahl.Count) Kategorie(n) reinigen?" } else { "$($Wahl.Count) Maßnahme(n) ausführen?" }
    if ($nurReinigung) {
        $d.FindName('Icon').Text = [string][char]0xE74D
        $d.FindName('Einleitung').Text = 'Nach „Jetzt ausführen“ wird mit Adminrechten gelöscht - nicht umkehrbar. Unmittelbar vor jeder Kategorie wird erneut gemessen und geprüft, ob ihr Programm läuft. Die Größen unten stammen aus der letzten Analyse.'
    }
    $d.FindName('Liste').ItemsSource = $Wahl
    $hinweise = New-Object System.Collections.Generic.List[string]
    if (@($Wahl | Where-Object { $_.Id -eq 'Reinigen' }).Count) { $hinweise.Add('• Gelöscht wird nur in den fest hinterlegten Ordnern der Kategorie, Datei für Datei. Gerade benutzte Dateien bleiben liegen; läuft ein zugehöriges Programm (z. B. Chrome), wird die Kategorie übersprungen.') }
    if (@($Wahl | Where-Object { $_.Id -eq 'Reinigen' -and $_.ParameterJson -match '"Papierkorb"' }).Count) { $hinweise.Add('• Der Papierkorb wird endgültig geleert - auch verwaiste Reste, die Windows selbst nicht mehr anzeigt.') }
    if (@($Wahl | Where-Object { $_.Id -eq 'TreiberpaketEntfernen' }).Count) { $hinweise.Add('• Vor dem Entfernen von Treiberpaketen wird ein Wiederherstellungspunkt angelegt. Ist das nicht möglich (Windows erlaubt nur einen pro 24 Stunden), wird ein höchstens einen Tag alter vorausgesetzt - sonst bleiben die Pakete unangetastet.') }
    if (@($Wahl | Where-Object { $_.Umkehrbar }).Count) { $hinweise.Add('• Umkehrbare Maßnahmen lassen sich danach unter „Verlauf“ mit einem Klick zurücknehmen.') }
    if (@($Wahl | Where-Object { -not $_.Umkehrbar -and $_.Id -ne 'Reinigen' }).Count) { $hinweise.Add('• Nicht umkehrbare Maßnahmen löschen Dateien oder Treiberpakete, die nicht mehr gebraucht werden.') }
    if (@($Wahl | Where-Object { $_.HatDauer }).Count) { $hinweise.Add('• Einige Maßnahmen dauern länger - der PC bleibt währenddessen normal benutzbar.') }
    $d.FindName('Hinweise').Text = (@($hinweise) -join "`n")
    $Script:DialogAntwort = $false
    $d.FindName('Ausfuehren').Add_Click({ param($s, $e) $Script:DialogAntwort = $true; [System.Windows.Window]::GetWindow($s).Close() })
    $d.FindName('Abbrechen').Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).Close() })
    if ($Vorschau) { return $d }
    [void]$d.ShowDialog()
    return $Script:DialogAntwort
}

function ConvertTo-Aktionen {
    param($Ergebnis)
    $liste = New-Object 'System.Collections.Generic.List[SdAktion]'
    foreach ($a in @($Ergebnis.Aktionen)) {
        if (-not $a) { continue }
        $x = New-Object SdAktion
        $x.Titel = [string]$a.Titel; $x.Meldung = [string]$a.Meldung; $x.Status = [string]$a.Status
        $x.Umkehrbar = (@($a.Rueckgaengig | Where-Object { $_ }).Count -gt 0)
        $liste.Add($x)
    }
    return , $liste
}

function Get-Zusammenfassung {
    param($Aktionen)
    $teile = foreach ($s in 'erledigt', 'simuliert', 'übersprungen', 'abgelehnt', 'fehlgeschlagen') {
        $n = @($Aktionen | Where-Object { $_.Status -eq $s }).Count
        if ($n) { "$n $s" }
    }
    return (@($teile) -join '  ·  ')
}

function Show-ErgebnisDialog {
    param($Ergebnis, [string] $Datei)
    $d = New-Dialog $Script:ErgebnisXaml
    $aktionen = ConvertTo-Aktionen $Ergebnis
    $istRueck = ([string]$Ergebnis.Art -eq 'Rückgängig')
    $istReinigung = Test-IstReinigung $Ergebnis
    $wort = if ($istReinigung) { 'Reinigung' } else { 'Behebung' }
    $d.FindName('Titel').Text = if ($Ergebnis.Fehler) { "$wort abgebrochen" } elseif ($istRueck) { 'Rückgängig gemacht' } else { "$wort abgeschlossen" }
    $d.FindName('Zusammenfassung').Text = Get-Zusammenfassung $aktionen
    $info = @()
    if ($Ergebnis.Fehler) { $info += "Fehler: $($Ergebnis.Fehler)" }
    $frei = Get-FreigegebenBytes $Ergebnis
    if ($frei -gt 0) { $info += $(if ($Ergebnis.Simulation) { 'Würde freigeben: ' } else { 'Freigegeben: ' }) + [SdFormat]::Groesse($frei) }
    if ($Ergebnis.Wiederherstellungspunkt -and -not $istRueck -and -not $istReinigung) { $info += "Wiederherstellungspunkt: $($Ergebnis.Wiederherstellungspunkt)" }
    $erledigt = @($aktionen | Where-Object { $_.Status -eq 'erledigt' }).Count
    # Neu geprueft bzw. gemessen wird danach automatisch (Complete-Bereinigung) - kein eigener Knopf mehr
    if ($erledigt -and -not $Ergebnis.Simulation) {
        $info += $(if ($istReinigung) { 'Nach dem Schließen wird automatisch neu gemessen - der Vergleich zeigt, was tatsächlich gelöscht wurde.' } else { 'Nach dem Schließen wird automatisch neu geprüft - der Vergleich zeigt, was sich bei den Befunden geändert hat.' })
    }
    $d.FindName('Info').Text = ($info -join "`n")
    $d.FindName('Liste').ItemsSource = $aktionen
    $kannZurueck = (-not $istRueck) -and (-not $Ergebnis.Simulation) -and (@($aktionen | Where-Object { $_.Status -eq 'erledigt' -and $_.Umkehrbar }).Count -gt 0)
    $d.FindName('Rueckgaengig').Visibility = if ($kannZurueck) { 'Visible' } else { 'Collapsed' }
    $Script:DialogAntwort = 'Schliessen'
    $d.FindName('Rueckgaengig').Add_Click({ param($s, $e) $Script:DialogAntwort = 'Rueckgaengig'; [System.Windows.Window]::GetWindow($s).Close() })
    $d.FindName('Schliessen').Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).Close() })
    if ($Vorschau) { return $d }
    [void]$d.ShowDialog()
    return $Script:DialogAntwort
}

# ============================================================================
#  Verlauf der Bereinigungen
# ============================================================================

function Update-Verlauf {
    $liste = New-Object 'System.Collections.ObjectModel.ObservableCollection[SdLauf]'
    if (Test-Path -LiteralPath $Script:BereinigungsOrdner) {
        $dateien = @(Get-ChildItem -LiteralPath $Script:BereinigungsOrdner -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^(Ergebnis|Rueckgaengig)-\d{8}_\d{6}\.json$' } | Sort-Object LastWriteTime -Descending)
        $eintraege = @(foreach ($f in $dateien) {
            try { [pscustomobject]@{ Datei = $f; Inhalt = (Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json) } } catch { }
        })
        # Rueckgaengig zaehlt nur, wenn es wirklich lief: Simulationen zaehlen nicht, mit Fehlern nur als „teilweise“
        # (dann bleibt der Knopf aktiv - ein erneuter Versuch setzt dieselben alten Werte noch einmal, das ist unschaedlich)
        $zurueck = @{}
        foreach ($e in $eintraege) {
            $r = $e.Inhalt
            if ([string]$r.Art -ne 'Rückgängig' -or -not $r.Bezug -or $r.Simulation) { continue }
            $zeilen = @($r.Aktionen | Where-Object { $_ })
            $ganz = -not $r.Fehler -and $zeilen.Count -gt 0 -and -not @($zeilen | Where-Object { $_.Status -ne 'erledigt' }).Count
            if ($ganz) { $zurueck[[string]$r.Bezug] = 'ganz' }
            elseif (-not $zurueck.ContainsKey([string]$r.Bezug)) { $zurueck[[string]$r.Bezug] = 'teilweise' }
        }
        foreach ($e in $eintraege) {
            $x = New-Object SdLauf
            $x.Datei = $e.Datei.FullName
            # In den Dateien heisst es weiter „Bereinigung“ - angezeigt wird „Behebung“
            $x.Art = if ($e.Inhalt.Simulation) { 'Simulation' } elseif (Test-IstReinigung $e.Inhalt) { 'Reinigung' } elseif ($e.Inhalt.Art -and [string]$e.Inhalt.Art -ne 'Bereinigung') { [string]$e.Inhalt.Art } else { 'Behebung' }
            $x.Zeit = try { ([datetime]$e.Inhalt.Zeit).ToString('dd.MM.yyyy HH:mm') } catch { $e.Datei.LastWriteTime.ToString('dd.MM.yyyy HH:mm') }
            $x.Wiederherstellungspunkt = if ($e.Inhalt.Wiederherstellungspunkt) { [string]$e.Inhalt.Wiederherstellungspunkt } else { '-' }
            $x.Fehler = [string]$e.Inhalt.Fehler
            $akt = ConvertTo-Aktionen $e.Inhalt
            foreach ($a in $akt) { $x.Aktionen.Add($a) }
            $x.Zusammenfassung = Get-Zusammenfassung $akt
            $frei = Get-FreigegebenBytes $e.Inhalt
            if ($frei -gt 0) { $x.Zusammenfassung += '  ·  ' + [SdFormat]::Groesse($frei) + $(if ($e.Inhalt.Simulation) { ' (simuliert)' } else { ' frei' }) }
            $rueck = $zurueck[$e.Datei.Name]
            if ($rueck -eq 'ganz') { $x.Zusammenfassung += '  ·  rückgängig gemacht' }
            elseif ($rueck -eq 'teilweise') { $x.Zusammenfassung += '  ·  nur teilweise rückgängig gemacht - erneut versuchen möglich' }
            # Nur Behebungen DIESES PCs (alte Dateien ohne Kennung stammen von PC 1); auch fehlgeschlagene mit Sicherung
            $eigen = if ($e.Inhalt.Maschine) { [string]$e.Inhalt.Maschine -eq $Script:MaschinenId } else { $Script:MaschinenId -eq $Script:Pc1Id }
            if (-not $eigen) { $x.Zusammenfassung += '  ·  von einem anderen PC' }
            $x.KannRueckgaengig = $eigen -and ($x.Art -eq 'Behebung') -and ($rueck -ne 'ganz') -and (@($akt | Where-Object { $_.Status -in 'erledigt', 'fehlgeschlagen' -and $_.Umkehrbar }).Count -gt 0)
            $liste.Add($x)
        }
    }
    $vl = Get-Element 'VerlaufListe'
    $vl.ItemsSource = $liste
    if ($liste.Count) { $vl.SelectedIndex = 0 }
    (Get-Element 'VerlaufLeer').Visibility = if ($liste.Count) { 'Collapsed' } else { 'Visible' }
}

# ============================================================================
#  Aktualisieren: neu pruefen und mit dem vorigen Stand vergleichen
# ============================================================================

function Get-Vorher {
    return @{ Daten = $Script:Daten; Kategorien = $Script:Kategorien; ReinigungStand = $Script:ReinigungStand; Ausgefuehrt = $Script:Ausgefuehrt.Clone() }
}

# Aktualisieren prueft immer schnell nach - die gruendliche Pruefung (5-10 min) waere dafuer zu lang;
# was nur gruendlich geprueft wird, zeigt der Vergleich als 'diesmal nicht geprueft'
function Get-AktualisierenModus {
    return 'Schnell'
}

function Start-Aktualisieren {
    if ($Script:Lauf) { return }
    Start-Diagnose (Get-AktualisierenModus) (Get-Vorher)
    if ($Script:Lauf) { (Get-Element 'OverlayTitel').Text = 'Wird aktualisiert …' }
}

function Get-StatusStufe {
    param([string] $Status)
    switch ($Status) { 'Kritisch' { return 3 } 'Warnung' { return 2 } 'Info' { return 1 } default { return 0 } }
}

function Get-ModusStufe {
    param([string] $Modus)
    switch ($Modus) { 'Gruendlich' { return 3 } 'Tief' { return 2 } default { return 1 } }
}

# Befunde je Pruefung - bei doppelten Eintraegen (z. B. zwei gleich benannte Datentraeger) zaehlt der schlechteste
function Get-BefundIndex {
    param($Daten)
    $index = [ordered]@{}
    foreach ($b in @($Daten.Befunde | Where-Object { $_ } | Sort-Object BereichRang, Rang, Nr)) {
        $k = "$($b.Typ)|$($b.Pruefung)"
        if (-not $index.Contains($k) -or (Get-StatusStufe $b.Status) -gt (Get-StatusStufe $index[$k].Status)) { $index[$k] = $b }
    }
    return $index
}

function New-VergleichZeile {
    param([string] $Gruppe, [string] $Status, [string] $Titel, [string] $Meldung)
    $x = New-Object SdAktion
    $x.Gruppe = $Gruppe; $x.Status = $Status; $x.Titel = $Titel; $x.Meldung = $Meldung
    return $x
}

# Was ist seit -Vorher behoben bzw. geloescht, was bleibt offen, was ist neu?
function Get-Vergleich {
    param($VorherDaten, $VorherKategorien, $Ausgefuehrt)
    if (-not $Ausgefuehrt) { $Ausgefuehrt = @{} }
    $v = @{ Zeilen = (New-Object 'System.Collections.Generic.List[SdAktion]'); MitBefunden = $false; MitReinigung = $false
            Behoben = 0; Offen = 0; Neu = 0; KategorienOffen = 0; Frei = [int64]0 }
    # Reihenfolge im Dialog: behoben, noch offen, neu, zuletzt Neutrales (nicht geprueft, praktisch leer)
    $behoben = New-Object 'System.Collections.Generic.List[SdAktion]'
    $offen = New-Object 'System.Collections.Generic.List[SdAktion]'
    $neu = New-Object 'System.Collections.Generic.List[SdAktion]'
    $rest = New-Object 'System.Collections.Generic.List[SdAktion]'

    if ($VorherDaten -and $Script:Daten -and -not [object]::ReferenceEquals($VorherDaten, $Script:Daten)) {
        $v.MitBefunden = $true
        $alt = Get-BefundIndex $VorherDaten
        $jetzt = Get-BefundIndex $Script:Daten
        $deckt = (Get-ModusStufe $Script:Daten.Modus) -ge (Get-ModusStufe $VorherDaten.Modus)
        $gesehen = @{}
        # Archivierte Befunde bleiben auch im Vergleich aussen vor
        $archiviert = @{}
        foreach ($o in @($Script:Liste | Where-Object { $_.IstArchiviert })) { $archiviert[$o.Schluessel] = $true }
        # „nicht lesbar“ ist keine Behebung: solche Befunde - und fehlende Befunde aus einem Bereich mit nicht lesbarer
        # Quelle (die Pruefung steigt dann aus) - gelten als „diesmal nicht beurteilbar“
        $unlesbar = @{}
        foreach ($k in @($jetzt.Keys)) { if ([string]$jetzt[$k].Ergebnis -match '^(Status )?nicht lesbar') { $unlesbar[[string]$jetzt[$k].Bereich] = $true } }
        foreach ($k in @($alt.Keys)) {
            $a = $alt[$k]
            $war = Get-StatusStufe $a.Status
            $m = $Ausgefuehrt["B:$k"]
            if (($war -lt 2 -and -not $m) -or ($archiviert.ContainsKey($k) -and -not $m)) { continue }
            $gesehen[$k] = $true
            $massnahme = if ($m) { "  ·  Maßnahme $($m.Status): $($m.Meldung)" } else { '' }
            if (-not $jetzt.Contains($k)) {
                if ($unlesbar.ContainsKey([string]$a.Bereich)) { $rest.Add((New-VergleichZeile 'Befunde' 'unveraendert' $a.Pruefung "diesmal nicht beurteilbar - Datenquelle nicht lesbar (vorher $($a.Status): $($a.Ergebnis))")) }
                elseif ($deckt) { $behoben.Add((New-VergleichZeile 'Befunde' 'behoben' $a.Pruefung "wird nicht mehr gemeldet (vorher $($a.Status): $($a.Ergebnis))")) }
                else { $rest.Add((New-VergleichZeile 'Befunde' 'unveraendert' $a.Pruefung ("diesmal nicht geprüft - nur im Modus $($VorherDaten.Modus)" -replace 'Gruendlich|Tief', 'Gründlich'))) }
                continue
            }
            $n = $jetzt[$k]
            $ist = Get-StatusStufe $n.Status
            $frueher = if ([string]$n.Ergebnis -ne [string]$a.Ergebnis) { "  (vorher: $($a.Ergebnis))" } else { '' }
            if ([string]$n.Ergebnis -match '^(Status )?nicht lesbar') {
                $rest.Add((New-VergleichZeile 'Befunde' 'unveraendert' $n.Pruefung "diesmal nicht beurteilbar: $($n.Ergebnis) (vorher $($a.Status): $($a.Ergebnis))"))
            } elseif ($ist -lt 2 -and $ist -lt $war) {
                $behoben.Add((New-VergleichZeile 'Befunde' 'behoben' $n.Pruefung "jetzt $($n.Status): $($n.Ergebnis)"))
            } else {
                $wie = if ($ist -lt $war) { 'besser, aber weiterhin' } else { 'weiterhin' }
                $offen.Add((New-VergleichZeile 'Befunde' 'offen' $n.Pruefung "$wie $($n.Status): $($n.Ergebnis)$frueher$massnahme"))
            }
        }
        foreach ($k in @($jetzt.Keys)) {
            $n = $jetzt[$k]
            if ($gesehen.ContainsKey($k) -or $archiviert.ContainsKey($k) -or (Get-StatusStufe $n.Status) -lt 2) { continue }
            if ($alt.Contains($k) -and (Get-StatusStufe $alt[$k].Status) -ge 2) { continue }
            $neu.Add((New-VergleichZeile 'Befunde' 'neu' $n.Pruefung "$($n.Status): $($n.Ergebnis)"))
        }
        $v.Behoben = $behoben.Count
        $v.Offen = $offen.Count
        $v.Neu = $neu.Count
        $v.Zeilen.AddRange($behoben); $v.Zeilen.AddRange($offen); $v.Zeilen.AddRange($neu); $v.Zeilen.AddRange($rest)
    }

    if ($VorherKategorien -and $Script:Kategorien -and -not [object]::ReferenceEquals($VorherKategorien, $Script:Kategorien)) {
        $v.MitReinigung = $true
        $weg = New-Object 'System.Collections.Generic.List[SdAktion]'
        $bleibt = New-Object 'System.Collections.Generic.List[SdAktion]'
        $neutral = New-Object 'System.Collections.Generic.List[SdAktion]'
        $altK = @{}
        foreach ($k in $VorherKategorien) { if ($k) { $altK[$k.Id] = $k } }
        foreach ($k in $Script:Kategorien) {
            if (-not $k -or -not $altK.ContainsKey($k.Id)) { continue }
            $a = $altK[$k.Id]
            $m = $Ausgefuehrt["K:$($k.Id)"]
            if ($a.Messbar -and $k.Messbar) {
                $weniger = [int64]$a.Bytes - [int64]$k.Bytes
                if ($weniger -ge 1MB -or ($m -and $weniger -gt 0)) {
                    $v.Frei += $weniger
                    $weg.Add((New-VergleichZeile 'Reinigung' 'behoben' $k.Name ('{0} → {1}  ·  {2} weniger' -f $a.GroesseText, $k.GroesseText, [SdFormat]::Groesse($weniger))))
                } elseif ($m -and $k.Bytes -lt 1MB) {
                    $wieLeer = if ($k.IstLeer) { 'leer' } else { "praktisch leer ($($k.GroesseText))" }
                    $neutral.Add((New-VergleichZeile 'Reinigung' 'unveraendert' $k.Name "$wieLeer  ·  $($m.Meldung)"))
                } elseif ($m) {
                    $programm = (@($k.Laeuft -split ',\s*' | ForEach-Object { (Get-ProgrammInfo $_)[0] } | Sort-Object -Unique) -join ', ')
                    $grund = if ($k.Laeuft) { "$programm läuft gerade - ganz beenden (auch im Infobereich) und erneut reinigen" }
                             elseif ($m.Status -ne 'erledigt') { $m.Meldung }
                             else { "$($m.Meldung) - seitdem wieder angelegt oder in Benutzung" }
                    $bleibt.Add((New-VergleichZeile 'Reinigung' 'offen' $k.Name "weiterhin $($k.GroesseText)  ·  $grund"))
                }
            } elseif ($m) {
                $neutral.Add((New-VergleichZeile 'Reinigung' 'unveraendert' $k.Name "ohne Adminrechte nicht messbar  ·  $($m.Meldung)"))
            }
        }
        $v.KategorienOffen = $bleibt.Count
        $v.Zeilen.AddRange($weg); $v.Zeilen.AddRange($bleibt); $v.Zeilen.AddRange($neutral)
    }
    return $v
}

function Show-Vergleich {
    param($Vorher, [switch] $NurReinigung)
    $v = Get-Vergleich $(if (-not $NurReinigung) { $Vorher.Daten }) $Vorher.Kategorien $Vorher.Ausgefuehrt
    $d = New-Dialog $Script:VergleichXaml
    $teile = @()
    if ($v.MitBefunden) { $teile += "$($v.Behoben) behoben  ·  $($v.Offen) noch offen" + $(if ($v.Neu) { "  ·  $($v.Neu) neu" } else { '' }) }
    if ($v.MitReinigung) {
        $teile += $(if ($v.Frei -gt 0) { [SdFormat]::Groesse($v.Frei) + ' weniger belegt' } else { 'Reinigung: nichts weniger belegt' }) +
                  $(if ($v.KategorienOffen) { ", $($v.KategorienOffen) Kategorie(n) nicht geleert" } else { '' })
    }
    $d.FindName('Titel').Text = if ($v.Zeilen.Count) { 'Aktualisiert - Vergleich mit vorher' } else { 'Aktualisiert - keine Veränderung' }
    $d.FindName('Zusammenfassung').Text = ($teile -join '  ·  ')
    $info = if ($v.MitBefunden) {
        $text = "Verglichen mit dem Stand vom $($Vorher.Daten.ZeitText) (Modus $($Vorher.Daten.Modus))."
        if ($Vorher.Daten.Modus -ne $Script:Daten.Modus) { $text += " Diesmal im Modus $($Script:Daten.Modus)." }
        $text -replace 'Gruendlich', 'Gründlich'
    } elseif ($v.MitReinigung -and $Vorher.ReinigungStand) { "Verglichen mit der Messung vom $($Vorher.ReinigungStand)." }
    elseif ($v.MitReinigung) { 'Verglichen mit der vorigen Messung.' }
    else { 'Es gab noch keinen früheren Stand zum Vergleichen.' }
    if (-not $v.Zeilen.Count -and ($v.MitBefunden -or $v.MitReinigung)) { $info += ' Nichts Kritisches, keine Warnungen und bei der Reinigung keine nennenswerte Veränderung.' }
    $d.FindName('Info').Text = $info
    $ansicht = [System.Windows.Data.CollectionViewSource]::GetDefaultView($v.Zeilen)
    [void]$ansicht.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription 'Gruppe'))
    $d.FindName('Liste').ItemsSource = $ansicht
    $d.FindName('Rahmen').Visibility = if ($v.Zeilen.Count) { 'Visible' } else { 'Collapsed' }
    $d.FindName('Schliessen').Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).Close() })
    if ($Vorschau) { return $d }
    [void]$d.ShowDialog()
}

# Rollbereiche laufen unten auf 28 px weich aus, statt mitten in einer Zeile hart abgeschnitten zu wirken.
# Absolute Koordinaten, neu bei jeder Groessenaenderung: ein relativer Verlauf bezieht sich auf die Inhaltshoehe
# und wuerde bei kurzem Inhalt dessen letzte Zeile abdunkeln.
function Set-Ausblenden {
    param($Element)
    $h = $Element.ActualHeight
    if ($h -lt 60) { $Element.OpacityMask = $null; return }
    $v = [System.Windows.Media.LinearGradientBrush]::new()
    $v.MappingMode = [System.Windows.Media.BrushMappingMode]::Absolute
    $v.StartPoint = [System.Windows.Point]::new(0, $h - 28); $v.EndPoint = [System.Windows.Point]::new(0, $h)
    $v.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.Colors]::Black, 0))
    $v.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.Colors]::Transparent, 1))
    $v.Freeze()
    $Element.OpacityMask = $v
}

# ============================================================================
#  Entwicklung: Zuverlaessigkeitsindex und freier Speicher ueber alle bisherigen Laeufe
# ============================================================================
$Script:EntwicklungCache = @{}
$Script:EntwicklungSerien = $null

# Aus einer Befunddatei nur das Noetige behalten (zwischengespeichert je Dateiname)
function Read-EntwicklungsPunkt {
    param($Datei)
    if ($Script:EntwicklungCache.ContainsKey($Datei.Name)) { return $Script:EntwicklungCache[$Datei.Name] }
    $p = $null
    try {
        $d = Get-Content -LiteralPath $Datei.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        $zeit = [datetime]::ParseExact([string]$d.Zeit, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
        $tage = New-Object System.Collections.Generic.List[object]
        $z = @($d.Befunde | Where-Object { $_ -and $_.Typ -eq 'Zuverlaessigkeit' })[0]
        foreach ($t in @($z.Details | Where-Object { $_ })) {
            if ([string]$t.Tag -notmatch '^(\d{2})\.(\d{2})\.$') { continue }
            $tag = [int]$Matches[1]; $monat = [int]$Matches[2]
            $jahr = if ($monat -gt $zeit.Month) { $zeit.Year - 1 } else { $zeit.Year }
            try { $tage.Add([pscustomobject]@{ X = [datetime]::new($jahr, $monat, $tag); Y = [double]$t.Letzter }) } catch { }
        }
        $lw = [ordered]@{}
        foreach ($b in @($d.Befunde | Where-Object { $_ -and $_.Typ -eq 'Laufwerk' })) {
            if ([string]$b.Pruefung -notmatch 'Laufwerk ([A-Z]):') { continue }
            $l = $Matches[1]
            if ([string]$b.Ergebnis -match '\((\d+) %\)') { $lw[$l] = [double]$Matches[1] }
        }
        $p = [pscustomobject]@{ Zeit = $zeit; Tage = $tage.ToArray(); Laufwerke = $lw; Eigen = (Test-EigeneDaten $d) }
    } catch { }
    $Script:EntwicklungCache[$Datei.Name] = $p
    return $p
}

function Update-Entwicklung {
    $dateien = @(Get-ChildItem -LiteralPath $Script:BerichtOrdner -Filter 'Befunde-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object -Last 150)
    $punkte = @(foreach ($f in $dateien) { $p = Read-EntwicklungsPunkt $f; if ($p -and $p.Eigen) { $p } })
    # Index: je Tag gilt der zuletzt gemeldete Wert
    $index = @{}
    foreach ($p in $punkte) { foreach ($t in @($p.Tage)) { $index[$t.X] = $t.Y } }
    $indexPunkte = @($index.Keys | Sort-Object | ForEach-Object { [pscustomobject]@{ X = $_; Y = $index[$_] } })
    # Speicher: je Laufwerk ein Punkt pro Lauf
    $laufwerke = [ordered]@{}
    foreach ($p in $punkte) {
        foreach ($l in @($p.Laufwerke.Keys)) {
            if (-not $laufwerke.Contains($l)) { $laufwerke[$l] = New-Object System.Collections.Generic.List[object] }
            $laufwerke[$l].Add([pscustomobject]@{ X = $p.Zeit; Y = $p.Laufwerke[$l] })
        }
    }
    $farben = 'Accent', 'Info', 'Ok', 'Warn', 'Crit', 'Muted'
    $i = 0
    $speicher = @(foreach ($l in @($laufwerke.Keys | Sort-Object)) { @{ Name = "${l}:"; Farbe = $farben[$i % $farben.Count]; Punkte = $laufwerke[$l].ToArray() }; $i++ })
    $Script:EntwicklungSerien = @{ Index = @(@{ Name = 'Index'; Farbe = 'Accent'; Punkte = $indexPunkte }); Speicher = $speicher }
    $tief = if ($indexPunkte.Count) { ($indexPunkte | Sort-Object Y | Select-Object -First 1).Y } else { 0 }
    (Get-Element 'EntwicklungInfoIndex').Text = if ($indexPunkte.Count) { 'aktuell {0:N2}  ·  Tiefstwert {1:N2}  ·  {2} Tage' -f $indexPunkte[-1].Y, $tief, $indexPunkte.Count } else { 'noch keine Werte' }
    (Get-Element 'EntwicklungInfoSpeicher').Text = if ($punkte.Count) { "$($punkte.Count) Prüfungen seit $($punkte[0].Zeit.ToString('dd.MM.yyyy'))  ·  gestrichelt = $($Script:MinFreiProzent) % Mindestwert" } else { '' }
    Show-Entwicklung
}

$Script:MinFreiProzent = 10

function Show-Entwicklung {
    if (-not $Script:EntwicklungSerien) { return }
    $werte = @($Script:EntwicklungSerien.Index[0].Punkte | ForEach-Object { $_.Y })
    $min = if ($werte.Count) { [Math]::Max(0, [Math]::Floor(($werte | Sort-Object | Select-Object -First 1)) - 1) } else { 0 }
    Show-Kurve (Get-Element 'KurveIndex') $Script:EntwicklungSerien.Index $min 10
    Show-Kurve (Get-Element 'KurveSpeicher') $Script:EntwicklungSerien.Speicher 0 100 $Script:MinFreiProzent ' %'
}

function Add-KurvenText {
    param($Leinwand, [string] $Text, [double] $X, [double] $Y, [string] $Farbe = 'Muted', [double] $Groesse = 11)
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $Text; $t.FontSize = $Groesse
    $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $Farbe)
    [System.Windows.Controls.Canvas]::SetLeft($t, $X); [System.Windows.Controls.Canvas]::SetTop($t, $Y)
    [void]$Leinwand.Children.Add($t)
}

function Add-KurvenLinie {
    param($Leinwand, [double] $X1, [double] $Y1, [double] $X2, [double] $Y2, [string] $Farbe, [double] $Dicke = 1, [bool] $Gestrichelt = $false)
    $l = New-Object System.Windows.Shapes.Line
    $l.X1 = $X1; $l.Y1 = $Y1; $l.X2 = $X2; $l.Y2 = $Y2; $l.StrokeThickness = $Dicke
    $l.SetResourceReference([System.Windows.Shapes.Shape]::StrokeProperty, $Farbe)
    if ($Gestrichelt) { $dc = New-Object System.Windows.Media.DoubleCollection; $dc.Add(4); $dc.Add(3); $l.StrokeDashArray = $dc }
    [void]$Leinwand.Children.Add($l)
}

# Liniendiagramm; Farben als Ressourcen, damit sie beim Designwechsel mitgehen
function Show-Kurve {
    param($Leinwand, [object[]] $Serien, [double] $YMin, [double] $YMax, [double] $Grenze = -1, [string] $Einheit = '')
    $Leinwand.Children.Clear()
    $b = $Leinwand.ActualWidth; $h = $Leinwand.ActualHeight
    if ($b -lt 80 -or $h -lt 50) { return }
    $alle = @($Serien | ForEach-Object { $_.Punkte } | Where-Object { $_ } | Sort-Object X)
    if (-not $alle.Count) { Add-KurvenText $Leinwand 'Noch keine Werte - sie entstehen mit jeder Prüfung.' 4 4 'Muted' 13; return }
    $xMin = $alle[0].X; $xMax = $alle[-1].X
    if (($xMax - $xMin).TotalHours -lt 1) { $xMin = $xMin.AddHours(-12); $xMax = $xMax.AddHours(12) }
    $links = 34; $rechts = 70; $oben = 8; $unten = 22
    $pb = $b - $links - $rechts; $ph = $h - $oben - $unten
    $spanne = ($xMax - $xMin).TotalSeconds
    if ($YMax -le $YMin) { $YMax = $YMin + 1 }
    for ($k = 0; $k -le 4; $k++) {
        $wert = $YMin + ($YMax - $YMin) * $k / 4
        $y = $oben + (1 - $k / 4) * $ph
        Add-KurvenLinie $Leinwand $links $y ($links + $pb) $y 'Linie'
        Add-KurvenText $Leinwand $(if (($YMax - $YMin) -lt 8) { '{0:N1}' -f $wert } else { '{0:N0}' -f $wert }) 0 ($y - 8)
    }
    if ($Grenze -ge $YMin -and $Grenze -le $YMax) {
        $y = $oben + (1 - ($Grenze - $YMin) / ($YMax - $YMin)) * $ph
        Add-KurvenLinie $Leinwand $links $y ($links + $pb) $y 'Crit' 1 $true
    }
    foreach ($t in @(0, 0.5, 1)) {
        $x = $links + $t * $pb
        Add-KurvenText $Leinwand ($xMin.AddSeconds($spanne * $t).ToString('dd.MM.')) ($x - 14) ($oben + $ph + 4)
    }
    $schilder = New-Object System.Collections.Generic.List[object]
    foreach ($s in $Serien) {
        $pkte = @($s.Punkte | Where-Object { $_ } | Sort-Object X)
        if (-not $pkte.Count) { continue }
        $linie = New-Object System.Windows.Shapes.Polyline
        $linie.StrokeThickness = 2; $linie.StrokeLineJoin = 'Round'
        $linie.SetResourceReference([System.Windows.Shapes.Shape]::StrokeProperty, $s.Farbe)
        $pc = [System.Windows.Media.PointCollection]::new()
        foreach ($p in $pkte) {
            $wert = [Math]::Max($YMin, [Math]::Min($YMax, [double]$p.Y))
            $pc.Add([System.Windows.Point]::new($links + (($p.X - $xMin).TotalSeconds / $spanne) * $pb, $oben + (1 - ($wert - $YMin) / ($YMax - $YMin)) * $ph))
        }
        $linie.Points = $pc
        [void]$Leinwand.Children.Add($linie)
        $ende = $pc[$pc.Count - 1]
        $punkt = New-Object System.Windows.Shapes.Ellipse
        $punkt.Width = 7; $punkt.Height = 7
        $punkt.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, $s.Farbe)
        [System.Windows.Controls.Canvas]::SetLeft($punkt, $ende.X - 3.5); [System.Windows.Controls.Canvas]::SetTop($punkt, $ende.Y - 3.5)
        [void]$Leinwand.Children.Add($punkt)
        $zahl = if (($YMax - $YMin) -lt 8) { '{0:N2}' -f $pkte[-1].Y } else { '{0:N0}' -f $pkte[-1].Y }
        $text = if ($s.Name -eq 'Index') { $zahl } else { "$($s.Name) $zahl$Einheit" }
        $schilder.Add([pscustomobject]@{ X = $ende.X + 8; Y = $ende.Y - 8; Text = $text; Farbe = $s.Farbe })
    }
    # Beschriftungen am Linienende nicht uebereinander: von oben nach unten mindestens 14 px Abstand
    # und unten nicht in die Datumszeile rutschen
    $liste = @($schilder | Sort-Object Y)
    $ys = New-Object 'double[]' $liste.Count
    $letztes = -1000
    for ($k = 0; $k -lt $liste.Count; $k++) { $ys[$k] = [Math]::Max($liste[$k].Y, $letztes + 14); $letztes = $ys[$k] }
    $grenzeUnten = $oben + $ph - 8
    for ($k = $liste.Count - 1; $k -ge 0; $k--) { $ys[$k] = [Math]::Min($ys[$k], $grenzeUnten - ($liste.Count - 1 - $k) * 14) }
    for ($k = 0; $k -lt $liste.Count; $k++) { Add-KurvenText $Leinwand $liste[$k].Text $liste[$k].X $ys[$k] $liste[$k].Farbe 11.5 }
}

# ============================================================================
#  Sofort-Aktionen (harmlos, ohne Adminrechte)
# ============================================================================

function Invoke-SofortAktion {
    param($Loesung)
    $p = $null
    try { $p = $Loesung.ParameterJson | ConvertFrom-Json } catch { }
    switch ($Loesung.Id) {
        'SeiteOeffnen' {
            $seite = [string]$p.Seite
            if ($seite -eq 'devmgmt') { Start-Process -FilePath (Join-Path ([Environment]::GetFolderPath('System')) 'devmgmt.msc') }
            # Speicherdiagnose fragt selbst, ob jetzt neu gestartet werden soll - die App startet nichts neu
            elseif ($seite -eq 'mdsched') { Start-Process -FilePath (Join-Path ([Environment]::GetFolderPath('System')) 'mdsched.exe') }
            elseif ($seite -match '^(ms-settings:[a-z0-9\-]+|windowsdefender:(//[a-z]+)?)$') { Start-Process -FilePath $seite }
            else { Show-Hinweis 'Nicht erlaubt' "Die Seite '$seite' steht nicht auf der Liste der erlaubten Einstellungsseiten." }
        }
        'BefehlKopieren' {
            # Feste Liste - aus der Befunddatei kommen nur Kennung und Laufwerksbuchstabe, nie Befehlstext
            $lw = ([string]$p.Laufwerk).ToUpper()
            if ($lw -and $lw -notmatch '^[A-Z]$') { Show-Hinweis 'Nicht erlaubt' "Ungültiger Laufwerksbuchstabe '$lw'."; return }
            $befehl = switch ([string]$p.Befehl) {
                'ChkdskLesen'     { "chkdsk ${lw}:" }
                'ChkdskScan'      { "chkdsk ${lw}: /scan" }
                'ChkdskReparatur' { "chkdsk ${lw}: /f" }
                'DismAufraeumen'  { 'Dism /online /Cleanup-Image /StartComponentCleanup' }
                'WmiReparatur'    { 'winmgmt /salvagerepository' }
            }
            if (-not $befehl -or ($befehl -match 'chkdsk' -and -not $lw)) { Show-Hinweis 'Nicht erlaubt' "Der Befehl '$($p.Befehl)' steht nicht auf der Liste der erlaubten Befehle."; return }
            [System.Windows.Clipboard]::SetText($befehl)
            try { Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoExit -NoLogo -NoProfile' -Verb RunAs }
            catch { Show-Hinweis 'Nicht geöffnet' "Was? Die Admin-PowerShell wurde nicht geöffnet.`n`nWarum? Die Adminrechte wurden nicht erteilt.`n`nWas tun? Der Befehl liegt trotzdem in der Zwischenablage:`n$befehl"; return }
            Show-Hinweis 'Befehl bereit' "In der Zwischenablage liegt:`n$befehl`n`nIm blauen Admin-Fenster Strg+V (oder Rechtsklick) drücken und mit Enter bestätigen. Erst dann läuft er - vorher passiert nichts."
        }
        'ReinigungOeffnen' {
            (Get-Element 'Reiter').SelectedItem = Get-Element 'ReiterReinigung'
            if (-not $Script:Kategorien) { Start-Analyse }
        }
        'MauszeigerFreigeben' {
            if ([SdFenster]::MauszeigerFreigeben()) { Show-Hinweis 'Mauszeiger freigegeben' 'Die Sperre ist aufgehoben - der Mauszeiger kann sich wieder auf allen Bildschirmen bewegen.' }
            else { Show-Hinweis 'Nicht möglich' 'Windows hat das Aufheben der Sperre abgelehnt.' }
        }
        default { Show-Hinweis 'Unbekannte Aktion' "Die Aktion '$($Loesung.Id)' ist der App nicht bekannt." }
    }
}

# ============================================================================
#  Aufbau
# ============================================================================

$Script:App = [System.Windows.Application]::Current
if (-not $Script:App) {
    $Script:App = New-Object System.Windows.Application
    $Script:App.ShutdownMode = [System.Windows.ShutdownMode]::OnMainWindowClose
}
$Script:App.Resources = New-XamlObjekt $Script:DesignXaml
Import-Einstellungen
Set-Design

# Sicherheitsnetz: ein Fehler in einem Knopf soll eine Meldung zeigen statt die App zu schliessen (Fehler.log im Werkzeugordner)
$Script:App.Add_DispatcherUnhandledException({
    param($s, $e)
    $e.Handled = $true
    $fehler = $e.Exception; while ($fehler -is [System.Management.Automation.MethodInvocationException] -and $fehler.InnerException) { $fehler = $fehler.InnerException }
    # Erste Zeile = Meldung (für die Liste im Fehlerbericht), danach Code und Stapel
    Write-Fehlerlog ("$($fehler.Message)`r`n0x{0:X8} {1}`r`n$($e.Exception)" -f $fehler.HResult, $fehler.GetType().Name)
    Update-FehlerPunkt
    if ($Script:FehlerOffen) { return }
    $Script:FehlerOffen = $true
    try { Show-Hinweis 'Etwas hat nicht geklappt' "Was? $($fehler.Message)`n`nWarum? Ein Fehler im Programm - die App bleibt offen, nichts wurde beendet.`n`nWas tun? Den Schritt noch einmal versuchen. Hilft das nicht: oben rechts " + '„Fehler melden“' + " (Symbol mit rotem Punkt) - das packt den Fehlercode in eine Datei für den Entwickler." }
    catch { } finally { $Script:FehlerOffen = $false }
})

try {
    $ico = [System.Drawing.Icon]::ExtractAssociatedIcon((Join-Path ([Environment]::GetFolderPath('System')) 'msinfo32.exe'))
    $Script:Symbol = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon($ico.Handle, [System.Windows.Int32Rect]::Empty, [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
} catch { $Script:Symbol = $null }

$Script:Fenster = New-XamlObjekt $Script:FensterXaml
Set-Titelleiste $Script:Fenster

$Script:Uhr = New-Object System.Windows.Threading.DispatcherTimer
$Script:Uhr.Interval = [TimeSpan]::FromMilliseconds(400)
$Script:Uhr.Add_Tick({
    try { Update-Lauf; Update-LaufMini }
    catch { $Script:Uhr.Stop(); $Script:Lauf = $null; Hide-Overlay; Show-Hinweis 'Fehler' $_.Exception.Message }
})

# Nach einer Bereinigung vergleicht jede neue Pruefung bzw. Messung mit dem Stand davor
(Get-Element 'StartKnopf').Add_Click({ Start-Diagnose (Get-GewaehlterModus) $(if ($Script:Ausgefuehrt.Count) { Get-Vorher }) })
(Get-Element 'EinstellungenKnopf').Add_Click({ Show-Einstellungen })
# „Entwicklung“ erst beim Öffnen laden; SelectionChanged kommt auch aus Listen darin hoch - nur den Reiter selbst beachten
(Get-Element 'Reiter').Add_SelectionChanged({ param($s, $e) if ($e.OriginalSource -eq $s -and (Test-EntwicklungOffen)) { Update-Entwicklung } })
foreach ($a in 'Massnahmen', 'Aenderungen', 'Entwicklung') {
    (Get-Element "Verlauf$a").Add_Checked({ param($s, $e) Set-VerlaufAnsicht ($s.Name -replace '^Verlauf', '') })
}
(Get-Element 'QuellenKnopf').Add_Click({ Show-Quellen })
(Get-Element 'FehlerKnopf').Add_Click({ Show-Fehlerbericht })
(Get-Element 'UpdateKnopf').Add_Click({ Show-Update })
(Get-Element 'QuellenZurueckKnopf').Add_Click({ Show-Quellen })
foreach ($n in 'KurveIndex', 'KurveSpeicher') { (Get-Element $n).Add_SizeChanged({ Show-Entwicklung }) }
# Am Wurzel-Grid statt am Fenster: feuert auch in der Vorschau, die nur den Inhalt in fester Groesse rendert
foreach ($n in 'BefundListe', 'Detail', 'ReinigungScroll', 'ReinigungInfoScroll') { (Get-Element $n).Add_SizeChanged({ param($s, $e) Set-Ausblenden $s }) }
# Gesamtnote: „geprüft vor …“ jede Minute nachführen; beim Start ggf. automatisch prüfen
$Script:NotenUhr = New-Object System.Windows.Threading.DispatcherTimer
$Script:NotenUhr.Interval = [TimeSpan]::FromMinutes(1)
$Script:NotenUhr.Add_Tick({ try { Update-Note } catch { } })
$Script:Fenster.Add_ContentRendered({ $Script:NotenUhr.Start(); try { Show-ProfilHinweis; Start-AutoPruefung } catch { }; try { Start-UpdatePruefung } catch { } })
(Get-Element 'BereinigenKnopf').Add_Click({ Start-Bereinigung @(Get-AusgewaehlteLoesungen) 'Befunde' })
(Get-Element 'ReinigenKnopf').Add_Click({ Start-Bereinigung @(Get-ReinigungsWahl) 'Reinigung' })
(Get-Element 'AnalysierenKnopf').Add_Click({ Start-Analyse $(if (@($Script:Ausgefuehrt.Keys | Where-Object { $_ -like 'K:*' }).Count) { Get-Vorher }) })
(Get-Element 'AnalysierenLeerKnopf').Add_Click({ Start-Analyse })
(Get-Element 'OverlayAbbrechen').Add_Click({ Stop-Lauf })
(Get-Element 'OverlayMinimieren').Add_Click({ Set-OverlayKlein $true })
(Get-Element 'LaufMini').Add_Click({ Set-OverlayKlein $false })
# Zurueck in der App (z. B. nachdem Chrome geschlossen wurde): Hinweise „läuft gerade“ frisch halten
$Script:Fenster.Add_Activated({ if (-not $Script:Lauf) { try { $null = Update-LaufendeProgramme @() } catch { } } })
$Script:Fenster.Add_PreviewKeyDown({
    param($s, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::F5 -and -not $Script:Lauf) { $e.Handled = $true; Start-Aktualisieren }
    if ($e.Key -eq [System.Windows.Input.Key]::Escape -and $Script:Lauf -and (Get-Element 'Overlay').Visibility -eq 'Visible') { $e.Handled = $true; Stop-Lauf; return }
    if ($e.Key -eq [System.Windows.Input.Key]::Escape -and (Get-Element 'Reiter').SelectedItem -eq (Get-Element 'ReiterQuellen')) { $e.Handled = $true; Show-Quellen }
})
(Get-Element 'ReinigungEmpfohlenKnopf').Add_Click({ foreach ($k in @($Script:Kategorien)) { if ($k) { $k.Ausgewaehlt = ($k.Vorauswahl -and -not $k.IstLeer) } } })
$Script:ReinigungFilter = 'Immer'
(Get-Element 'ReinigungFilterImmer').Add_Checked({ $Script:ReinigungFilter = 'Immer'; Update-ReinigungsFilter })
(Get-Element 'ReinigungFilterAnsehen').Add_Checked({ $Script:ReinigungFilter = 'Ansehen'; Update-ReinigungsFilter })
(Get-Element 'ReinigungKeineKnopf').Add_Click({ foreach ($k in @($Script:Kategorien)) { if ($k) { $k.Ausgewaehlt = $false } } })
(Get-Element 'AuswahlLeerenKnopf').Add_Click({ foreach ($l in @(Get-AlleLoesungen)) { $l.Ausgewaehlt = $false } })
(Get-Element 'EmpfohleneKnopf').Add_Click({
    foreach ($b in @($Script:Liste)) {
        if ($b.Status -eq 'OK' -or $b.IstArchiviert -or $b.Empfehlung -notin @('Entfernen', 'Beheben')) { continue }
        foreach ($l in $b.Loesungen) { if ($l.IstBereinigung -and -not $l.HatDauer) { $l.Ausgewaehlt = $true } }
    }
    Set-Filter 'Loesbar'
})
(Get-Element 'OrdnerKnopf').Add_Click({
    if (-not (Test-Path -LiteralPath $Script:BerichtOrdner)) { Show-Hinweis 'Noch leer' 'Es gibt noch keine Berichte.'; return }
    Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$Script:BerichtOrdner`""
})
(Get-Element 'FilterOffen').Add_Checked({ if (-not $Script:FilterWirdGesetzt) { $Script:Filter = 'Offen'; Update-Liste } })
(Get-Element 'FilterInfo').Add_Checked({ if (-not $Script:FilterWirdGesetzt) { $Script:Filter = 'Info'; Update-Liste } })
# Erklaerung der Stufen: per Info-Symbol oder Rechtsklick auf den Schalter auf- und zuklappen
(Get-Element 'ModusInfoKnopf').Add_Checked({ (Get-Element 'ModusInfo').Visibility = 'Visible' })
(Get-Element 'ModusInfoKnopf').Add_Unchecked({ (Get-Element 'ModusInfo').Visibility = 'Collapsed' })
(Get-Element 'ModusSchalter').Add_MouseRightButtonUp({ param($s, $e) $k = Get-Element 'ModusInfoKnopf'; $k.IsChecked = -not $k.IsChecked; $e.Handled = $true })
(Get-Element 'FilterAlle').Add_Checked({ if (-not $Script:FilterWirdGesetzt) { $Script:Filter = 'Alle'; Update-Liste } })
(Get-Element 'FilterArchiv').Add_Checked({ if (-not $Script:FilterWirdGesetzt) { $Script:Filter = 'Archiv'; Update-Liste } })
(Get-Element 'ArchivKnopf').Add_Click({ $b = (Get-Element 'BefundListe').SelectedItem; if ($b) { Set-Archiviert $b (-not $b.IstArchiviert) } })
(Get-Element 'Suche').Add_TextChanged({ param($s, $e) $Script:SuchText = $s.Text.Trim().ToLowerInvariant(); Update-Liste })
(Get-Element 'BefundListe').Add_SelectionChanged({ Update-Detail })
(Get-Element 'RueckgaengigKnopf').Add_Click({
    $sel = (Get-Element 'VerlaufListe').SelectedItem
    if ($sel -and $sel.KannRueckgaengig) { Start-Rueckgaengig $sel.Datei }
})
# Sofort-Knoepfe in den Loesungs-Karten (Tag = SdLoesung)
$Script:Fenster.AddHandler([System.Windows.Controls.Button]::ClickEvent, [System.Windows.RoutedEventHandler] {
    param($s, $e)
    $q = $e.OriginalSource
    if ($q -is [System.Windows.Controls.Button] -and $q.Tag -is [SdLoesung]) { Invoke-SofortAktion $q.Tag }
})
[SdLoesung]::add_AuswahlGeaendert([EventHandler] { Update-Auswahl })
[SdKategorie]::add_AuswahlGeaendert([EventHandler] { Update-ReinigungsAuswahl })

# Berichte begrenzen (Johannes, 28.10.: 60 Tage als Kompromiss aus Nachvollziehbarkeit und Platz).
# Nur bekannte Dateinamen direkt in Berichte\, Alter aus dem Zeitstempel im Namen; die neueste Datei jeder Art bleibt
# immer (letzter Stand, Snapshot fuer den naechsten Vergleich). Die Engine selbst loescht nie - sie bleibt rein lesend.
function Remove-AlteBerichte {
    param([string] $Ordner = $Script:BerichtOrdner, [int] $Tage = 60)
    if ($Script:NurLesen -or -not (Test-Path -LiteralPath $Ordner)) { return 0 }
    $grenze = (Get-Date).Date.AddDays(-$Tage)
    $muster = '^(Befunde|Reinigung|Diagnose|Snapshot|Bericht)-(\d{4}-\d{2}-\d{2})_\d{6}(-\d+)?\.(json|html)$'
    $dateien = @(Get-ChildItem -LiteralPath $Ordner -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $muster })
    $geloescht = 0
    foreach ($gruppe in ($dateien | Group-Object { ([regex]::Match($_.Name, $muster)).Groups[1].Value })) {
        $neueste = $gruppe.Group | Sort-Object Name -Descending | Select-Object -First 1
        foreach ($f in $gruppe.Group) {
            if ($f.FullName -eq $neueste.FullName) { continue }
            $tag = [datetime]::ParseExact(([regex]::Match($f.Name, $muster)).Groups[2].Value, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
            if ($tag -ge $grenze) { continue }
            try { [IO.File]::Delete($f.FullName); $geloescht++ } catch { }
        }
    }
    return $geloescht
}

# PC erkennen, dann Archiv und letztes Ergebnis DIESES PCs laden
$Script:MaschinenId = Get-MaschinenId
try { [void](Remove-AlteBerichte) } catch { }
Update-FehlerPunkt
Import-Archiv
$letzte = Get-NeuesteEigeneDatei 'Befunde-*.json'
if ($letzte) {
    try { Import-Ergebnis $letzte.FullName } catch { (Get-Element 'Untertitel').Text = 'Letztes Ergebnis konnte nicht geladen werden (' + $_.Exception.Message + ' / Zeile ' + $_.InvocationInfo.ScriptLineNumber + ') - bitte neu prüfen.' }
    if ($Script:LetzterModus -in 'Tief', 'Gruendlich') { (Get-Element 'ModusSchalter').IsChecked = $true }
} else {
    (Get-Element 'ListeLeer').Text = 'Noch keine Prüfung. Wähle oben einen Modus und klicke auf „Prüfung starten“.'
    (Get-Element 'ListeLeer').Visibility = 'Visible'
    (Get-Element 'DetailLeerText').Text = 'Hier erscheinen die Befunde mit Erklärung und Lösung.'
}
$reinigungGeladen = $false
try { $reinigungGeladen = Import-NeuesteReinigung } catch { }
if (-not $reinigungGeladen) {
    (Get-Element 'ReinigungLeerText').Text = 'Noch nicht gemessen. Die Analyse misst in wenigen Sekunden und ohne Adminrechte, wie viel sich in jeder Kategorie freigeben lässt - gelöscht wird dabei nichts.'
    (Get-Element 'ReinigungLeer').Visibility = 'Visible'
    (Get-Element 'ReinigungScroll').Visibility = 'Collapsed'
    Update-ReinigungsAuswahl
}
Update-Verlauf

# ============================================================================
#  Vorschau (Test) oder Start
# ============================================================================
function Save-Vorschau {
    param($Element, [int] $Breite, [int] $Hoehe, [string] $Datei)
    $Element.Measure((New-Object System.Windows.Size($Breite, $Hoehe)))
    $Element.Arrange((New-Object System.Windows.Rect(0, 0, $Breite, $Hoehe)))
    $Element.UpdateLayout()
    $Script:App.Dispatcher.Invoke([Action] { }, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
    $Element.UpdateLayout()
    $bmp = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($Breite, $Hoehe, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $bmp.Render($Element)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bmp))
    $fs = [IO.File]::Create($Datei)
    try { $enc.Save($fs) } finally { $fs.Close() }
}

if ($Vorschau) {
    $lb = Get-Element 'BefundListe'
    $ziel = @($Script:Liste | Where-Object { $_.HatBereinigung -and $_.IstOffen })[0]
    if (-not $ziel) { $ziel = @($Script:Liste | Where-Object { $_.HatBereinigung })[0] }
    if ($ziel) { Set-Filter 'Alle'; $lb.SelectedItem = $ziel; foreach ($l in $ziel.Loesungen) { if ($l.IstBereinigung) { $l.Ausgewaehlt = $true } } }
    Save-Vorschau $Script:Fenster.Content 1280 860 $Vorschau
    # Ganze Befund-Karte (auch den Teil unterhalb des sichtbaren Bereichs)
    $karte = (Get-Element 'Detail').Content
    $karte.Measure((New-Object System.Windows.Size(800, [double]::PositiveInfinity)))
    Save-Vorschau $karte 800 ([int][Math]::Ceiling($karte.DesiredSize.Height)) ($Vorschau -replace '\.png$', '-karte.png')
    (Get-Element 'Reiter').SelectedItem = Get-Element 'ReiterVerlauf'
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-verlauf.png')
    (Get-Element 'Reiter').SelectedItem = Get-Element 'ReiterReinigung'
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-reinigung.png')
    # Entwicklung: erst Layout (Leinwandgroesse), dann zeichnen, dann speichern
    (Get-Element 'Reiter').SelectedItem = Get-Element 'ReiterVerlauf'
    Set-VerlaufAnsicht 'Entwicklung'
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-entwicklung.png')
    Show-Entwicklung
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-entwicklung.png')
    Set-VerlaufAnsicht 'Aenderungen'
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-aenderungen.png')
    Set-VerlaufAnsicht 'Massnahmen'
    Show-Quellen
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-quellen.png')
    # Fortschrittsfenster mit beiden Balken; Lebenszeichen am eigenen Prozess gemessen
    $Script:Lauf = @{ Art = 'Diagnose' }
    Show-Overlay 'Prüfung läuft - Gründlich' 'Liest: Komponentenspeicher-Tiefenscan (DISM)' $false
    Set-OverlayProzent 42
    Update-Aktivitaet @{ Prozess = (Get-Process -Id $PID); Log = $null }
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-fortschritt.png')
    Set-OverlayKlein $true
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-minimiert.png')
    $Script:Lauf = $null
    Hide-Overlay
    (Get-Element 'Reiter').SelectedIndex = 0
    $wahl = @(Get-AusgewaehlteLoesungen)
    if ($wahl.Count) {
        $dlg = Show-Bestaetigung $wahl
        Save-Vorschau $dlg.Content 660 560 ($Vorschau -replace '\.png$', '-dialog.png')
    }
    $rwahl = @(Get-ReinigungsWahl)
    if ($rwahl.Count) {
        $dlg = Show-Bestaetigung $rwahl
        Save-Vorschau $dlg.Content 660 760 ($Vorschau -replace '\.png$', '-reinigung-dialog.png')
    }
    # Archiv: Karte eines Befunds ohne App-Loesung, dann archiviert (nur im Speicher, Archiv.json bleibt unberuehrt)
    $ohneLoesung = @($Script:Liste | Where-Object { $_.IstOffen -and -not $_.HatBereinigung })[0]
    if ($ohneLoesung) {
        Set-Filter 'Alle'; $lb.SelectedItem = $ohneLoesung
        $karte = (Get-Element 'Detail').Content
        $karte.Measure((New-Object System.Windows.Size(800, [double]::PositiveInfinity)))
        Save-Vorschau $karte 800 ([int][Math]::Ceiling($karte.DesiredSize.Height)) ($Vorschau -replace '\.png$', '-archiv.png')
        $ohneLoesung.ArchiviertAm = (Get-Date).ToString('dd.MM.yyyy HH:mm'); $ohneLoesung.IstArchiviert = $true
        Update-Zaehler; Update-Liste
        Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-archiviert.png')
        $ohneLoesung.IstArchiviert = $false; Update-Zaehler
    }
    # Aktualisieren: vorletzter gegen letzten Stand, die letzte echte Reinigung gilt als ausgefuehrt
    $befundDateien = @(Get-ChildItem -LiteralPath $Script:BerichtOrdner -Filter 'Befunde-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    $messDateien = @(Get-ChildItem -LiteralPath $Script:BerichtOrdner -Filter 'Reinigung-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    if ($befundDateien.Count -ge 2 -and $messDateien.Count -ge 2) {
        foreach ($f in @(Get-ChildItem -LiteralPath $Script:BereinigungsOrdner -Filter 'Ergebnis-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending)) {
            $erg = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $planDatei = Join-Path $Script:BereinigungsOrdner ($f.Name -replace '^Ergebnis', 'Plan')
            if ($erg.Simulation -or -not (Test-IstReinigung $erg) -or -not (Test-Path -LiteralPath $planDatei)) { continue }
            if (-not @($erg.Aktionen | Where-Object { $_ -and $_.Status -eq 'erledigt' }).Count) { continue }
            $planWahl = foreach ($a in @((Get-Content -LiteralPath $planDatei -Raw -Encoding UTF8 | ConvertFrom-Json).Aktionen)) {
                $x = New-Object SdLoesung
                $x.Id = [string]$a.Id; $x.Titel = [string]$a.Titel; $x.BefundNr = [int]$a.BefundNr
                $x.ParameterJson = ConvertTo-Json -InputObject $a.Parameter -Compress
                $x
            }
            Add-Ausgefuehrt @($planWahl) $erg
            break
        }
        $vergleichVorher = @{
            Daten = (Get-Content -LiteralPath $befundDateien[1].FullName -Raw -Encoding UTF8 | ConvertFrom-Json)
            Kategorien = (ConvertTo-Kategorien (Get-Content -LiteralPath $messDateien[1].FullName -Raw -Encoding UTF8 | ConvertFrom-Json))
            ReinigungStand = ''; Ausgefuehrt = $Script:Ausgefuehrt.Clone()
        }
        Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-aktualisieren.png')
        $inhalt = (Show-Vergleich $vergleichVorher).Content
        $inhalt.Measure((New-Object System.Windows.Size(680, [double]::PositiveInfinity)))
        Save-Vorschau $inhalt 680 ([int][Math]::Ceiling($inhalt.DesiredSize.Height)) ($Vorschau -replace '\.png$', '-vergleich.png')
    }
    # Ergebnisdialog der neuesten Behebung (erscheint nur noch bei Fehlern/Uebersprungenem, sonst wird gleich neu geprueft)
    $letztesErgebnis = Get-ChildItem -LiteralPath $Script:BereinigungsOrdner -Filter 'Ergebnis-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
    if ($letztesErgebnis) {
        $erg = Get-Content -LiteralPath $letztesErgebnis.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        $inhalt = (Show-ErgebnisDialog $erg $letztesErgebnis.FullName).Content
        $inhalt.Measure((New-Object System.Windows.Size(620, [double]::PositiveInfinity)))
        Save-Vorschau $inhalt 620 ([int][Math]::Ceiling($inhalt.DesiredSize.Height)) ($Vorschau -replace '\.png$', '-ergebnis.png')
    }
    # Hinweis „Programme schließen“ mit Chrome und Steam als Beispiel
    $beispiel = foreach ($p in 'chrome', 'steam') {
        $x = New-Object SdProgramm
        $name, $anleitung = Get-ProgrammInfo $p
        $x.Programm = $name; $x.Anleitung = $anleitung
        $k = @($Script:Kategorien | Where-Object { $_ -and @((Get-KategorieProzesse)[$_.Id]) -contains $p })[0]
        if ($k) { $x.Kategorien.Add("$($k.Name) ($($k.GroesseText))") }
        $x
    }
    $inhalt = (Show-ProgrammeSchliessen @($beispiel) $true).Content
    $inhalt.Measure((New-Object System.Windows.Size(640, [double]::PositiveInfinity)))
    Save-Vorschau $inhalt 640 ([int][Math]::Ceiling($inhalt.DesiredSize.Height)) ($Vorschau -replace '\.png$', '-programme.png')
    # Designs: Einstellungen und das Fenster im Star-Wars-Design (Einstellungen.json bleibt unberuehrt)
    $vorherDesign = $Script:Einstellungen.Design
    $inhalt = (Show-Einstellungen).Content
    $inhalt.Measure((New-Object System.Windows.Size(640, [double]::PositiveInfinity)))
    Save-Vorschau $inhalt 640 ([int][Math]::Ceiling($inhalt.DesiredSize.Height)) ($Vorschau -replace '\.png$', '-einstellungen.png')
    $inhalt = (Show-Fehlerbericht).Content
    $inhalt.Measure((New-Object System.Windows.Size(620, [double]::PositiveInfinity)))
    Save-Vorschau $inhalt 620 ([int][Math]::Ceiling($inhalt.DesiredSize.Height)) ($Vorschau -replace '\.png$', '-fehlerbericht.png')
    $Script:Update = [pscustomobject]@{ Version = '9.9'; Datum = '02.11.2026 03:20'; Neu = @('Fehler melden: Fehlercodes als Datei an PC 1', 'Update-Knopf statt Neuinstallation') }
    $inhalt = (Show-Update).Content
    $inhalt.Measure((New-Object System.Windows.Size(560, [double]::PositiveInfinity)))
    Save-Vorschau $inhalt 560 ([int][Math]::Ceiling($inhalt.DesiredSize.Height)) ($Vorschau -replace '\.png$', '-update.png')
    $Script:Update = $null
    $Script:Einstellungen.Design = 'StarWars'; Set-Design
    (Get-Element 'Reiter').SelectedIndex = 0
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-starwars.png')
    Save-Vorschau $Script:Fenster.Content 1920 1080 ($Vorschau -replace '\.png$', '-starwars-gross.png')
    (Get-Element 'Reiter').SelectedItem = Get-Element 'ReiterReinigung'
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-starwars-reinigung.png')
    $inhalt = (Show-Einstellungen).Content
    $inhalt.Measure((New-Object System.Windows.Size(640, [double]::PositiveInfinity)))
    Save-Vorschau $inhalt 640 ([int][Math]::Ceiling($inhalt.DesiredSize.Height)) ($Vorschau -replace '\.png$', '-starwars-einstellungen.png')
    $Script:Einstellungen.Design = 'CloneWars'; Set-Design
    (Get-Element 'Reiter').SelectedIndex = 0
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-clonewars.png')
    $Script:Einstellungen.Design = 'Dunkel'; Set-Design
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-dunkel.png')
    $Script:Einstellungen.Design = 'Hell'; Set-Design
    Save-Vorschau $Script:Fenster.Content 1280 860 ($Vorschau -replace '\.png$', '-hell.png')
    $Script:Einstellungen.Design = $vorherDesign; Set-Design
    return
}

[void]$Script:App.Run($Script:Fenster)
