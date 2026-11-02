# Systemdiagnose - Updates

Update-Quelle der Windows-App Systemdiagnose. Die App laedt beim Start `manifest.json` + `manifest.sig`,
prueft die RSA-Signatur und die SHA-256-Pruefsummen und ersetzt dann die Dateien aus `app/`.

Nicht von Hand aendern - wird von `Installer\Update-veroeffentlichen.ps1` erzeugt.
