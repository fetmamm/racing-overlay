# Overlay for Zwift

Ett första skrivbordsprogram i Python som samlar in träningsdata och visar en liten live-sammanfattning i ett separat fönster.

## Vad programmet gor just nu

- visar en minimalistisk overlay med fokus pa textvarden
- later anvandaren valja om fonstret ska vara alltid overst
- later anvandaren ange vikt i kg for att rakna ut watt per kilo
- räknar ut:
  - aktuell effekt
  - aktuell watt per kilo
  - 5 minuters rullande snitteffekt
  - 20 minuters rullande snitteffekt
  - genomsnittlig puls
  - maxpuls
  - genomsnittlig kadens
  - genomsnittlig hastighet
- innehaller en sensor-dialog dar du kan:
  - valja `effektmatare`, `pulsmatare` och `kadenssensor`
  - soka efter tillgangliga enheter via `Bluetooth LE` eller `ANT+`
  - tilldela vald enhet till ratt sensortyp
- ar forberett for en hybridlosning:
  - effekt, puls och kadens via sensorer
  - hastighet via OCR/skarmavlasning fran Zwift
- innehaller fortfarande en fungerande demokalla nar inga sensorer ar valda

## Sessionkontroller

Programmet har tre knappar:

- `Starta` för att börja en session
- `Pausa` för att stoppa inflödet men behålla statistiken
- `Nollställ` för att rensa sessionen helt

Det finns ocksa en knapp `Sensorer` som oppnar konfigurationen for sensorval.

## Sensorer och bibliotek

Bluetooth-sokning anvander `bleak` om det ar installerat och forsoker nu klassa sensorer som puls, effekt och kadens utifran BLE-tjanster och namn.

```powershell
py -m pip install bleak
```

ANT+-sokning letar nu efter kompatibel ANT+-dongel i Windows. Själva live-lasningen av sensorer over ANT+ ar fortfarande nasta steg. Nar du vill ta det steget ar det naturligt att installera:

```powershell
py -m pip install openant
py -m pip install pyusb
```

## Hastighet fran Zwift

Hastighet ska inte komma fran en extern sensor i denna losning. Tanken i projektet ar i stallet:

- effekt, puls och kadens valjs som sensorer i sensorfonstren
- hastighet lases av fran Zwift-fonstret via OCR

OCR-delen ar fortfarande en separat nasta etapp.

## Starta programmet

Kör:

```powershell
py app.py
```

Programmet startar i demoläge så att du direkt kan se att summeringarna fungerar.

Du kan ocksa starta programmet utan terminal genom att dubbelklicka pa:

```text
Start Zwift Overlay.vbs
```

## Sparade installningar

Programmet sparar installningar till `overlay_config.json` i projektmappen.

Det som sparas mellan sessioner ar bland annat:

- vikt
- om fonstret ska vara alltid overst
- valda sensorer

Andringar sparas nar du uppdaterar dem och igen nar programmet stangs.

## Nästa steg för riktig Zwift-data

### 1. Skärmläsning/OCR

Lägg till ett OCR-flöde som:

- tar skärmdumpar av ett valt område i Zwift
- läser av siffror för puls, watt, kadens och hastighet
- skickar in datapunkter till samma beräkningsmotor som redan finns

Vanliga Python-paket för detta:

- `mss` för skärmdump
- `opencv-python` för bildbehandling
- `pytesseract` eller Windows OCR för textigenkänning

### 2. ANT+ / Bluetooth

Lägg till en källa som läser direkt från sensorer:

- ANT+: ofta via `openant`
- Bluetooth Low Energy: ofta via `bleak`

Det bästa långsiktigt är att köra sensorvägen när det går, eftersom den är stabilare än OCR.

## Projektstruktur

- `app.py` startar programmet
- `zwift_overlay/models.py` innehåller datamodeller
- `zwift_overlay/stats.py` räknar ut sammanfattningar
- `zwift_overlay/ui.py` visar fönstret
- `zwift_overlay/sources/` innehåller datakällor

## Viktig notering

Jag har byggt en stabil grund som fungerar utan externa bibliotek. För att läsa riktig data från Zwift-skärmen eller hårdvara behöver vi i nästa steg välja:

- OCR-baserad lösning
- Bluetooth
- ANT+
- eller en kombination där sensorer är primär källa och OCR är reserv

## EXE for anvandare

For anvandare som inte vill installera Python:

1. Oppna repo -> `Releases`.
2. Ladda ner `Zwift Overlay.exe`.
3. Dubbelklicka pa filen for att starta programmet.

## Bygg EXE i GitHub

Projektet innehaller en GitHub Action i:

- `.github/workflows/release-windows.yml`

Nar du pushar en tagg som borjar med `v` (t.ex. `v0.3.7`) bygger workflowen en portable:

- `Zwift Overlay.exe`

