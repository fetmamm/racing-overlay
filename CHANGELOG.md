# Changelog

Alla viktiga ändringar i projektet dokumenteras här.

## [Unreleased]

### Added
- Update-knapp i toppraden (vänster om mailknappen) med symbol för uppdatering.
- Versionskontroll vid start som kan visa popup med `Update` och `Cancel`.
- Könsval i Settings under Profile (`Male`/`Female`, endast ett val åt gången, `Male` som default).
- Cat-val i Settings under Profile (`A`, `B`, `C`, `D`).
- Ny inställning i Settings: `Show W/kg warnings`.
- Info-knapp bredvid Cat i Settings som visar alla kategori-limits (male/female, 5 min, 20 min) i ett separat fönster.

### Changed
- Update-knappen visar nu både aktuell version och senaste version även när appen redan är uppdaterad, med knapp till GitHub Releases.
- `Extra W/kg column` använder nu `95%` som defaultvärde.
- Settings-fonter låsta till stabil storlek så texten blir mer enhetlig.
- `Refresh sensors` blir bara grå när användaren själv trycker på knappen (inte under auto reconnect).
- Config för EXE sparas nu i `%APPDATA%\Zwift Overlay\overlay_config.json` istället för bredvid `.exe`.
- Build/release-flöde i GitHub Actions använder återanvändbara releases `Latest` (main) och `Stable` (stable).
- README uppdaterad så den speglar nuvarande funktioner och releaseflöde.
- Versionsvisning använder nu format utan `(Beta)`.
- W/kg-varningslogik använder kön + cat-specifika gränser för 5 min, 20 min och session average.
- När både `Show W/kg warnings` och `Extra W/kg column` är aktiva färgas endast 90/95%-kolumnens värden.
- Category-limits popup är omgjord till två tydliga sektioner (`Male`/`Female`) med fetstilta rubriker och utan `Sex`-kolumn.
- Info-knappen vid Cat använder nu tydlig infosymbol (`ℹ`).
- `AGENTS.md` uppdaterad med en samlad lokal regelöversikt från chatten.
- Formatering/tecken i `AGENTS.md` justerad så regellistan visas korrekt.
- Settings-fönstrets standardbredd är minskad så fönstret blir smalare.
- Cat-valet och infoknappen i Settings/Profile är flyttade närmare texten `Cat`.
- `Speed / AVG` är inte längre ikryssad som default.
- Cat har nu alternativet `None` för att köra utan vald kategori/limits.

### Fixed
- Email-knappen stoppar inte längre på saknad supportadress utan öppnar relevant mailtjänst med bekräftelse först.
- Discord-knappen i Contact öppnar nu supportservern som default.
- `Show W/kg warnings` blir nu urkryssad och inaktiverad när Cat är `None`.
- Konfliktmarkörer i `ui.py` som kunde stoppa appstart är borttagna.
- Felaktig Tk-fontsträng i Settings (`expected integer but got "UI"`) är fixad.
- Race condition mellan `Pause/Stop` och snabb `Start` som kunde låsa live-värden är åtgärdad med robust stopp/start-hantering.
- Discord-länk i push-notis kapslas i `<...>` för att undvika preview.
- Discord-notiser visar nu beskrivningsrader per commit även vid flera commits i samma push.
- Save i Settings hanterar UI-skalningsjämförelse robust så fel som `invalid literal for int() with base 10: '10s'` inte stoppar sparning.

## [v0.4.4] - 2026-04-17

### Added
- Grundläggande overlay-funktion med livevärden, sessionkontroller och settingsfönster.
