# Changelog

Alla viktiga ändringar i projektet dokumenteras här.

## [Unreleased]

### Added
- Update-knapp i toppraden (vänster om mailknappen) med symbol för uppdatering.
- Versionskontroll vid start som kan visa popup med `Update` och `Cancel`.
- Könsval i Settings under Profile (`Male`/`Female`, endast ett val åt gången, `Male` som default).

### Changed
- Settings-fonter låsta till stabil storlek så texten blir mer enhetlig.
- `Refresh sensors` blir bara grå när användaren själv trycker på knappen (inte under auto reconnect).
- Config för EXE sparas nu i `%APPDATA%\Zwift Overlay\overlay_config.json` istället för bredvid `.exe`.
- Build/release-flöde i GitHub Actions använder återanvändbara releases `Latest` (main) och `Stable` (stable).
- README uppdaterad så den speglar nuvarande funktioner och releaseflöde.

### Fixed
- Konfliktmarkörer i `ui.py` som kunde stoppa appstart är borttagna.
- Felaktig Tk-fontsträng i Settings (`expected integer but got "UI"`) är fixad.
- Race condition mellan `Pause/Stop` och snabb `Start` som kunde låsa live-värden är åtgärdad med robust stopp/start-hantering.
- Discord-länk i push-notis kapslas i `<...>` för att undvika preview.
- Discord-notiser visar nu beskrivningsrader per commit även vid flera commits i samma push.

## [v0.4.4] - 2026-04-17

### Added
- Grundläggande overlay-funktion med livevärden, sessionkontroller och settingsfönster.
