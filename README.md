# Aperio Imago 
# ImageScope-Batch-Extractor

Automated batch extraction and conversion of whole-slide images with ImageScope.

![Aperio Imago](AperioImago.png)


ImageScope Batch Export
Automatically exports complete `.svs` slides from ImageScope as compressed TIFF files (`TIF:LZW`).


*English Setup*

You need Windows, Aperio ImageScope, and PowerShell.
Install `winapp` in PowerShell:
```powershell
winget install Microsoft.winappcli --source winget
```
Download and Open the script and enter your SVS folder:
```powershell
$InputFolder = "E:\dataimagescope"
```
For the first test, process only one file:
```powershell
$MaxFiles = 1
```
**dont forget to set this to 0 again after testing!**

Run
Open PowerShell in the script folder and run:
```powershell
Set-ExecutionPolicy -Scope Process Bypass
& ".\ImageScope_FullBatch_RECURSIVE_ALL_Q99_v4.ps1"
```
If the test works, set `$MaxFiles = 0` to process all slides.
Keep Windows unlocked and do not use the mouse or keyboard while the script is running. Results and errors are saved in `ImageScope_batch_log.csv`. Restarting the script skips completed files.

---
*Deutsch*

Das Skript exportiert vollständige `.svs`-Slides automatisch aus ImageScope als komprimierte TIFF-Dateien (`TIF:LZW`).
Einrichtung
Du brauchst Windows, Aperio ImageScope und PowerShell.
Installiere `winapp` in PowerShell:
```powershell
winget install Microsoft.winappcli --source winget
```
Downloade und Öffne das Skript und trage deinen SVS-Ordner ein:
```powershell
$InputFolder = "E:\dataimagescope"
```
Teste zuerst nur eine Datei:
```powershell
$MaxFiles = 1
```
**nicht vergessen das wieder auf 0 zu setzen vor dem Testen!**
Start
Öffne PowerShell im Skriptordner und führe Folgendes aus:
```powershell
Set-ExecutionPolicy -Scope Process Bypass
& ".\ImageScope_FullBatch_RECURSIVE_ALL_Q99_v4.ps1"
```
Wenn der Test funktioniert, setze `$MaxFiles = 0`, um alle Slides zu verarbeiten.
Windows muss entsperrt bleiben. Benutze Maus und Tastatur während des Laufs nicht. Ergebnisse und Fehler stehen in `ImageScope_batch_log.csv`. Beim Neustart werden fertige Dateien übersprungen.
