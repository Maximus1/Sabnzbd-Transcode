#region Konfiguration
# Pfad zu FFmpeg (muss korrekt sein)
$ffmpegPath = "F:\media-autobuild_suite-master\local64\bin-video\ffmpeg.exe"
$mkvextractPath = "C:\Program Files\MKVToolNix\mkvextract.exe"

# Ziel-Lautheit in LUFS (z. B. -14 für YouTube, -23 für Rundfunk)
$targetLoudness = -18
$filePath = ''
$extensions = @('.mkv', '.mp4', '.avi', '.m2ts')

# Encoder-Voreinstellungen (können angepasst werden)
$crfTarget = 22
$encoderPreset = 'medium'
$audioCodecBitrate192 = '192k'
$audioCodecBitrate128 = '128k'
$videoCodecHEVC = 'HEVC'
$force720p = $false


# Standard-Dateierweiterung für die Ausgabe
$targetExtension = '.mkv'

#endregion

#region Hilfsfunktionen
# Funktion zur Überprüfung ob die Datei bereits normalisiert ist
function Test-IsNormalized {
    param (
        [string]$file
    )

    if (!(Test-Path $mkvextractPath)) {
        Write-Error "mkvextract.exe nicht gefunden unter $mkvextractPath"
        return $false
    }

    # Temporäre Datei
    $tempXml = [System.IO.Path]::GetTempFileName()

    # Tags extrahieren und in Datei schreiben
    & $mkvextractPath tags "$file" > $tempXml 2>$null

    # Inhalt lesen (ohne BOM-Probleme)
    $xmlText = Get-Content -Path $tempXml -Raw -Encoding UTF8

    # Temporäre Datei löschen
    Remove-Item $tempXml -Force

    if ([string]::IsNullOrWhiteSpace($xmlText)) {
        return $false
    }

    try {
        [xml]$xml = $xmlText
    } catch {
        Write-Warning "Konnte XML nicht parsen: $_"
        return $false
    }

    # Prüfen ob NORMALIZED-Tag vorhanden ist
    $normalized = $xml.SelectNodes('//Simple[Name="NORMALIZED"]/String') |
                  Where-Object { $_.InnerText -eq 'true' }

    return ($normalized.Count -gt 0)
}
# Funktion zum Überprüfen der Medieninformationen
function Get-MediaInfo {# Funktion zum Extrahieren von Mediendaten mit FFmpeg
    param (
        [string]$filePath # Pfad zur Eingabedatei
    )
    $mediaInfo = @{} # Hashtable zum Speichern der Mediendaten
    try {
        # Prüfen, ob die Datei existiert
        if (!(Test-Path $filePath)) {
            Write-Host "FEHLER: Datei nicht gefunden: $filePath" -ForegroundColor Red
            return $null
        }
        # FFmpeg-Prozess starten, um Mediendaten zu extrahieren
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $ffmpegPath
        $startInfo.Arguments = "-i `"$filePath`"" # FFmpeg-Argumente für die Eingabe
        $startInfo.RedirectStandardError = $true # StandardError umleiten, um die Ausgabe zu erfassen
        $startInfo.UseShellExecute = $false # ShellExecute deaktivieren, um die Umleitung zu ermöglichen
        $startInfo.CreateNoWindow = $true # Kein Konsolenfenster erstellen

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        $process.Start() | Out-Null # Prozess starten und Ausgabe verwerfen

        $infoOutput = $process.StandardError.ReadToEnd() # StandardError lesen
        $process.WaitForExit() # Warten, bis der Prozess beendet ist

        # Überspringe bereits normalisierte Dateien
        if ($file.Name -match "_normalized") {
            Write-Host "Lösche _normalized Datei: $($file.Name)" -ForegroundColor green
            continue
        }
        # fps extrahieren
        if ($infoOutput -match "fps,\s*(\d+(\.\d+)?)") {
            $mediaInfo.FPS = [double]$matches[1]
            $script:fps = [double]$matches[1]
            #Write-Host "Extrahierte FPS: $($mediaInfo.FPS)" -ForegroundColor DarkCyan
        }
        else {
            Write-Host "WARNUNG: Konnte FPS nicht aus der FFmpeg-Ausgabe extrahieren" -ForegroundColor Yellow
            $mediaInfo.FPS = 0
        }
        #Dateigröße extrahieren
        $Size = (Get-Item $filePath).Length
        $FileSizei = "{0:N2} MB" -f ($Size / 1MB)
        Write-Host "Dateigröße: $($FileSizei)" -ForegroundColor DarkCyan

        # Dauer extrahieren
        if ($infoOutput -match "Duration:\s*(\d+):(\d+):(\d+)\.(\d+)") {
            $hours = [int]$matches[1]
            $minutes = [int]$matches[2]
            $seconds = [int]$matches[3]
            $milliseconds = [int]$matches[4]

            $totalSeconds = $hours * 3600 + $minutes * 60 + $seconds + ($milliseconds / 100)
            $mediaInfo.Duration = $totalSeconds
            $mediaInfo.DurationFormatted1 = "{0:D2}:{1:D2}:{2:D2}.{3:D2}" -f $hours, $minutes, $seconds, $milliseconds
        }
        else {
            Write-Host "WARNUNG: Konnte Dauer nicht aus der FFmpeg-Ausgabe extrahieren" -ForegroundColor Yellow
            $mediaInfo.Duration = 0
            $mediaInfo.DurationFormatted = "00:00:00.00"
        }
        # Video Codec extrahieren (robuster, erkennt auch Klammern und weitere Zeichen)
        if ($infoOutput -match "Stream\s+#\d+:\d+(?:\([\w-]+\))?:\s+Video:\s*([^\s,]+)") {
            $mediaInfo.VideoCodec = $matches[1]
        }
        elseif ($infoOutput -match "Video:\s*([^\s,]+)") {
            $mediaInfo.VideoCodec = $matches[1]
        }
        elseif ($infoOutput -match "Video:\s*([^\s,]+)\s*\(") {
            $mediaInfo.VideoCodec = $matches[1]
        }
        else {
            $mediaInfo.VideoCodec = "Unbekannt"
        }
        # Auflösung extrahieren (robuster, erkennt auch Zeilen wie "720x576 [SAR ...]")
        if ($infoOutput -match "Stream\s+#\d+:\d+(?:\([\w-]+\))?:\s+Video:.*?,\s+(\d+)x(\d+)") {
            $mediaInfo.Resolution = "$($matches[1])x$($matches[2])"
        }
        elseif ($infoOutput -match "Video:.*?,\s*[^\s,]+,\s*(\d+)x(\d+)") {
            $mediaInfo.Resolution = "$($matches[1])x$($matches[2])"
        }
        elseif ($infoOutput -match "Video:.*?(\d+)x(\d+)\s*\[SAR") {
            $mediaInfo.Resolution = "$($matches[1])x$($matches[2])"
        }
        else {
            $mediaInfo.Resolution = "Unbekannt"
        }
        #check if interlaced
        # Prüfe, ob das Video interlaced ist (mit idet-Filter)
        try {
            $startInfoIdet = New-Object System.Diagnostics.ProcessStartInfo
            $startInfoIdet.FileName = $ffmpegPath
            $startInfoIdet.Arguments = "-i `"$filePath`" -filter:v idet -frames:v 1500 -an -f null NUL"
            $startInfoIdet.RedirectStandardError = $true
            $startInfoIdet.UseShellExecute = $false
            $startInfoIdet.CreateNoWindow = $true

            $processIdet = New-Object System.Diagnostics.Process
            $processIdet.StartInfo = $startInfoIdet
            $processIdet.Start() | Out-Null
            $idetOutput = $processIdet.StandardError.ReadToEnd()
            $processIdet.WaitForExit()

            # Suche nach idet summary
            $multiFrameMatches = [regex]::Matches($idetOutput, "Multi frame detection:\s*TFF:\s*(\d+)\s*BFF:\s*(\d+)\s*Progressive:\s*(\d+)\s*Undetermined:\s*(\d+)")
            if ($multiFrameMatches.Count -gt 0) {
                $lastMatch = $multiFrameMatches[$multiFrameMatches.Count - 1]
                $tff = [int]$lastMatch.Groups[1].Value
                $bff = [int]$lastMatch.Groups[2].Value
                $prog = [int]$lastMatch.Groups[3].Value
                $undet = [int]$lastMatch.Groups[4].Value
                $mediaInfo.IDET_TFF = $tff
                $mediaInfo.IDET_BFF = $bff
                $mediaInfo.IDET_Progressive = $prog
                $mediaInfo.IDET_Undetermined = $undet
                if (($tff + $bff) -gt $prog) {
                    $mediaInfo.Interlaced = $true
                    $interlaced = $true
                } else {
                    $mediaInfo.Interlaced = $false
                    $interlaced = $false
                }
            } else {
                $mediaInfo.Interlaced = $false
                $interlaced = $false
            }
        }
        catch {
            $mediaInfo.Interlaced = $false
            $interlaced = $false
        }
        # Audiokanäle extrahieren
        # Versuche, die Anzahl der Audiokanäle aus verschiedenen FFmpeg-Ausgabeformaten zu extrahieren
        if ($infoOutput -match "Audio:.*?,\s*\d+\s*Hz,\s*([0-9\.]+)\s*channels?") {
            $mediaInfo.AudioChannels = [int]$matches[1]
        }
        elseif ($infoOutput -match "Audio:.*?,\s*\d+\s*Hz,\s*([a-zA-Z0-9\.\(\)\[\]\-]+),") {
            # Erkenne Kanalbezeichnungen wie '5.1', '7.1', 'mono', 'stereo', '5.1(side)' etc.
            $channelDesc = $matches[1]
            switch -Regex ($channelDesc) {
            "mono"   { $mediaInfo.AudioChannels = 1; break }
            "stereo" { $mediaInfo.AudioChannels = 2; break }
            "2\.1"   { $mediaInfo.AudioChannels = 3; break }
            "3\.0"   { $mediaInfo.AudioChannels = 3; break }
            "4\.0"   { $mediaInfo.AudioChannels = 4; break }
            "5\.0"   { $mediaInfo.AudioChannels = 5; break }
            "5\.1"   { $mediaInfo.AudioChannels = 6; break }
            "6\.1"   { $mediaInfo.AudioChannels = 7; break }
            "7\.1"   { $mediaInfo.AudioChannels = 8; break }
            default  {
                # Versuche, eine Zahl am Anfang zu extrahieren (z.B. '6 channels')
                if ($channelDesc -match "^(\d+)") {
                $mediaInfo.AudioChannels = [int]$matches[1]
                } else {
                $mediaInfo.AudioChannels = 0
                }
            }
            }
        }
        elseif ($infoOutput -match "Audio:.+?stereo") {
            $mediaInfo.AudioChannels = 2
        }
        elseif ($infoOutput -match "Audio:.+?mono") {
            $mediaInfo.AudioChannels = 1
        }
        else {
            $mediaInfo.AudioChannels = 0
        }
        # Audio Codec extrahieren (robusteres Pattern, um auch Klammern und Bindestriche zu erfassen)
        if ($infoOutput -match "Stream\s+#\d+:\d+(?:\([\w-]+\))?:\s+Audio:\s*([\w\-\d]+)") {
            $mediaInfo.AudioCodec = $matches[1]
        }
        elseif ($infoOutput -match "Audio:\s*([\w\-\d]+)") {
            $mediaInfo.AudioCodec = $matches[1]
        }
        else {
            $mediaInfo.AudioCodec = "Unbekannt"
        }
        Write-Host "Video: $($mediaInfo.DurationFormatted1) | $($mediaInfo.VideoCodec) | $($mediaInfo.Resolution) | Interlaced: $($mediaInfo.Interlaced)" -ForegroundColor DarkCyan
        Write-Host "Audio: $($mediaInfo.AudioChannels) Kanäle: | $($mediaInfo.AudioCodec)" -ForegroundColor DarkCyan
    }
    catch {
        Write-Host "FEHLER: Fehler beim Abrufen der Mediendaten: $_" -ForegroundColor Red
        return $null
    }
    return $mediaInfo
}
function Get-MediaInfo2 { #keine file vorhanden prüfung
    param (
        [string]$filePath
    )
    $mediaInfo = @{}
    try {
# Führe FFmpeg aus, um Informationen über die Datei zu erhalten
        # Führe expliziten Befehl statt Umleitung aus, um korrekte Fehlerausgabe zu erhalten
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $ffmpegPath
        $startInfo.Arguments = "-i `"$outputFile`""
        $startInfo.RedirectStandardError = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        $process.Start() | Out-Null

        $infoOutput = $process.StandardError.ReadToEnd()
        $process.WaitForExit()

# Extrahiere die Dauer mit einem verbesserten Regex-Pattern
        if ($infoOutput -match "Duration:\s*(\d+):(\d+):(\d+)\.(\d+)") {
            $hours = [int]$matches[1]
            $minutes = [int]$matches[2]
            $seconds = [int]$matches[3]
            $milliseconds = [int]$matches[4]
            $totalSeconds = $hours * 3600 + $minutes * 60 + $seconds + ($milliseconds / 100)
            $mediaInfo.Duration = $totalSeconds
            $mediaInfo.DurationFormatted = "{0:D2}:{1:D2}:{2:D2}.{3:D2}" -f $hours, $minutes, $seconds, $milliseconds
        }
        else {
            Write-Host "WARNUNG: Konnte Dauer nicht aus der FFmpeg-Ausgabe extrahieren" -ForegroundColor Yellow
            $mediaInfo.Duration = 0
            $mediaInfo.DurationFormatted = "00:00:00.00"
        }
# Video Codec extrahieren (robuster, erkennt auch Klammern und weitere Zeichen)
        if ($infoOutput -match "Stream\s+#\d+:\d+(?:\([\w-]+\))?:\s+Video:\s*([^\s,]+)") {
            $mediaInfo.VideoCodec = $matches[1]
        }
        elseif ($infoOutput -match "Video:\s*([^\s,]+)") {
            $mediaInfo.VideoCodec = $matches[1]
        }
        elseif ($infoOutput -match "Video:\s*([^\s,]+)\s*\(") {
            $mediaInfo.VideoCodec = $matches[1]
        }
        else {
            $mediaInfo.VideoCodec = "Unbekannt"
        }
# Auflösung extrahieren (robuster, erkennt auch Zeilen wie "720x576 [SAR ...]")
        if ($infoOutput -match "Stream\s+#\d+:\d+(?:\([\w-]+\))?:\s+Video:.*?,\s+(\d+)x(\d+)") {
            $mediaInfo.Resolution = "$($matches[1])x$($matches[2])"
        }
        elseif ($infoOutput -match "Video:.*?,\s*[^\s,]+,\s*(\d+)x(\d+)") {
            $mediaInfo.Resolution = "$($matches[1])x$($matches[2])"
        }
        elseif ($infoOutput -match "Video:.*?(\d+)x(\d+)\s*\[SAR") {
            $mediaInfo.Resolution = "$($matches[1])x$($matches[2])"
        }
        else {
            $mediaInfo.Resolution = "Unbekannt"
        }
        # Audiokanäle extrahieren
        # Versuche, die Anzahl der Audiokanäle aus verschiedenen FFmpeg-Ausgabeformaten zu extrahieren
        if ($infoOutput -match "Audio:.*?,\s*\d+\s*Hz,\s*([0-9\.]+)\s*channels?") {
            $mediaInfo.AudioChannels = [int]$matches[1]
        }
        elseif ($infoOutput -match "Audio:.*?,\s*\d+\s*Hz,\s*([a-zA-Z0-9\.\(\)\[\]\-]+),") {
            # Erkenne Kanalbezeichnungen wie '5.1', '7.1', 'mono', 'stereo', '5.1(side)' etc.
            $channelDesc = $matches[1]
            switch -Regex ($channelDesc) {
            "mono"   { $mediaInfo.AudioChannels = 1; break }
            "stereo" { $mediaInfo.AudioChannels = 2; break }
            "2\.1"   { $mediaInfo.AudioChannels = 3; break }
            "3\.0"   { $mediaInfo.AudioChannels = 3; break }
            "4\.0"   { $mediaInfo.AudioChannels = 4; break }
            "5\.0"   { $mediaInfo.AudioChannels = 5; break }
            "5\.1"   { $mediaInfo.AudioChannels = 6; break }
            "6\.1"   { $mediaInfo.AudioChannels = 7; break }
            "7\.1"   { $mediaInfo.AudioChannels = 8; break }
            default  {
                # Versuche, eine Zahl am Anfang zu extrahieren (z.B. '6 channels')
                if ($channelDesc -match "^(\d+)") {
                $mediaInfo.AudioChannels = [int]$matches[1]
                } else {
                $mediaInfo.AudioChannels = 0
                }
            }
            }
        }
        elseif ($infoOutput -match "Audio:.+?stereo") {
            $mediaInfo.AudioChannels = 2
        }
        elseif ($infoOutput -match "Audio:.+?mono") {
            $mediaInfo.AudioChannels = 1
        }
        else {
            $mediaInfo.AudioChannels = 0
        }

        # Audio Codec extrahieren (robusteres Pattern, um auch Klammern und Bindestriche zu erfassen)
        if ($infoOutput -match "Stream\s+#\d+:\d+(?:\([\w-]+\))?:\s+Audio:\s*([\w\-\d]+)") {
            $mediaInfo.AudioCodec = $matches[1]
        }
        elseif ($infoOutput -match "Audio:\s*([\w\-\d]+)") {
            $mediaInfo.AudioCodec = $matches[1]
        }
        else {
            $mediaInfo.AudioCodec = "Unbekannt"
        }
        Write-Host "Video: $($mediaInfo.DurationFormatted) | $($mediaInfo.VideoCodec) | $($mediaInfo.Resolution)" -ForegroundColor DarkCyan
        Write-Host "Audio: $($mediaInfo.AudioChannels) Kanäle: | $($mediaInfo.AudioCodec)" -ForegroundColor DarkCyan
    }
    catch {
        Write-Host "Fehler beim Abrufen der Mediendaten: $_" -ForegroundColor Red
        $mediaInfo.Duration = 0
        $mediaInfo.DurationFormatted = "00:00:00.00"
        $mediaInfo.AudioChannels = 0
    }
    return $mediaInfo
}
function Get-LoudnessInfo {# Funktion zur Lautstärkeanalyse mit FFmpeg
    param (
        [string]$filePath # Pfad zur Eingabedatei
    )
    try {
# Temporäre Datei erstellen, um die FFmpeg-Ausgabe zu speichern
        $tempOutputFile = [System.IO.Path]::GetTempFileName()
 # FFmpeg-Prozess starten, um die Lautstärke zu analysieren
        Write-Host "Starte FFmpeg zur Lautstärkeanalyse..." -ForegroundColor Cyan
        $ffmpegProcess = Start-Process -FilePath $ffmpegPath -ArgumentList "-i", "`"$($filePath)`"", "-hide_banner", "-threads", "12", "-filter_complex", "[0:a:0]ebur128=metadata=1", "-f", "null", "NUL" -NoNewWindow -PassThru -RedirectStandardError $tempOutputFile
        $ffmpegProcess.WaitForExit()
# Ausgabe aus der temporären Datei lesen
        $ffmpegOutput = Get-Content -Path $tempOutputFile -Raw
# Temporäre Datei löschen
        Remove-Item -Path $tempOutputFile -Force -ErrorAction SilentlyContinue
        return $ffmpegOutput
    }
    catch {
        Write-Host "FEHLER: Fehler beim Ausführen von FFmpeg: $_" -ForegroundColor Red
        return $null
    }
}
function Set-VolumeGain {# Funktion zur Anpassung der Lautstärke mit FFmpeg
    param (
        [string]$filePath, # Pfad zur Eingabedatei
        [double]$gain, # Der anzuwendende Gain-Wert in dB
        [string]$outputFile, # Pfad für die Ausgabedatei
        [int]$audioChannels, # Anzahl der Audiokanäle in der Eingabedatei
        [string]$videoCodec, # Video Codec der Eingabedatei
        [bool]$interlaced = $false # Gibt an, ob das Video interlaced ist

    )
    try {
# FFmpeg-Argumente basierend auf der Anzahl der Audiokanäle, Bildgröße und Quellcodec erstellen
        Write-Host "Starte FFmpeg zur Lautstärkeanpassung..." -NoNewline -ForegroundColor Cyan
        $ffmpegArguments = @()
        $ffmpegArguments = @(
            "-hide_banner", # FFmpeg-Banner ausblenden
            "-loglevel", "error", # Nur Fehler anzeigen
            "-stats", # Statistiken anzeigen
            "-y", # Überschreibe Ausgabedateien ohne Nachfrage
            "-i", "`"$($filePath)`"" # Eingabedatei
        )


# konvertieren, wenn der Video Codec nicht HEVC ist
        if ($videoCodec -ne $videoCodecHEVC -AND -not $force720p -match "true") {
            Write-Host "Video Transode..." -NoNewline -ForegroundColor Cyan
            $ffmpegArguments += @(
                "-c:v", "libx265", # Video-Codec auf HEVC setzen
                "-avoid_negative_ts", "make_zero", # Negative Timestamps vermeiden
                "-preset", $encoderPreset, # Encoder-Voreinstellung verwenden
                "-crf", $crfTarget,  # CRF-Wert für die Qualität
                "-x265-params", "aq-mode=3:psy-rd=2.0:psy-rdoq=1.0:rd=4"
                if ($interlaced -eq $true) {
                    Write-Host "Deinterlacing..." -NoNewline -ForegroundColor Cyan
                    "-vf", "yadif=0:-1:0,hqdn3d=1.5:1.5:6:6" # Deinterlacing-Filter anwenden
                }else{
                    "-vf", "hqdn3d=1.5:1.5:6:6"
                }
            )
        }
# Überprüfen, ob die Auflösung 720p ist und die Variable $force720p gesetzt ist
        if ($force720p -match "true") {
            Write-Host "Video transcode mit resize..." -NoNewline -ForegroundColor Cyan
            $ffmpegArguments += @(
                "-c:v", "libx265", # Video-Codec auf HEVC setzen
                "-avoid_negative_ts", "make_zero", # Negative Timestamps vermeiden
                "-preset", $encoderPreset, # Encoder-Voreinstellung verwenden
                "-crf", $crfTarget,  # CRF-Wert für die Qualität
                "-x265-params", "aq-mode=3:psy-rd=2.0:psy-rdoq=1.0:rd=4"
                if ($interlaced -eq $true) {
                    Write-Host "Deinterlacing und Scaling..." -NoNewline -ForegroundColor Cyan
                    "-vf", "yadif=0:-1:0,scale=1280:720,hqdn3d=1.5:1.5:6:6" # Deinterlacing-Filter anwenden
                } else {
                    Write-Host "Scaling auf 720p..." -NoNewline -ForegroundColor Cyan
                    "-vf", "scale=1280:720,hqdn3d=1.5:1.5:6:6" # Auflösung auf 720p skalieren
                }
            )
        }
# Überprüfen, ob der Video-Codec bereits HEVC ist
        if ($videoCodec -eq $videoCodecHEVC -AND -not $force720p) {
            Write-Host "Video copy..." -NoNewline -ForegroundColor Cyan
            # Video-Codec beibehalten, wenn er bereits HEVC ist
            $ffmpegArguments += @(
                "-c:v", "copy" # Video Codec kopieren
            )
        }
# Überprüfen, ob die Anzahl der Audiokanäle größer als 2 ist
        if ($audioChannels -gt 2) {
            Write-Host "Audio transcode Surround..." -NoNewline -ForegroundColor Cyan
            $ffmpegArguments += @(
                "-c:a", "libfdk_aac", # Audio-Codec auf AAC setzen
                "-profile:a", "aac_he", # AAC-Profil setzen
                "-ac", $audioChannels, # Anzahl der Audiokanäle beibehalten
                "-channel_layout", "5.1" # Kanal-Layout setzen
            )
        }
# Überprüfen, ob die Anzahl der Audiokanäle gleich 2 ist
        if ($audioChannels -eq 2) {
            Write-Host "Audio transcode Stereo..." -NoNewline -ForegroundColor Cyan
            $ffmpegArguments += @(
                "-c:a", "libfdk_aac", # Audio-Codec auf AAC setzen
                "-profile:a", "aac_he", # AAC-Profil setzen
                "-b:a", $audioCodecBitrate192 # Audio-Bitrate setzen
            )
        }
# Überprüfen, ob die Anzahl der Audiokanäle kleiner oder gleich 1 ist
        if ($audioChannels -le 1) {
            Write-Host "Audio transcode Mono..." -NoNewline -ForegroundColor Cyan
            $ffmpegArguments += @(
                "-c:a", "libfdk_aac", # Audio-Codec auf AAC setzen
                "-profile:a", "aac_he", # AAC-Profil setzen
                "-b:a", $audioCodecBitrate128 # Audio-Bitrate setzen
            )
        }
# Metadaten setzen, untertitel kopieren und Lautstärke anpassen
        $ffmpegArguments += @(
            Write-Host "Lautstärke anpassung und Metadaten..."  -ForegroundColor Cyan
            "-af", "volume=${gain}dB", # Lautstärke anpassen
            "-c:s", "copy", # Untertitel kopieren
            "-metadata", "LUFS=$targetLoudness", # LUFS-Metadaten setzen
            "-metadata", "gained=$gain", # Gain-Metadaten setzen
            "-metadata", "normalized=true", # Normalisierungs-Metadaten setzen
            "`"$($outputFile)`"" # Ausgabedatei
        )
        Write-Host "FFmpeg-Argumente: $($ffmpegArguments -join ' ')" -ForegroundColor DarkCyan
# FFmpeg-Prozess zur Lautstärkeanpassung starten
        Start-Process -FilePath $ffmpegPath -ArgumentList $ffmpegArguments -NoNewWindow -Wait -PassThru -ErrorAction Stop
        Write-Host "Lautstärkeanpassung abgeschlossen für: $($filePath)" -ForegroundColor Green

    }
    catch {
        Write-Host "FEHLER: Fehler bei der Lautstärkeanpassung: $_" -ForegroundColor Red
    }
}
function Test-OutputFile {# Überprüfe die Ausgabedatei, sobald der Prozess abgeschlossen ist
    param (
        [string]$outputFile,
        [string]$sourceFile,
        [object]$sourceInfo,
        [string]$targetExtension
        )
    Write-Host "Überprüfe Ausgabedatei und Quelldatei" -ForegroundColor Cyan
# Warte kurz, um sicherzustellen, dass die Datei vollständig geschrieben wurde
    Start-Sleep -Seconds 2

    Test_Fileintregity -Outputfile $outputFile -ffmpegPath $ffmpegPath -destFolder $destFolder -file $sourceFile

    $outputInfo = Get-MediaInfo2 -filePath $outputFile
# Überprüfe, ob die Ausgabedatei korrekt erfasst wurde
    if ($outputInfo.Duration -eq 0 -or $outputInfo.AudioChannels -eq 0) {
        Write-Host "  FEHLER: Konnte Mediendaten für die Ausgabedatei nicht korrekt extrahieren." -ForegroundColor Red
        return $false
    }else {
        Write-Host "  Die Ausgabedatei wurde erfolgreich erfasst." -ForegroundColor Green
        Write-Host "  Quelldatei-Dauer: $($sourceInfo.DurationFormatted1) | Audiokanäle: $($sourceInfo.AudioChannels)" -ForegroundColor Blue
        Write-Host "  Ausgabedatei-Dauer: $($outputInfo.DurationFormatted) | Audiokanäle: $($outputInfo.AudioChannels)" -ForegroundColor Blue
        #Dateigröße extrahieren
        $Size = (Get-Item $outputFile).Length
        $FileSizeo = "{0:N2} MB" -f ($Size / 1MB)
        $Size = (Get-Item $sourceFile).Length
        $FileSizei = "{0:N2} MB" -f ($Size / 1MB)
        Write-Host "  Quelldatei-Größe: $($FileSizei)" -ForegroundColor DarkCyan
        Write-Host "  Ausgabedatei-Größe: $($FileSizeo)" -ForegroundColor DarkCyan
    }
# Überprüfe die Laufzeit beider Dateien (mit einer kleinen Toleranz von 1 Sekunde)
    $durationDiff = [Math]::Abs($sourceInfo.Duration - $outputInfo.Duration)
    if ($durationDiff -gt 1) {
        Write-Host "  WARNUNG: Die Laufzeiten unterscheiden sich um $durationDiff Sekunden!" -ForegroundColor Red
        return $false
    }else {
        Write-Host "  Die Laufzeiten stimmen überein." -ForegroundColor Green
    }
# Überprüfe die Anzahl der Audiokanäle beider Dateien
    if ($sourceInfo.AudioChannels -ne $outputInfo.AudioChannels) {
        Write-Host "  WARNUNG: Die Anzahl der Audiokanäle hat sich geändert! (Quelle: $($sourceInfo.AudioChannels), Ausgabe: $($outputInfo.AudioChannels))" -ForegroundColor Red
        return $false
    }else {
        Write-Host "  Die Anzahl der Audiokanäle ist gleich geblieben." -ForegroundColor Green
    }
    return $true
}
function Test_Fileintregity {# Überprüfe die Integrität der Datei (Streamfehler)
    param (
        [Parameter(Mandatory = $true)]
        [string]$outputFile,

        [Parameter(Mandatory = $true)]
        [string]$ffmpegPath,

        [Parameter(Mandatory = $true)]
        [string]$destFolder,

        [Parameter(Mandatory = $true)]
        [string]$file
    )

    $logDatei = Join-Path -Path $destFolder -ChildPath "MKV_Überprüfung.log"
    $date = Get-Date -Format "yyyy-MM-dd"

    Write-Host "Überprüfe Ausgabedatei: $outputFile"

# Temporäre Datei für FFmpeg-Fehlerausgabe
    $tempFehlerDatei = [System.IO.Path]::GetTempFileName()

# FFmpeg-Argumente vorbereiten
    $argumentso = @()
    $argumentso = @(
        "-v", "error",
        "-i", "`"$outputFile`"",
        "-f", "null",
        "-"
    ) -join ' '

    $argumentsi = @()
    $argumentsi = @(
        "-v", "error",
        "-i", "`"$file`"",
        "-f", "null",
        "-"
    ) -join ' '

# Ffmpeg Prozess konfigurieren für prüfung $outputFile
    $processInfoo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfoo.FileName = $ffmpegPath
    $processInfoo.Arguments = $argumentso
    $processInfoo.Arguments = $argumentsi
    $processInfoo.RedirectStandardError = $true
    $processInfoo.RedirectStandardOutput = $true
    $processInfoo.UseShellExecute = $false
    $processInfoo.CreateNoWindow = $true

# FFmpeg starten für prüfung $outputFile
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfoo
    $process.Start() | Out-Null

# Fehler auslesen und warten
    [string]$ffmpegFehlero = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    $exitCodeo = $process.ExitCode

# Fehlerausgabe zwischenspeichern
    $ffmpegFehlero | Out-File -FilePath $tempFehlerDatei -Encoding UTF8

# Auswertung des Exitcodes für das Ausgabe-File
    if ($exitCodeo -eq 0 -and [string]::IsNullOrWhiteSpace($ffmpegFehlero)) {
        Write-Host "OK: $outputFile" -ForegroundColor Green
    } else {
        "`n==== Überprüfung gestartet am $($date) ====" | Add-Content $logDatei
        Write-Host "FEHLER in Datei: $outputFile" -ForegroundColor Red
        Add-Content $logDatei "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $outputFile - FEHLER:"
        Add-Content $logDatei $ffmpegFehlero
        Add-Content $logDatei "$file wird auf fehler in der Quelle geprüft."
        Add-Content $logDatei "----------------------------------------"

# Ffmpeg Prozess konfigurieren für prüfung $inputFile
        $processInfoi = New-Object System.Diagnostics.ProcessStartInfo
        $processInfoi.FileName = $ffmpegPath
        $processInfoi.Arguments = $argumentsi
        $processInfoi.RedirectStandardError = $true
        $processInfoi.RedirectStandardOutput = $true
        $processInfoi.UseShellExecute = $false
        $processInfoi.CreateNoWindow = $true

# FFmpeg starten für prüfung $inputFile
        Write-Host "Überprüfe Quelldatei: $file"
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfoi
        $process.Start() | Out-Null

# Fehler auslesen und warten
        [string]$ffmpegFehleri = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        $exitCodei = $process.ExitCode

# Fehlerausgabe zwischenspeichern
        $ffmpegFehleri | Out-File -FilePath $tempFehlerDatei -Encoding UTF8

# Auswertung des Exitcodes für das inputFile
        if ($exitCodei -eq 0 -and [string]::IsNullOrWhiteSpace($ffmpegFehleri)) {
            Write-Host "OK: $file" -ForegroundColor Green
            Remove-Item $outputFile -Force
            Remove-Item $logDatei -Force
        } else {
# Wenn beide Dateien Fehler haben, gebe eine Warnung aus
            Write-Host "FEHLER in Datei: $file" -ForegroundColor Red
            Write-Host "$file und $outputFile haben beide fehler." -ForegroundColor Red
            Write-Host "Ersetze Quelldatei mit Ausgabedatei." -ForegroundColor green
            Add-Content $logDatei "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $file - FEHLER:"
            Add-Content $logDatei $ffmpegFehleri
            Add-Content $logDatei "$file und $outputFile haben beide fehler."
            Add-Content $logDatei "Ersetze Quelldatei mit Ausgabedatei."
            Add-Content $logDatei "----------------------------------------"
        }
    }
# Lösche die temporäre Fehlerdatei
    Remove-Item $tempFehlerDatei -Force

    "`n==== Überprüfung beendet am $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ====" | Add-Content $logDatei
    Write-Host "Überprüfung abgeschlossen. Ergebnis in: $logDatei"
}
function Remove-Files {# Funktion zum Aufräumen und Umbenennen von Dateien
    param (
        [string]$outputFile,
        [string]$sourceFile,
        [string]$targetExtension
    )
    try {
# Temporäre Datei für Umbenennung
# Wenn Test-OutputFile $true zurückgibt, lösche die Quelldatei
        if ($true) {
            $tempFile = [System.IO.Path]::Combine((Split-Path -Path $sourceFile), "$([System.IO.Path]::GetFileNameWithoutExtension($sourceFile))_temp$([System.IO.Path]::GetExtension($sourceFile))")
# Datei umbenennen mit Zwischenschritt um Namenskollisionen zu vermeiden
            Rename-Item -Path $outputFile -NewName $tempFile -Force
            Remove-Item -Path $sourceFile -Force
            $finalFile = [System.IO.Path]::Combine((Split-Path -Path $sourceFile), "$([System.IO.Path]::GetFileNameWithoutExtension($sourceFile))$targetExtension")
            Rename-Item -Path $tempFile -NewName $finalFile -Force
            Write-Host "  Erfolg: Quelldatei gelöscht und normalisierte Datei umbenannt zu $([System.IO.Path]::GetFileName($sourceFile))" -ForegroundColor Green
        } else {
# Wenn Test-OutputFile $false zurückgibt, lösche die Ausgabedatei
            Write-Host "  FEHLER: Test-OutputFile ist fehlgeschlagen. Test-OutputFile wird gelöscht." -ForegroundColor Red
            Remove-Item -Path $outputFile -Force
        }
    }
    catch {
        Write-Host "  FEHLER bei Umbenennung/Löschen: $_" -ForegroundColor Red
    }
}
#endregion

#region Hauptskript
# Ordnerauswahldialog anzeigen
Add-Type -AssemblyName System.Windows.Forms
$PickFolder = New-Object -TypeName System.Windows.Forms.OpenFileDialog
$PickFolder.FileName = 'Mediafolder'
$PickFolder.Filter = 'Folder Selection|*.*'
$PickFolder.AddExtension = $false
$PickFolder.CheckFileExists = $false
$PickFolder.Multiselect = $false
$PickFolder.CheckPathExists = $true
$PickFolder.ShowReadOnly = $false
$PickFolder.ReadOnlyChecked = $true
$PickFolder.ValidateNames = $false

$result = $PickFolder.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $destFolder = Split-Path -Path $PickFolder.FileName
    Write-Host -Object "Ausgewählter Ordner: $destFolder" -ForegroundColor Green

# Alle MKV-Dateien im ausgewählten Ordner rekursiv suchen (schnellere Methode)
    $startTime = Get-Date
    $mkvFiles = [System.IO.Directory]::EnumerateFiles($destFolder, '*.*', [System.IO.SearchOption]::AllDirectories) | Where-Object { $extensions -contains [System.IO.Path]::GetExtension($_).ToLowerInvariant() }
    $mkvFileCount = ($mkvFiles | Measure-Object).Count
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Host "Dateiscan-Zeit: $($duration.TotalSeconds) Sekunden" -ForegroundColor Yellow


# Jede MKV-Datei verarbeiten
    foreach ($file in $mkvFiles) {
        Write-Host "$mkvFileCount MKV-Dateien verbleibend." -ForegroundColor Green
        $mkvFileCount --
        Write-Host "Verarbeite Datei: $file" -ForegroundColor Cyan

# Prüfen, ob bereits NORMALIZED
        if (Test-IsNormalized $file) {
        Write-Host "Datei ist bereits normalisiert. Überspringe: $($file)" -ForegroundColor DarkGray
        continue
        }

# Mediendaten extrahieren
        $sourceInfo = Get-MediaInfo -filePath $file
        if (!$sourceInfo) {
            Write-Host "FEHLER: Konnte Mediendaten nicht extrahieren. Überspringe Datei." -ForegroundColor Red
            continue
        }

# Überprüfen, ob der Dateiname dem Serienmuster entspricht (z. B. S01E01)
        if ($file -match "S\d+E\d+") {
            Write-Host "Datei erkannt als Serientitel. Prüfe auf 720p anpassung." -ForegroundColor Yellow
# Setze Variable, um die 720p-Auflösung zu erzwingen
            if ($sourceInfo.Resolution -match "^(\d+)x(\d+)$") {
                $width = [int]$matches[1]
                $height = [int]$matches[2]
                if ($width -gt 1280 -or $height -gt 720) {
                    Write-Host "Aktuelle Auflösung: $($sourceInfo.Resolution). Größe wird auf 720p angepasst." -ForegroundColor Yellow
                    $force720p = $true
                } else {
                    Write-Host "Aktuelle Auflösung: $($sourceInfo.Resolution). Keine Größenanpassung notwendig." -ForegroundColor Yellow
                    $force720p = $false
                }
            } else {
                Write-Host "Aktuelle Auflösung: $($sourceInfo.Resolution). Keine Größenanpassung" -ForegroundColor Yellow
                $force720p = $false
            }
        } else {
# Setze Variable, um die Standardauflösung beizubehalten
            $force720p = $false
        }

# Lautstärkeinformationen extrahieren
        $ffmpegOutput = Get-LoudnessInfo -filePath $file
        if (!$ffmpegOutput) {
            Write-Host "FEHLER: Konnte Lautstärkeinformationen nicht extrahieren. Überspringe Datei." -ForegroundColor Red
            continue
        }

# Integrierte Lautheit (LUFS) extrahieren
        if ($ffmpegOutput -match "I:\s*([-\d\.]+)\s*LUFS") {
            $integratedLoudness = [double]$matches[1]
            $gain = $targetLoudness - $integratedLoudness # Notwendigen Gain berechnen

# Wenn der Gain größer als 0.1 dB ist, Lautstärke anpassen
            if ([math]::Abs($gain) -gt 0.2) {
                Write-Host "Passe Lautstärke an um $gain dB" -ForegroundColor Yellow
# Ausgabedatei erstellen
                $outputFile = [System.IO.Path]::Combine((Get-Item $file).DirectoryName, "$([System.IO.Path]::GetFileNameWithoutExtension($file))_normalized$($targetExtension)")
                Set-VolumeGain -filePath $file -gain $gain -outputFile $outputFile -audioChannels $sourceInfo.AudioChannels -videoCodec $sourceInfo.VideoCodec -interlaced $sourceInfo.Interlaced
# Überprüfen der Ausgabedatei
                Test-OutputFile -outputFile $outputFile -sourceFile $file -sourceInfo $sourceInfo -targetExtension $targetExtension
# Aufräumen und Umbenennen der Ausgabedatei
                Remove-Files -outputFile $outputFile -sourceFile $file -targetExtension $targetExtension
            }
            else {
                try {
                    Write-Host "Lautstärke bereits im Zielbereich. Keine Anpassung notwendig." -ForegroundColor Green
                    $ffmpegArgumentscopy = @()  # Array sauber neu initialisieren
                    $outputFile = [System.IO.Path]::Combine((Get-Item $file).DirectoryName, "$([System.IO.Path]::GetFileNameWithoutExtension($file))_normalized$($targetExtension)")

                    $ffmpegArgumentscopy += @(
                        "-hide_banner", # FFmpeg-Banner ausblenden
                        "-loglevel", "error", # Nur Fehler anzeigen
                        "-stats", # Statistiken anzeigen
                        "-y", # Überschreibe Ausgabedateien ohne Nachfrage
                        "-i", "`"$($file)`"", # Eingabedatei
                        "-c:v", "copy", # Video kopieren
                        "-c:a", "copy", # Audio kopieren
                        "-c:s", "copy", # Untertitel kopieren
                        "-metadata", "LUFS=$targetLoudness", # LUFS-Metadaten setzen
                        "-metadata", "gained=0", # Gain-Metadaten setzen
                        "-metadata", "normalized=true", # Normalisierungs-Metadaten setzen
                        "`"$($outputFile)`"" # Ausgabedatei
                    )
                    Write-Host "FFmpeg-Argumente: $($ffmpegArgumentscopy -join ' ')" -ForegroundColor DarkCyan
                    Start-Process -FilePath $ffmpegPath -ArgumentList $ffmpegArgumentscopy -NoNewWindow -Wait -PassThru -ErrorAction Stop
                    # Überprüfen der Ausgabedatei
                    Test-OutputFile -outputFile $outputFile -sourceFile $file -sourceInfo $sourceInfo -targetExtension $targetExtension
                    # Aufräumen und Umbenennen der Ausgabedatei
                    Remove-Files -outputFile $outputFile -sourceFile $file -targetExtension $targetExtension
                }
                catch {
                    Write-Host "FEHLER: Fehler beim schreiben der Metadaten." -ForegroundColor Red
                }
            }
        }
        else {
            Write-Host "WARNUNG: Keine LUFS-Informationen gefunden. Überspringe Lautstärkeanpassung." -ForegroundColor Yellow
        }
        Write-Host "Verarbeitung abgeschlossen für: $file" -ForegroundColor Green
        Write-Host "--------------------------------------------------" -ForegroundColor DarkGray
    }
# Nachbereitung: Lösche alle _normalized Dateien
    Write-Host "Starte Nachbereitung: Suche und lösche _normalized Dateien..." -ForegroundColor Cyan
    $normalizedFiles = [System.IO.Directory]::EnumerateFiles($destFolder, "*_normalized*$targetExtension", [System.IO.SearchOption]::AllDirectories)
    foreach ($normalizedFile in $normalizedFiles) {
        try {
            Remove-Item -Path $normalizedFile -Force
            Write-Host "  Gelöscht: $normalizedFile" -ForegroundColor Green
        }
        catch {
            Write-Host "  FEHLER: Konnte Datei nicht löschen $normalizedFile : $_" -ForegroundColor Red
        }
    }
    Write-Host "Alle Dateien verarbeitet." -ForegroundColor Green
}
else {
    Write-Host "Ordnerauswahl abgebrochen." -ForegroundColor Yellow
}
#endregion
