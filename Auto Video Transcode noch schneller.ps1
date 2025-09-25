#region Konfiguration
$ffmpegPath = $null
$mkvextractPath = $null
$targetLoudness = $null
$filePath = $null
$extensions = $null
$crfTargetm = $null
$crfTargets = $null
$encoderPreset = $null
$audioCodecBitrate160 = $null
$audioCodecBitrate128 = $null
$videoCodecHEVC = $null
$targetExtension = $null
$script:filesnotnorm = $null
$script:filesnorm = $null
$qualitätFilm = $null
$qualitätSerie = $null
# Pfad zu FFmpeg (muss korrekt sein)
$ffmpegPath = "F:\media-autobuild_suite-master1\local64\bin-video\ffmpeg.exe"
$mkvextractPath = "C:\Program Files\MKVToolNix\mkvextract.exe"

# Ziel-Lautheit in LUFS (z. B. -14 für YouTube, -23 für Rundfunk)
$targetLoudness = -18
$filePath = ''
$extensions = @('.mkv', '.mp4', '.avi', '.m2ts')

# Encoder-Voreinstellungen (können angepasst werden)
$crfTargetm = 18
$crfTargets = 20
$encoderPreset = 'medium'
#$encoderPreset = 'slow'
$audioCodecBitrate160 = '160k'
$audioCodecBitrate128 = '128k'
$videoCodecHEVC = 'HEVC'


# Standard-Dateierweiterung für die Ausgabe
$targetExtension = '.mkv'
$script:filesnotnorm = @()
$script:filesnorm = @()

# === Rekodierungsentscheidung basierend auf Laufzeit und Dateigröße ===
$qualitätFilm = "hoch"
$qualitätSerie = "hoch"


#endregion

#region Hilfsfunktionen

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

    try {
        # Tags extrahieren
        & $mkvextractPath tags "$file" > $tempXml 2>$null

        # Inhalt lesen
        $xmlText = Get-Content -Path $tempXml -Raw -Encoding UTF8

        if ([string]::IsNullOrWhiteSpace($xmlText)) {
            throw "Extrahierter XML-Inhalt ist leer oder ungültig."
        }

        try {
            [xml]$xml = $xmlText
        } catch {
            Write-Warning "Konnte XML nicht parsen (vermutlich keine MKV): $file"
            $script:filesnotnorm += $file
            return $false
        }

        # Prüfen ob NORMALIZED-Tag vorhanden ist
        $normalized = $xml.SelectNodes('//Simple[Name="NORMALIZED"]/String') |
                      Where-Object { $_.InnerText -eq 'true' }

        if ($null -eq $normalized -or $normalized.Count -eq 0) {
            Write-Host "Die Datei ist nicht normalisiert: $file" -ForegroundColor Green
            $script:filesnotnorm += $file
            return $false
        } else {
            $script:filesnorm += $file
            return $true
        }
    }
    catch {
        # Fehlerhafte Dateien trotzdem als "nicht normalisiert" markieren
        Write-Warning "Fehler beim Verarbeiten von $file`nGrund: $_"
        $script:filesnotnorm += $file
        return $false
    }
    finally {
        if (Test-Path $tempXml) {
            Remove-Item $tempXml -Force
        }
    }
}
function Get-MediaInfo {
    param ([string]$filePath)

    if (!(Test-Path -LiteralPath $filePath)) {
        Write-Host "FEHLER: Datei nicht gefunden: $filePath" -ForegroundColor Red
        return $null
    }

    $ffmpegOutput = Get-FFmpegOutput -FilePath $filePath
    $mediaInfo = @{}

    # Basis-Infos hinzufügen
    $mediaInfo += Get-BasicVideoInfo -Output $ffmpegOutput -FilePath $filePath
    $mediaInfo += Get-ColorAndHDRInfo -Output $ffmpegOutput
    $mediaInfo += Get-AudioInfo -Output $ffmpegOutput
    $mediaInfo += Get-InterlaceInfo -FilePath $filePath

    # Serienerkennung
    $mediaInfo += Is-Series -filename $filePath

    # Recode-Analyse (liefert ggf. zusätzliche Dauer)
    $mediaInfo += Get-RecodeAnalysis -MediaInfo $mediaInfo -FilePath $filePath -IsSource

    # Sicherung der Dauerwerte (nur wenn vorhanden und noch nicht gesetzt)
    if ($mediaInfo.Duration -and -not $mediaInfo.ContainsKey("Duration1")) {
        $mediaInfo.Duration1 = $mediaInfo.Duration
    }
    if ($mediaInfo.DurationFormatted -and -not $mediaInfo.ContainsKey("DurationFormatted1")) {
        $mediaInfo.DurationFormatted1 = $mediaInfo.DurationFormatted
    }

    # Konsolenausgabe
    Write-Host "Video: $($mediaInfo.DurationFormatted1) | $($mediaInfo.VideoCodec) | $($mediaInfo.Resolution) | Interlaced: $($mediaInfo.Interlaced)" -ForegroundColor DarkCyan
    Write-Host "Audio: $($mediaInfo.AudioChannels) Kanäle | $($mediaInfo.AudioCodec)" -ForegroundColor DarkCyan
    return $mediaInfo
}
#region Hilfsfunktionen zu Get-MediaInfo
function Get-FFmpegOutput {
    param ([string]$FilePath)

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $ffmpegPath
    $startInfo.Arguments = "-i `"$FilePath`""
    $startInfo.RedirectStandardError = $true
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $process.Start() | Out-Null
    $output = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    return $output
}
function Get-BasicVideoInfo {
    param (
        [string]$Output,
        [string]$FilePath
    )
    $info = @{}
    $size = (Get-Item $FilePath).Length
    $info.FileSizeBytes = $size

    if ($Output -match "fps,\s*(\d+(\.\d+)?)") {
        $info.FPS = [double]$matches[1]
    }

    if ($Output -match "Duration:\s*(\d+):(\d+):(\d+)\.(\d+)") {
        $h = [int]$matches[1]; $m = [int]$matches[2]; $s = [int]$matches[3]; $ms = [int]$matches[4]
        $info.Duration = $h * 3600 + $m * 60 + $s + ($ms / 100)
        $info.DurationFormatted1 = "{0:D2}:{1:D2}:{2:D2}.{3:D2}" -f $h, $m, $s, $ms
    }

    if ($Output -match "Video:\s*([^\s,]+)") {
        $info.VideoCodec = $matches[1]
    }

    if ($Output -match "Video:.*?,\s+(\d+)x(\d+)") {
        $info.Resolution = "$($matches[1])x$($matches[2])"
    }
    return $info
}
function Get-ColorAndHDRInfo {
    param ([string]$Output)
    $info = @{}

    if ($Output -match "yuv\d{3}p(\d{2})\w*\(([^)]*)\)") {
        $info.BitDepth = [int]$matches[1]
        $info.Is12BitOrMore = $info.BitDepth -ge 12
        $colorParts = $matches[2].Split("/")
        foreach ($part in $colorParts) {
            switch ($part.Trim()) {
                { $_ -match "^(tv|pc)$" }     { $info.ColorRange = $_ }
                { $_ -match "^bt\d+" }        { $info.ColorPrimaries = $_ }
                { $_ -match "smpte|hlg|pq" }  { $info.TransferCharacteristics = $_ }
            }
        }
    } elseif ($Output -match "yuv\d{3}p(\d{2})") {
        $info.BitDepth = [int]$matches[1]
        $info.Is12BitOrMore = $info.BitDepth -ge 12
    } else {
        $info.BitDepth = 8
        $info.Is12BitOrMore = $false
    }

    if ($Output -match "(HDR10\+?|Dolby\s+Vision|HLG|PQ|BT\.2020|smpte2084|arib-std-b67)") {
        $info.HDR = $true
        $info.HDR_Format = $matches[1]
    } else {
        $info.HDR = $false
        $info.HDR_Format = "Kein HDR"
    }

    return $info
}
function Get-AudioInfo {
    param ([string]$Output)
    $info = @{}

    if ($Output -match "Audio:.*?,\s*\d+\s*Hz,\s*([0-9\.]+)\s*channels?") {
        $info.AudioChannels = [int]$matches[1]
    } elseif ($Output -match "Audio:.*?,\s*\d+\s*Hz,\s*([^\s,]+),") {
        switch -Regex ($matches[1]) {
            "mono"   { $info.AudioChannels = 1 }
            "stereo" { $info.AudioChannels = 2 }
            "5\.1"   { $info.AudioChannels = 6 }
            "7\.1"   { $info.AudioChannels = 8 }
            default  { $info.AudioChannels = 0 }
        }
    }

    if ($Output -match "Audio:\s*([\w\-\d]+)") {
        $info.AudioCodec = $matches[1]
    }

    return $info
}
function Get-InterlaceInfo {
    param ([string]$FilePath)
    $info = @{}

    try {
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $ffmpegPath
        $startInfo.Arguments = "-i `"$FilePath`" -filter:v idet -frames:v 1500 -an -f null NUL"
        $startInfo.RedirectStandardError = $true
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $startInfo
        $proc.Start() | Out-Null
        $output = $proc.StandardError.ReadToEnd()
        $proc.WaitForExit()

        $match = [regex]::Matches($output, "Multi frame detection:\s*TFF:\s*(\d+)\s*BFF:\s*(\d+)\s*Progressive:\s*(\d+)")
        if ($match.Count -gt 0) {
            $last = $match[$match.Count - 1]
            $tff = [int]$last.Groups[1].Value
            $bff = [int]$last.Groups[2].Value
            $prog = [int]$last.Groups[3].Value
            $info.Interlaced = ($tff + $bff) -gt $prog
        }
    } catch {
        $info.Interlaced = $false
    }

    return $info
}
function Is-Series {
    param([string]$filename)
    $info = @{}
    if ($filename -match "S\d+E\d+") {
        $info.IsSeries = $true
        if ($mediaInfo.Resolution -match "^(\d+)x(\d+)$") {
            $width = [int]$matches[1]
            $height = [int]$matches[2]
            if ($width -gt 1280 -or $height -gt 720) {
                $info.resize = $true
                $info.Force720p = $true
                Write-Host "Auflösung > 1280x720 erkannt: Resize und Force720p aktiviert." -ForegroundColor Yellow
            } else {
                $info.resize = $false
                $info.Force720p = $false
            }
        }
        Write-Host "Serie erkannt: $filename" -ForegroundColor Green
    }
    else {
        $info.IsSeries = $false
        $info.Force720p = $false
        Write-Host "Keine Serie erkannt: $filename"
    }
    return $info
}
function Get-RecodeAnalysis {
    param (
        [hashtable]$MediaInfo,
        [string]$FilePath
    )

    $fileSizeBytes = (Get-Item $FilePath).Length
    $fileSizeMB = $fileSizeBytes / 1MB

    $duration = $MediaInfo.Duration
    # Erwartete Größe berechnen
    $expectedSizeMB = Calculate-ExpectedSizeMB -durationSeconds $duration -isSeries $MediaInfo.IsSeries
    # Wenn Datei >50% größer als erwartet, dann Recode empfohlen
    if ($fileSizeMB -gt ($expectedSizeMB * 1.5)) {
        $mediaInfo = @{ RecodeRecommended = $true }
        Write-Host "Recode empfohlen: Datei ist deutlich größer als erwartet ($([math]::Round($fileSizeMB,2)) MB > $expectedSizeMB MB)" -ForegroundColor Yellow
    }
    else {
        $mediaInfo = @{ RecodeRecommended = $false }
    }

    return $mediaInfo
}
function Calculate-ExpectedSizeMB {
    param (
        [double]$durationSeconds,
        [bool]$isSeries
    )

    # Qualitätsraten für Filme
    $filmRates = @{
        "niedrig" = 0.25
        "mittel"  = 0.4
        "hoch"    = 0.7
        "sehrhoch"= 1.0
    }
    # Qualitätsraten für Serien
    $serieRates = @{
        "niedrig" = 0.1
        "mittel"  = 0.14
        "hoch"    = 0.3
        "sehrhoch"= 0.5
    }

    if ($mediaInfo.IsSeries -eq $true) {
        $quality = $qualitätSerie.ToLower()
        $rates = $serieRates
    }
    else {
        $quality = $qualitätFilm.ToLower()
        $rates = $filmRates
    }

    if (-not $rates.ContainsKey($quality)) {
        Write-Warning "Qualität '$quality' nicht definiert. Nutze 'mittel'."
        $quality = "mittel"
    }

    $mbPerSecond = $rates[$quality]
    $expectedSizeMB = [math]::Round($mbPerSecond * $durationSeconds, 2)
    return $expectedSizeMB
}
#endregion

function Get-MediaInfo2 {
    param (
        [string]$filePath
    )

    $mediaInfoout = @{}

    try {
        # FFmpeg-Analyse durchführen
        $infoOutput = Get-FFmpegOutput -FilePath $filePath

        # Dauer & Formatierte Dauer extrahieren
        if ($infoOutput -match "Duration:\s*(\d+):(\d+):(\d+)\.(\d+)") {
            $h = [int]$matches[1]
            $m = [int]$matches[2]
            $s = [int]$matches[3]
            $ms = [int]$matches[4]
            $mediaInfoout.Duration = $h * 3600 + $m * 60 + $s + ($ms / 100)
            $mediaInfoout.DurationFormatted = "{0:D2}:{1:D2}:{2:D2}.{3:D2}" -f $h, $m, $s, $ms
        } else {
            Write-Host "WARNUNG: Konnte Dauer nicht extrahieren" -ForegroundColor Yellow
            $mediaInfoout.Duration = 0
            $mediaInfoout.DurationFormatted = "00:00:00.00"
        }

        # Video Infos (Codec, Auflösung)
        $videoInfo = Get-BasicVideoInfo -Output $infoOutput -FilePath $filePath
        $videoInfo.Remove("Duration") | Out-Null
        $videoInfo.Remove("DurationFormatted") | Out-Null
        $mediaInfoout += $videoInfo

        # Audio Infos (Codec, Kanäle)
        $mediaInfoout += Get-AudioInfo -Output $infoOutput

        # Log-Ausgabe
        Write-Host "Video: $($mediaInfoout.DurationFormatted) | $($mediaInfoout.VideoCodec) | $($mediaInfoout.Resolution)" -ForegroundColor DarkCyan
        Write-Host "Audio: $($mediaInfoout.AudioChannels) Kanäle | $($mediaInfoout.AudioCodec)" -ForegroundColor DarkCyan
    }
    catch {
        Write-Host "FEHLER: Medienanalyse fehlgeschlagen: $_" -ForegroundColor Red
        $mediaInfoout.Duration = 0
        $mediaInfoout.DurationFormatted = "00:00:00.00"
        $mediaInfoout.AudioChannels = 0
        $mediaInfoout.VideoCodec = "Fehler"
        $mediaInfoout.AudioCodec = "Fehler"
        $mediaInfoout.Resolution = "Unbekannt"
    }
    return $mediaInfoout
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
        [bool]$interlaced, # Gibt an, ob das Video interlaced ist
        [int]$bitDepth
    )
    try {
        Write-Host "Starte FFmpeg zur Lautstärkeanpassung..." -ForegroundColor Cyan

        $ffmpegArguments = @(
            "-hide_banner",
            "-loglevel", "error",
            "-stats",
            "-y",
            "-i", "`"$($filePath)`""
        )

# Prüfen ob BitDepth != 8 → immer reencode zu HEVC 8bit
        $needsReencodeDueToBitDepth = $false
        if ($bitDepth -ne 8) {
            Write-Host "⚠️ BitDepth ist $bitDepth, Reencode zu HEVC 8bit erforderlich" -ForegroundColor Yellow
        $needsReencodeDueToBitDepth = $true
        }

# Videocodec prüfen
        if ($mediaInfo.Force720p -or $mediaInfo.NeedsRecode -or $needsReencodeDueToBitDepth -or ($videoCodec -ne $videoCodecHEVC)) {
            Write-Host "🎞️ Transcode aktiv..." -ForegroundColor Cyan

            $ffmpegArguments += @(
                "-c:v", "libx265",
                "-pix_fmt", "yuv420p",  # 8 Bit erzwingen
                "-avoid_negative_ts", "make_zero",
                "-preset", $encoderPreset,
                "-x265-params", "aq-mode=3:psy-rd=2.0:psy-rdoq=1.0:rd=3:bframes=8:ref=4:me=3:subme=6:merange=32:deblock=-1,-1:scenecut=40:keyint=240:strong-intra-smoothing=0"
            )

# CRF je nach Serie/Film
            if ($sourceInfo.IsSeries -eq $true) {
                $ffmpegArguments += @("-crf", "$crfTargets")
                Write-Host "🎞️ Auf Serienauflösung-Anpassungen... $crfTargets" -ForegroundColor Cyan
            } else {
                $ffmpegArguments += @("-crf", "$crfTargetm")
                Write-Host "🎞️ Auf Filmauflösung-Anpassungen... $crfTargetm" -ForegroundColor Cyan
            }

# Filter setzen je nach Interlaced und Force720p
            if ($sourceInfo.Interlaced -eq $true) {
                if ($sourceInfo.Force720p -eq $true) {
                    Write-Host "↘️ Deinterlace + Scaling auf 720p" -ForegroundColor Cyan
                    $ffmpegArguments += @("-vf", "yadif=0:-1:0,scale=1280:720,hqdn3d=1.5:1.5:6:6")
                } else {
                    Write-Host "↘️ Deinterlace" -ForegroundColor Cyan
                    $ffmpegArguments += @("-vf", "yadif=0:-1:0,hqdn3d=1.5:1.5:6:6")
                }
            } elseif ($sourceInfo.Force720p -eq $true) {
                Write-Host "↘️ Scaling auf 720p" -ForegroundColor Cyan
                $ffmpegArguments += @("-vf", "scale=1280:720,hqdn3d=1.5:1.5:6:6")
            } else {
                $ffmpegArguments += @("-vf", "hqdn3d=1.5:1.5:6:6")
            }
        } else {
# Video kopieren, wenn keine Transcodierung nötig
        Write-Host "📼 Video wird kopiert (HEVC, 8 Bit und Größe OK)" -ForegroundColor Green
        $ffmpegArguments += @("-c:v", "copy")
        }

# AUDIO: Transcode je nach Kanalanzahl
        switch ($audioChannels) {
            { $_ -gt 2 } {
                Write-Host "🔊 Audio: Surround → Transcode" -ForegroundColor Cyan
                $ffmpegArguments += @(
                    "-c:a", "aac",
                    "-ac", $audioChannels,
                    "-b:a", $audioCodecBitrate160,
                    "-channel_layout", "5.1"
                )
            }
            2 {
                Write-Host "🔉 Audio: Stereo → Transcode" -ForegroundColor Cyan
                $ffmpegArguments += @(
                    "-c:a", "aac",
                    "-b:a", $audioCodecBitrate160,
                    "-ar", "48000"
                )
            }
            default {
                Write-Host "🔈 Audio: Mono → Transcode" -ForegroundColor Cyan
                $ffmpegArguments += @(
                    "-c:a", "aac",
                    "-b:a", $audioCodecBitrate128,
                    "-ar", "48000"
                )
            }
        }

# SUBS + Lautstärke + Metadaten
        $ffmpegArguments += @(
            "-af", "volume=${gain}dB",
            "-c:s", "copy",
            "-metadata", "LUFS=$targetLoudness",
            "-metadata", "gained=$gain",
            "-metadata", "normalized=true",
            "`"$($outputFile)`""
        )

        Write-Host "🧾 FFmpeg-Argumente: $($ffmpegArguments -join ' ')" -ForegroundColor DarkCyan

# FFmpeg-Prozess zur Lautstärkeanpassung starten
        $process = Start-Process -FilePath $ffmpegPath -ArgumentList $ffmpegArguments -NoNewWindow -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -eq 0) {
            Write-Host "Lautstärkeanpassung abgeschlossen für: $($filePath)" -ForegroundColor Green
            return $true
        } else {
            Write-Host "FEHLER: FFmpeg-Prozess mit Exit-Code $($process.ExitCode) beendet" -ForegroundColor Red
            return $false
        }

    }
    catch {Write-Host "FEHLER: Fehler bei der Lautstärkeanpassung: $_" -ForegroundColor Red}
}
function Test-OutputFile {# Überprüfe die Ausgabedatei, sobald der Prozess abgeschlossen ist
    param (
        [string]$outputFile,
        [string]$sourceFile,
        [object]$sourceInfo,
        [string]$targetExtension
        )
    Write-Host "Überprüfe Ausgabedatei und ggf. Quelldatei" -ForegroundColor Cyan
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
        if ($FileSizei -lt $FileSizeo) {
            Write-Host "  WARNUNG: Die Ausgabedatei ist größer als die Quelldatei!" -ForegroundColor Red
            continue
        }
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
        return $true
    }
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
    if (($exitCodeo -eq 0 -and [string]::IsNullOrWhiteSpace($ffmpegFehlero)) -or
        ($ffmpegFehlero -match "Application provided invalid, non monotonically increasing dts to muxer in stream 0")) {
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
        if (($exitCodei -eq 0 -and [string]::IsNullOrWhiteSpace($ffmpegFehleri)) -or
        ($ffmpegFehleri -match "Application provided invalid, non monotonically increasing dts to muxer in stream 0")) {
            Write-Host "OK: $file" -ForegroundColor Green
            Try {
                Remove-Item $outputFile -Force -ErrorAction Stop
                Remove-Item $logDatei -Force -ErrorAction Stop
            } catch {
                Write-Host "FEHLER beim Löschen von Dateien: $_" -ForegroundColor Red
                Write-Host "  Datei: $($_.Exception.ItemName)" -ForegroundColor Red
                Write-Host "  Fehlercode: $($_.Exception.HResult)" -ForegroundColor Red
                Write-Host "  Fehlertyp: $($_.Exception.GetType().FullName)" -ForegroundColor Red
            }
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
        [string]$targetExtension,
        [bool]$isOutputOk
    )
    try {
# Temporäre Datei für Umbenennung
# Wenn Test-OutputFile $true zurückgibt, lösche die Quelldatei
        if ($isOutputOk) {
            $tempFile = [System.IO.Path]::Combine((Split-Path -Path $sourceFile), "$([System.IO.Path]::GetFileNameWithoutExtension($sourceFile))_temp$([System.IO.Path]::GetExtension($sourceFile))")
# Datei umbenennen mit Zwischenschritt um Namenskollisionen zu vermeiden
            try {
                Rename-Item -Path $outputFile -NewName $tempFile -Force
                Remove-Item -Path $sourceFile -Force
                $finalFile = [System.IO.Path]::Combine((Split-Path -Path $sourceFile), "$([System.IO.Path]::GetFileNameWithoutExtension($sourceFile))$targetExtension")
                Rename-Item -Path $tempFile -NewName $finalFile -Force
                Write-Host "  Erfolg: Quelldatei gelöscht und normalisierte Datei umbenannt zu $([System.IO.Path]::GetFileName($sourceFile))" -ForegroundColor Green
            }
            catch {
                Write-Host "FEHLER beim Löschen von Dateien: $_" -ForegroundColor Red
                Write-Host "  Datei: $($_.Exception.ItemName)" -ForegroundColor Red
                Write-Host "  Fehlercode: $($_.Exception.HResult)" -ForegroundColor Red
                Write-Host "  Fehlertyp: $($_.Exception.GetType().FullName)" -ForegroundColor Red
            }
        } else {
# Wenn Test-OutputFile $false zurückgibt, lösche die Ausgabedatei
            Write-Host "  FEHLER: Test-OutputFile ist fehlgeschlagen. Test-OutputFile wird gelöscht." -ForegroundColor Red
            try {
                Remove-Item -Path $outputFile -Force
            }
            catch {
                Write-Host "FEHLER beim Löschen von Dateien: $_" -ForegroundColor Red
                Write-Host "  Datei: $($_.Exception.ItemName)" -ForegroundColor Red
                Write-Host "  Fehlercode: $($_.Exception.HResult)" -ForegroundColor Red
                Write-Host "  Fehlertyp: $($_.Exception.GetType().FullName)" -ForegroundColor Red
            }
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
        Write-Host "$mkvFileCount MKV-Dateien zur Tag-Prüfung verbleibend." -ForegroundColor Green
        $mkvFileCount --
        Write-Host "Verarbeite Datei: $file" -ForegroundColor Cyan

# Prüfen, ob bereits NORMALIZED
        if (Test-IsNormalized $file) {
        Write-Host "Datei ist bereits normalisiert. Überspringe: $($file)" -ForegroundColor DarkGray
        }
    }
    $mkvFileCount = ($filesnotnorm | Measure-Object).Count
    foreach ($file in $filesnotnorm) {
        Write-Host "`nStarte Verarbeitung der *nicht normalisierten* Datei: $file" -ForegroundColor Cyan
        Write-Host "$mkvFileCount MKV-Dateien zur Verarbeitung verbleibend." -ForegroundColor Green
        $mkvFileCount --

# Mediendaten extrahieren
        $sourceInfo = Get-MediaInfo -filePath $file -IsSource
        if (!$sourceInfo) {
            Write-Host "FEHLER: Konnte Mediendaten nicht extrahieren. Überspringe Datei." -ForegroundColor Red
            continue
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
#Qualität der Transkodierung einstellen.

# Ausgabedatei erstellen
                $outputFile = [System.IO.Path]::Combine((Get-Item $file).DirectoryName, "$([System.IO.Path]::GetFileNameWithoutExtension($file))_normalized$($targetExtension)")
                Set-VolumeGain -filePath $file -gain $gain -outputFile $outputFile -audioChannels $sourceInfo.AudioChannels -videoCodec $sourceInfo.VideoCodec -interlaced $sourceInfo.Interlaced -bitDepth $sourceInfo.BitDepth
# Überprüfen der Ausgabedatei und speichere das Ergebnis
                $isOutputOk = Test-OutputFile -outputFile $outputFile -sourceFile $file -sourceInfo $sourceInfo -targetExtension $targetExtension
# Aufräumen und Umbenennen der Ausgabedatei, Übergabe des Ergebnisses
                Remove-Files -outputFile $outputFile -sourceFile $file -targetExtension $targetExtension -isOutputOk $isOutputOk
            }
            else {
                try {
                    Write-Host "Lautstärke bereits im Zielbereich. Setze Metadaten." -ForegroundColor Green
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
                    $isOutputOk = Test-OutputFile -outputFile $outputFile -sourceFile $file -sourceInfo $sourceInfo -targetExtension $targetExtension
                    # Aufräumen und Umbenennen der Ausgabedatei
                    Remove-Files -outputFile $outputFile -sourceFile $file -targetExtension $targetExtension -isOutputOk $isOutputOk
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
    $normalizedFiles = [System.IO.Directory]::EnumerateFiles($destFolder, "*_normalized*", [System.IO.SearchOption]::AllDirectories)
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
