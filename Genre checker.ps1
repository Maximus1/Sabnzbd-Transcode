Add-Type -AssemblyName System.Windows.Forms

# Function to determine the genre
function Get-Genre {
    param (
        [string]$nfoContent
    )
    # Check if the NFO file contains the genre
    if ($nfoContent -match "<genre>Animation</genre>" -or
        $nfoContent -match "<genre>Zeichentrick</genre>" -or
        $nfoContent -match "<genre>Anime</genre>") {
        return $true
    } else {
        return $false
    }
}

# Funktion zur Lautstärkeanalyse mit FFmpeg
function Get-LoudnessInfo {
    param (
        [string]$filePath
    )
        
    try {
        # Für Windows: Nutze NUL statt /dev/null
        # Führe ffmpeg mit der Lautstärkeanalyse über ebur128 aus - Ausgabe in Variable erfassen
        $tempOutputFile = [System.IO.Path]::GetTempFileName()
        $ffmpegProcess = Start-Process -FilePath $ffmpegPath -ArgumentList "-i", "`"$($filePath)`"", "-hide_banner", "-filter_complex", "ebur128=metadata=1", "-f", "null", "NUL" -NoNewWindow -PassThru -RedirectStandardError $tempOutputFile
        $ffmpegProcess.WaitForExit()
         Write-Host "Analysieren fertig" -ForegroundColor Green

        # Lese die Ausgabe aus der temporären Datei
        $ffmpegOutput = Get-Content -Path $tempOutputFile -Raw
           
        # Lösche die temporäre Datei
        Remove-Item -Path $tempOutputFile -Force -ErrorAction SilentlyContinue
        if ($ffmpegOutput -match "I:\s*([-\d\.]+)\s*LUFS") {
            Write-Host "Integrierte Lautheit für $($file.Name): $integratedLoudness LUFS" -ForegroundColor yellow
        } else {
            if ($ffmpegOutput -match "Error|Invalid") {
                Write-Host "Fehler beim Verarbeiten von $($file.Name):" -ForegroundColor Red
                Write-Host $ffmpegOutput -ForegroundColor Red #Ausgabe der Fehlermeldung
            }
        }
        return $ffmpegOutput
    }
    catch {
        Write-Host "Fehler beim Ausführen von FFmpeg: $_" -ForegroundColor Red
        # Überprüfe, ob FFmpeg Fehler ausgibt
    
        return $null
    }
}

function Get-MediaInfo { #keine file vorhanden prüfung
    param (
        [string]$filePath
    )

    if ([string]::IsNullOrWhiteSpace($filePath) -or -not (Test-Path $filePath)) {
        Write-Host "Ungültiger oder nicht vorhandener Dateipfad: $filePath" -ForegroundColor Red
        return @{ Duration = 0; DurationFormatted = "00:00:00.00"; AudioChannels = 0 }
    }

    $mediaInfo = @{}

    try {
        # Führe FFmpeg aus, um Informationen über die Datei zu erhalten
        # Führe expliziten Befehl statt Umleitung aus, um korrekte Fehlerausgabe zu erhalten
        #$tempOutputFile = [System.IO.Path]::GetTempFileName()
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $ffmpegPath
        $startInfo.Arguments = "-i `"$filePath`""
        $startInfo.RedirectStandardError = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        $process.Start() | Out-Null

        $infoOutput = $process.StandardError.ReadToEnd()
        $process.WaitForExit()


        # Extrahiere die Anzahl der Audiokanäle mit verbessertem Regex
        # Match audio stream information to extract the number of audio channels
        # Regex breakdown:
        # - "Stream\s+#\d+:\d+" matches the stream identifier (e.g., "Stream #0:1")
        # - "(?:\([\w-]+\))?" optionally matches language tags in parentheses (e.g., "(eng)")
        # - "\s+Audio:.*?" matches the "Audio" stream type and any additional details
        # - "\d+\s+Hz" matches the sample rate (e.g., "48000 Hz")
        # - "(\d+(\.\d+)?)" captures the number of audio channels (e.g., "2" or "5.1")
        if ($infoOutput -match "Stream\s+#\d+:\d+(?:\([\w-]+\))?:\s+Audio:.*?,\s+\d+\s+Hz,\s+(\d+(\.\d+)?)") {
            $mediaInfo.AudioChannels = $matches[1]
            Write-Host "Extrahierte Audiokanäle: $($mediaInfo.AudioChannels)" -ForegroundColor DarkCyan
        }
        elseif ($infoOutput -match "Stream\s+#\d+:\d+(?:\([\w-]+\))?:\s+Audio:.+?stereo") {
            $mediaInfo.AudioChannels = 2
            Write-Host "Audioformat ist stereo (2 Kanäle)" -ForegroundColor DarkCyan
        }
        elseif ($infoOutput -match "Stream\s+#\d+:\d+(?:\([\w-]+\))?:\s+Audio:.+?mono") {
            $mediaInfo.AudioChannels = 1
            Write-Host "Audioformat ist mono (1 Kanal)" -ForegroundColor DarkCyan
        }
        else {
            Write-Host "WARNUNG: Konnte Audiokanäle nicht aus der FFmpeg-Ausgabe extrahieren" -ForegroundColor Yellow
            $mediaInfo.AudioChannels = 0
        }
    }
    catch {
        Write-Host "Fehler beim Abrufen der Mediendaten: $_" -ForegroundColor Red
        $mediaInfo.Duration = 0
        $mediaInfo.DurationFormatted = "00:00:00.00"
        $mediaInfo.AudioChannels = 0
    }
    return $mediaInfo
}

# Set FFmpeg Path
$ffmpegPath = "F:\media-autobuild_suite-master\local64\bin-video\ffmpeg.exe"  # Pfad zu ffmpeg.exe anpassen

# Browse for Folder
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Wähle einen Ordner aus, um nach MKV-Dateien zu suchen:"
$folderBrowser.RootFolder = "MyComputer"  # Optional: Setze das Startverzeichnis
$folderBrowser.ShowNewFolderButton = $false

$result = $folderBrowser.ShowDialog()

if ($result -eq "OK") {
    $selectedFolder = $folderBrowser.SelectedPath

    # Get all MKV files recursively
    $mkvFiles = Get-ChildItem -Path $selectedFolder -Filter "*.mkv" -Recurse

    # Process each MKV file
    foreach ($mkvFile in $mkvFiles) {
        # Remove everything in parentheses from the filename
        $filmName = [System.IO.Path]::GetFileNameWithoutExtension($mkvFile.Name) -replace '\(.*?\)'

        Write-Output "Film: $filmName"

        # Construct the path to the NFO file
        $nfoFile = [System.IO.Path]::ChangeExtension($mkvFile.FullName, ".nfo")

        # Function to determine if the genre is animation

        # Check if the NFO file exists
        if (Test-Path $nfoFile) {
            Write-Output "NFO-Datei gefunden: $nfoFile"

            # Read the content of the NFO file
            $nfoContent = Get-Content -Path $nfoFile -Raw

            # Check if the NFO file contains the genre
            if (Get-Genre -nfoContent $nfoContent) {
            Write-Output "Genre erkannt als Anime/Animation/Zeichentrick. Starte FFmpeg Konvertierung mit -tune animation..."
            $tuneAnimation = "-tune animation"
            } else {
            Write-Output "Genre nicht erkannt."
            $tuneAnimation = ""
            }

    $ffmpegOutput = Get-LoudnessInfo -filePath $mkvFile.FullName
    # Ziel-Lautstärke definieren (z.B. -23 LUFS)
    $targetLoudness = -23
    # Extrahiere die integrierte Lautheit (LUFS) mit verbessertem Regex-Pattern
    if ($ffmpegOutput -match "I:\s*([-\d\.]+)\s*LUFS") {
        $integratedLoudness = [double]$Matches[1]
        $gain = $targetLoudness - $integratedLoudness # Berechne den notwendigen Gain-Wert

        if ([math]::Abs($gain) -gt 0.1) {
            Write-Host "Passe Lautstärke um $gain dB an für: $($mkvFile.Name)" -ForegroundColor Yellow

            # FFmpeg CLI command
            $outputFile = [System.IO.Path]::Combine($mkvFile.DirectoryName, "$($filmName)_h265.mkv")

            # Überprüfe die Anzahl der Audiokanäle, um den richtigen FFmpeg-Befehl auszuwählen
            if ($audioChannels -igt 2) {
            # Für Dateien mit mehr als 2 Audiokanälen (z.B. 5.1)
            $ffmpegArguments = @(
                "-hide_banner",
                "-i", "`"$($mkvFile.FullName)`"",
                "-c:v", "libx265",
                "-preset", "medium",
                "-crf", "23",
                $tuneAnimation,  # Add -tune animation if genre matches
                "-c:a", "libfdk_aac",
                "-profile:a", "aac_he",
                "-ac", "6",
                "-channel_layout", "5.1",
                "`"$outputFile`""
            )
            }
            if ($audioChannels -le 2) {
            $ffmpegArguments = @(
                "-hide_banner",
                "-i", "`"$($mkvFile.FullName)`"",
                "-c:v", "libx265",
                "-preset", "medium",
                "-crf", "23",
                $tuneAnimation,  # Add -tune animation if genre matches
                "-c:a", "libfdk_aac",
                "-profile:a", "aac_he",
                "-b:a", "192k",
                "-c:s", "copy",
                "`"$outputFile`""
            )
            }
            # Start FFmpeg process
            Start-Process -FilePath $ffmpegPath -ArgumentList $ffmpegArguments -NoNewWindow -PassThru -Wait -ErrorAction Stop
        } else {
            Write-Output "Keine NFO-Datei gefunden für: $($mkvFile.Name)"
        }
    }
        Write-Output "------------------------"
    }
}
} else {
    Write-Output "Ordnerauswahl abgebrochen."
}
