#!/bin/bash

# Determine the script's own directory
SCRIPT_DIR=$(cd "$(dirname "${BASH_SOURCE[0]}")" &>/dev/null && pwd)

# --- CONFIGURATION ---
INPUT="$1"
DIR=$(dirname "$INPUT")
FILENAME=$(basename "$INPUT")
EXT="${FILENAME##*.}"
BASENAME="${FILENAME%.*}"
INTERMEDIATE_DIR="$DIR/intermediate"
TEMP_OUTPUT="$INTERMEDIATE_DIR/$BASENAME.mkv"
FINAL_OUTPUT="$DIR/11Transcoded/$BASENAME.mkv"
LOG="$DIR/$BASENAME.log"
TARGET_LOUDNESS=-18.0


# --- LOG ROTATION ---
if [[ -f "$LOG" && $(wc -l < "$LOG") -gt 5000 ]]; then
  mv "$LOG" "$LOG.old"
  touch "$LOG"
  echo "$(date) - Log rotated (>$LOG.old)" >> "$LOG"
fi

# --- VALIDATION ---
echo "$(date) - Initial input from SABnzbd: $INPUT" >> "$LOG"

# Check if input is a directory, and if so, find the largest video file inside
if [[ -d "$INPUT" ]]; then
  echo "$(date) - Input is a directory. Searching for video file inside: $INPUT" >> "$LOG"
  # Find the largest file with a common video extension, avoiding samples.
  # The -print0 and -read -d $'\\0' handles filenames with spaces or special characters.
  # Scans one level deep.
  FOUND_FILE=$(find "$INPUT" -maxdepth 1 -type f \( -iname '*.mkv' -o -iname '*.mp4' -o -iname '*.avi' -o -iname '*.ts' -o -iname '*.mov' \) -not -iname '*sample*' -print0 | xargs -0 du -b | sort -nr | head -1 | cut -f2-)

  if [[ -n "$FOUND_FILE" && -f "$FOUND_FILE" ]]; then
    echo "$(date) - Found video file in directory: $FOUND_FILE" >> "$LOG"
    INPUT="$FOUND_FILE" # Update INPUT to the found file
    # Re-evaluate DIR, FILENAME, EXT, BASENAME based on the new INPUT
    DIR=$(dirname "$INPUT")
    FILENAME=$(basename "$INPUT")
    EXT="${FILENAME##*.}"
    BASENAME="${FILENAME%.*}"
    echo "$(date) - Updated INPUT to: $INPUT" >> "$LOG"
    echo "$(date) - Updated DIR: $DIR, FILENAME: $FILENAME, EXT: $EXT, BASENAME: $BASENAME" >> "$LOG"
  else
    echo "$(date) - ERROR: Input was a directory, but no suitable video file found inside: $INPUT" >> "$LOG"
    exit 1
  fi
elif [[ ! -f "$INPUT" ]]; then # Original check if it's not a directory AND not a file
  echo "$(date) - ERROR: Input is not a file and not a processable directory: $INPUT" >> "$LOG"
  exit 1
fi
# Proceed with the now potentially updated INPUT

echo "$(date) - Processing file: $INPUT" >> "$LOG"
# The original validation 'if [[ -z "$INPUT" || ! -f "$INPUT" ]]' is now effectively covered
# by the logic above, as we ensure INPUT is a file or exit.

# --- CATEGORY DETECTION ---
SAB_PASSED_CATEGORY="$5" # SABnzbd passes category as the 5th argument

echo "$(date) - Argument \$5 (SABnzbd Category): '$SAB_PASSED_CATEGORY'" >> "$LOG"
echo "$(date) - Environment SAB_CAT: '$SAB_CAT'" >> "$LOG"

if [[ -n "$SAB_PASSED_CATEGORY" && "$SAB_PASSED_CATEGORY" != "None" && "$SAB_PASSED_CATEGORY" != "" ]]; then
  CATEGORY="$SAB_PASSED_CATEGORY"
  echo "$(date) - Using category from argument \$5: $CATEGORY" >> "$LOG"
elif [[ -n "$SAB_CAT" && "$SAB_CAT" != "None" && "$SAB_CAT" != "" ]]; then
  CATEGORY="$SAB_CAT" # Fallback to environment variable if $5 is not useful
  echo "$(date) - Using category from environment SAB_CAT: $CATEGORY" >> "$LOG"
else
  echo "$(date) - Category not found in argument or environment. Deriving from path: $DIR" >> "$LOG"
  # Check if DIR (derived from the actual video file path) contains specific category subfolders
  if [[ "$DIR" == *"/Movies"* || "$DIR" == *"/Movie"* ]]; then
    CATEGORY="movies"
  elif [[ "$DIR" == *"/TV"* ]]; then
    CATEGORY="tv"
  else
    CATEGORY="unknown"
  fi
  echo "$(date) - Derived category from path: $CATEGORY" >> "$LOG"
fi
CATEGORY=$(echo "$CATEGORY" | awk '{print tolower($0)}') # Ensure lowercase

echo "$(date) - Final detected category: $CATEGORY" >> "$LOG"

# --- CONVERSION / FILE HANDLING LOGIC ---
OPERATION_SUCCESSFUL=false # Flag to track if we have a file ready for API calls

# Convert to lowercase for extension matching
EXT_LOWER=$(echo "$EXT" | awk '{print tolower($0)}')

if [[ "$EXT_LOWER" =~ ^(mkv|avi|ts|mov)$ ]]; then
  echo "$(date) - Starting ffmpeg conversion for $INPUT to: $TEMP_OUTPUT" >> "$LOG"
  mkdir -p "$INTERMEDIATE_DIR" # Ensure intermediate directory exists

  # --- LOUDNESS MEASUREMENT ---
  echo "$(date) - Measuring loudness for $INPUT" >> "$LOG"
  LOUDNESS_INFO=$("$SCRIPT_DIR/ffmpeg" -i "$INPUT" -hide_banner -filter_complex "[0:a:0]ebur128=metadata=1" -f null - 2>&1)
  
  # Extract Integrated Loudness (I)
  INTEGRATED_LOUDNESS=$(echo "$LOUDNESS_INFO" | grep -o 'I: \s*-[0-9\.]\+' | grep -o '-[0-9\.]\+')
  
  AUDIO_FILTER_ARGS=("-c:a" "copy") # Default to copy

  if [[ -n "$INTEGRATED_LOUDNESS" ]]; then
    echo "$(date) - Measured Integrated Loudness: $INTEGRATED_LOUDNESS LUFS" >> "$LOG"
    # Calculate gain using bc (arbitrary-precision calculator)
    GAIN=$(echo "$TARGET_LOUDNESS - $INTEGRATED_LOUDNESS" | bc)
    echo "$(date) - Calculated gain: $GAIN dB" >> "$LOG"

    # Check if absolute gain is greater than a small threshold (e.g., 0.2) to avoid unnecessary re-encoding
    if (( $(echo "$GAIN < -0.2 || $GAIN > 0.2" | bc -l) )); then
      echo "$(date) - Applying volume gain of $GAIN dB. Audio will be re-encoded." >> "$LOG"
      AUDIO_FILTER_ARGS=("-c:a" "aac" "-b:a" "192k" "-af" "volume=${GAIN}dB")
    else
      echo "$(date) - Loudness is within target range. Audio will be copied." >> "$LOG"
    fi
  else
    echo "$(date) - WARNING: Could not determine loudness. Audio will be copied without change." >> "$LOG"
  fi

  "$SCRIPT_DIR/ffmpeg" -y -loglevel debug -i "$INPUT" -c:v libx265 -pix_fmt yuv420p -avoid_negative_ts make_zero -preset medium -x265-params aq-mode=3:psy-rd=2.0:psy-rdoq=1.0:rd=3:bframes=8:ref=4:me=3:subme=6:merange=32:deblock=-1,-1:scenecut=40:keyint=240:strong-intra-smoothing=0 -crf 20 -vf scale=1280:720 "${AUDIO_FILTER_ARGS[@]}" "$TEMP_OUTPUT" >> "$LOG" 2>&1

  if [ $? -eq 0 ]; then
    echo "$(date) - Conversion successful: $TEMP_OUTPUT" >> "$LOG"
    rm -f "$INPUT" # Delete original (e.g. .mkv)
    echo "$(date) - Deleted original file: $INPUT" >> "$LOG"
    # $FINAL_OUTPUT is already defined as $DIR/$BASENAME.mp4 (original dir, new .mp4 extension)
    mv "$TEMP_OUTPUT" "$FINAL_OUTPUT" # Move converted .mp4 from intermediate to final output path for Sonarr/Radarr
    echo "$(date) - Moved completed file to: $FINAL_OUTPUT" >> "$LOG"
    OPERATION_SUCCESSFUL=true
  else
    echo "$(date) - ERROR: Conversion failed for: $INPUT. Script will exit." >> "$LOG"
    exit 1 # Critical error, stop script.
  fi
else
  echo "$(date) - Unsupported file type: $EXT for $INPUT. Script will exit. No API calls will be made." >> "$LOG"
  exit 0 # Not a processing error, but nothing further to do for Sonarr/Radarr.
fi

# --- SHORT PAUSE TO AVOID RACE CONDITION ---
echo "$(date) - Sleeping briefly before API rescan..." >> "$LOG"
sleep 5

echo "$(date) - Processing complete for: $FILENAME" >> "$LOG"
exit 0
