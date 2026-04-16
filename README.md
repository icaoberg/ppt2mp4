# PPT to Video Converter

**Version 0.4**

Converts a PowerPoint presentation (`.pptx` or `.ppt`) to an MP4 video. An AI voice reads
the presenter notes for each slide. Slides without notes are displayed silently
for a configurable duration.

Supports two TTS backends: **Microsoft Neural TTS** (free, no API key) and **ElevenLabs** (requires an API key).

## How It Works

1. **Extract notes** — reads presenter notes from each slide using `python-pptx`
2. **Render slides** — converts the deck to images via LibreOffice (PPTX → PDF) and pdf2image (PDF → PNG)
3. **Generate narration** — creates MP3 audio for each slide's notes using the chosen TTS library (`microsoft` or `elevenlabs`)
4. **Assemble video** — combines slide images and audio into a single MP4 using `moviepy`

## Requirements

### Python packages

```bash
pip install python-pptx edge-tts "moviepy<2" pdf2image pillow elevenlabs
```

### System packages

**LibreOffice** (renders slides to PDF):

| OS | Command |
|----|---------|
| Ubuntu / Debian / WSL | `sudo apt-get install libreoffice` |
| macOS | `brew install --cask libreoffice` |
| Windows | [libreoffice.org/download](https://www.libreoffice.org/download/) |

**Poppler** (converts PDF pages to images):

| OS | Command |
|----|---------|
| Ubuntu / Debian / WSL | `sudo apt-get install poppler-utils` |
| macOS | `brew install poppler` |
| Windows | [github.com/oschwartz10612/poppler-windows](https://github.com/oschwartz10612/poppler-windows) |

**Extra ffmpeg codecs** (required for `.avi` and `.wmv` output on Ubuntu/Debian):

```bash
sudo apt install libavcodec-extra
```

## Usage

```bash
python ppt2movie.py <presentation.pptx> [options]
```

The output MP4 is saved in the same directory as the input file, using the same
base filename (e.g. `presentation.pptx` → `presentation.mp4`).

### Options

| Option | Default | Description |
|--------|---------|-------------|
| `--library NAME` | `microsoft` | TTS library: `microsoft` (edge-tts, free) or `elevenlabs` (requires `ELEVENLABS_API_KEY`) |
| `--voice NAME` | see below | TTS voice name or ID (default depends on `--library`) |
| `--silent SECONDS` | `3.0` | Duration to show slides that have no presenter notes |
| `--output PATH` | same name as input | Override the output MP4 path |
| `--author NAME` | `Ivan Cao-Berg` | Author name embedded in video metadata |
| `--group NAME` | `Biomedical Applications Group` | Group name embedded in video metadata |
| `--center NAME` | `Pittsburgh Supercomputing Center` | Center name embedded in video metadata |
| `--copyright TEXT` | current year + author/institution | Copyright string embedded in video metadata |

### Examples

```bash
# Basic usage (Microsoft TTS, default)
python ppt2movie.py "Taming Data Dragons - Introduction.pptx"

# Use ElevenLabs TTS
ELEVENLABS_API_KEY=your_key python ppt2movie.py presentation.pptx --library elevenlabs

# Use ElevenLabs with a specific voice
ELEVENLABS_API_KEY=your_key python ppt2movie.py presentation.pptx --library elevenlabs --voice Adam

# Use a British male voice (Microsoft)
python ppt2movie.py presentation.pptx --voice en-GB-RyanNeural

# Show silent slides for 5 seconds
python ppt2movie.py presentation.pptx --silent 5

# Specify a custom output path
python ppt2movie.py presentation.pptx --output ~/Videos/output.mp4

# Override metadata
python ppt2movie.py presentation.pptx --author "Jane Smith" --group "Research Lab"
```

### Browsing available voices

**Microsoft** (edge-tts):
```bash
edge-tts --list-voices
```

Some popular Microsoft voices:

| Voice name | Language / Style |
|------------|-----------------|
| `en-US-ChristopherNeural` | English (US), male (default) |
| `en-US-JennyNeural` | English (US), female |
| `en-US-GuyNeural` | English (US), male |
| `en-GB-SoniaNeural` | English (UK), female |
| `en-GB-RyanNeural` | English (UK), male |
| `en-AU-NatashaNeural` | English (AU), female |

**ElevenLabs**: pass a voice name (e.g. `Rachel`, `Adam`, `Bella`) or a voice ID from your ElevenLabs account. The default is `Rachel`.

> **Note:** ElevenLabs requires the `ELEVENLABS_API_KEY` environment variable to be set.

## Adding Presenter Notes in PowerPoint

1. Open your presentation in PowerPoint.
2. Select a slide.
3. Click **View → Notes** (or click the notes panel at the bottom of the screen).
4. Type the script you want the AI voice to read for that slide.
5. Save the file.

Slides with no notes will be shown silently for the duration set by `--silent`.

## Troubleshooting

**`LibreOffice not found`**
Install LibreOffice and make sure `libreoffice` or `soffice` is on your `PATH`.

**`poppler` / `pdfinfo` not found**
Install Poppler utilities (see Requirements above).

**Audio and slide counts do not match**
This can happen if the PPTX contains hidden slides. The script automatically
truncates to the shorter of the two lists and prints a warning.

**Poor slide rendering quality**
Increase the DPI by editing `dpi=150` in `convert_pptx_to_images()` inside
`ppt2movie.py`. Higher DPI produces sharper images but takes longer.

---

Copyright (C) 2026 icaoberg, Carnegie Mellon University

This program is free software: you can redistribute it and/or modify it under
the terms of the GNU General Public License as published by the Free Software
Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY
WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A
PARTICULAR PURPOSE. See the [GNU General Public License](LICENSE) for more details.
