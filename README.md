# Green Ocean Tracker

A Windows desktop app that keeps a **Google Sheet** in sync with the stock footage you collect from YouTube and the **CapCut** projects you cut it into.

It answers two questions for a video team:

- *Which source videos do we already have, and how good are they?* — the **All Videos** worksheet.
- *Which source videos went into each finished edit?* — the **Video Projects** worksheet.

Built with Python, [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter), [gspread](https://github.com/burnash/gspread), [RapidFuzz](https://github.com/rapidfuzz/RapidFuzz) and [yt-dlp](https://github.com/yt-dlp/yt-dlp).

---

## Contents

1. [Features](#features)
2. [How it works](#how-it-works)
3. [Requirements](#requirements)
4. [Installation](#installation)
5. [Google Sheet layout](#google-sheet-layout)
6. [Google Cloud setup](#google-cloud-setup)
7. [Configuration](#configuration)
8. [Using the app](#using-the-app)
9. [Troubleshooting](#troubleshooting)
10. [Security](#security)
11. [Project structure](#project-structure)
12. [Roadmap](#roadmap)

---

## Features

| Tab | What it does |
| --- | --- |
| **Submit Video** | Pick a local CapCut project. The app reads its source clips, fuzzy-matches them to titles in *All Videos*, and logs the finished edit (name, date, YouTube links of every source used) to *Video Projects*. Optionally links the project name to the final render on Google Drive. |
| **Download** | Paste one or more YouTube links (videos or playlists). The app downloads them with yt-dlp and adds only **new** videos to *All Videos* — duplicates are skipped. |
| **Add & Rate Video** | Add a single video by hand with a location, 0–10 ratings for eight areas (Pool, Backyard, Kitchen…), comments, area similarity and who found it. If the video already exists you're asked whether to overwrite it. |

Other behaviour worth knowing:

- YouTube links are normalised (`youtu.be/…`, `&t=4s`, `&list=…` are stripped to `https://www.youtube.com/watch?v=ID`) so the same video is never stored twice under different URLs.
- Downloads and title look-ups run in the background, so the window stays responsive and the progress bar updates live.
- Submitting a project name that already exists asks before overwriting.

---

## How it works

```
CapCut project ──► draft_content.json ──► source clip filenames
                                               │  clean + fuzzy match (score > 85)
                                               ▼
YouTube ──yt-dlp──► All Videos (Title, Link, ratings…) ──► matched YouTube links
                                                               │
                                                               ▼
                                         Video Projects (Project, Date, Videos Used)
```

1. CapCut keeps an index of local projects in `root_meta_info.json`. The app reads it to fill the project dropdown.
2. Each project has a `draft_content.json`; the app takes every path under `materials → videos → path` that has a video extension (`.mp4`, `.mov`, `.mkv`, …).
3. Filenames are cleaned (`My_Video-Clip.mp4` → `my video clip`) and compared to the cleaned titles in *All Videos* column A using RapidFuzz `WRatio`. A match counts only when the score is **above 85**.
4. The YouTube links of all matches are written, comma-separated, to a new row in *Video Projects*.

> **Tip:** Clips downloaded through the **Download** tab are saved as `<YouTube title>.<ext>`, which is exactly what the matcher expects. Renaming clips before importing them into CapCut makes matching less reliable.

---

## Requirements

- Windows 10 or 11
- Python 3.10+ (developed on 3.13)
- CapCut Desktop with at least one saved project
- A Google account and a Google Cloud project (free tier is fine)

---

## Installation

```powershell
git clone https://github.com/dimslalom/Green-Ocean-Tracker.git
cd Green-Ocean-Tracker

python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

`requirements.txt` installs:

| Package | Used for |
| --- | --- |
| `customtkinter` | Dark-mode GUI |
| `gspread`, `google-auth` | Reading/writing the Google Sheet |
| `google-api-python-client` | Validating Google Drive links and fetching the file name |
| `rapidfuzz` | Fuzzy filename ↔ title matching |
| `yt-dlp` | Downloading YouTube videos and fetching titles |

> yt-dlp works best when [FFmpeg](https://ffmpeg.org/download.html) is on your `PATH` (it is needed to merge separate video/audio streams into one file).

---

## Google Sheet layout

Create one spreadsheet with **two worksheets named exactly** as below. Put the headers in row 1.

### `All Videos`

| Col | Header | Filled by |
| --- | --- | --- |
| A | Title | Download, Add & Rate |
| B | YouTube Link | Download, Add & Rate |
| C | Original Location | Add & Rate |
| D | Pool | Add & Rate (0–10) |
| E | Backyard | Add & Rate (0–10) |
| F | Overall Indoor | Add & Rate (0–10) |
| G | Kitchen | Add & Rate (0–10) |
| H | Living room | Add & Rate (0–10) |
| I | Bathrooms | Add & Rate (0–10) |
| J | Full Home Video | Add & Rate (0–10) |
| K | Outdoor Kitchen | Add & Rate (0–10) |
| L | Comments | Add & Rate |
| M | *(your own column)* | **Never written by the app** — safe for formulas such as an average score |
| N | Area Similarity | Add & Rate |
| O | Who Found It | Add & Rate |

### `Video Projects`

| Col | Header | Value |
| --- | --- | --- |
| A | Project Name | Final video name (a clickable link if a Drive link was given) |
| B | Date | Submission date, `YYYY-MM-DD` |
| C | Videos Used | Comma-separated YouTube links of matched source videos |

---

## Google Cloud setup

The app talks to Google as a **service account** (a robot Google user). One-time setup:

1. **Create a project** — go to <https://console.cloud.google.com/>, open the project dropdown → **New Project**, name it (e.g. `GreenOceanTracker`) and select it.
2. **Enable APIs** — *APIs & Services → Library*, then enable **Google Sheets API** and **Google Drive API**.
3. **Create a service account** — *APIs & Services → Credentials → + Create credentials → Service account*. Any name works; no roles are needed. Click **Done**.
4. **Create a key** — open the service account → **Keys → Add key → Create new key → JSON**. A `.json` file downloads.
5. **Install the key** — move that file next to `app.py` and rename it to `credentials.json`.
6. **Share the sheet** — copy the service account's email (`name@project-id.iam.gserviceaccount.com`), open your spreadsheet → **Share**, paste the email and give it **Editor** access.

> Skipping step 6 is the most common cause of `SpreadsheetNotFound`.

---

## Configuration

Settings live at the top of [`app.py`](app.py). The two most common ones can also be set with environment variables, so you don't need to edit code:

| Setting | Env var | Default |
| --- | --- | --- |
| Spreadsheet name | `GOOGLE_SHEET_NAME` | `Capcut Video Tracker Test` |
| CapCut project index | `CAPCUT_ROOT_META_PATH` | `%LOCALAPPDATA%\CapCut\User Data\Projects\com.lveditor.draft\root_meta_info.json` |
| Service account key | — | `credentials.json` next to `app.py` |
| Match threshold | — | `85` (`MATCH_THRESHOLD`) |
| Rating areas | — | `RATING_AREAS` (must match columns D–K) |

Example (PowerShell, current session only):

```powershell
$env:GOOGLE_SHEET_NAME = "Green Ocean Video Tracker"
python app.py
```

If you moved CapCut's drafts folder (CapCut → Settings → *Projects location*), point `CAPCUT_ROOT_META_PATH` at the `root_meta_info.json` inside that folder.

---

## Using the app

```powershell
python app.py
```

### Submit Video

1. Choose the CapCut project from the dropdown.
2. *(Optional)* Paste the Google Drive link of the exported video and click **Validate & Fetch Title**. The app checks that it can open the file and fills in **Final Video Name** from the file name. If you get *Access Denied*, set the Drive file's sharing to **Anyone with the link** (or share it with the service account).
3. Enter or adjust **Final Video Name**.
4. Click **Submit to Database**.

If none of the project's clips match a title in *All Videos*, nothing is written — add the source videos first (via **Download** or **Add & Rate**).

### Download

1. Paste YouTube links — one per line, or separated by spaces/commas. Playlist links download every video in the playlist.
2. Choose a download folder (default `Downloads\CapCutDownloads`).
3. Click **Download + Add to Database**.

A video is treated as a duplicate if its normalised link **or** its normalised title already exists in *All Videos*.

### Add & Rate Video

1. Paste a YouTube URL. The title is fetched automatically when you leave the field (or click **Fetch Title**).
2. Fill in location, ratings, comments, area similarity and who found it.
3. Click **Submit to Database**. If the video already exists you can choose to overwrite its row.

---

## Troubleshooting

| Message | Fix |
| --- | --- |
| `CapCut root_meta_info.json not found` | Open CapCut and save a project at least once, or set `CAPCUT_ROOT_META_PATH`. |
| `credentials.json not found` | Put the key next to `app.py` with exactly that name. |
| `Google Sheet not found` | Check `GOOGLE_SHEET_NAME` matches the file name exactly, the sheet is shared with the service account, and both APIs are enabled. |
| `Worksheet missing` | Worksheets must be named `All Videos` and `Video Projects`. |
| `No matching videos found in the database` | The project's clip filenames don't resemble any title. Add the videos to *All Videos* first, or rename clips to match. |
| `google-api-python-client is not available` | `pip install -r requirements.txt` — needed for Drive link validation. |
| `yt-dlp is not installed` | `pip install -r requirements.txt`, then restart the app. |
| Downloads fail or have no audio | Update yt-dlp (`pip install -U yt-dlp`) and install FFmpeg. |
| `Access Denied` on a Drive link | Share the Drive file as *Anyone with the link*, or with the service account email. |

---

## Security

- `credentials.json` grants **edit access to every sheet shared with the service account**. It is listed in `.gitignore`; never commit or share it.
- If a key is ever exposed (pushed to GitHub, sent in chat, etc.), go to the service account → **Keys**, **delete** that key and create a new one. Removing the file from the repository does not remove it from git history.

---

## Project structure

```
Green Ocean Tracker/
├── app.py             # The whole application (helpers, SheetsClient, GUI)
├── requirements.txt   # Python dependencies
├── credentials.json   # Your service account key (git-ignored, you add this)
├── .gitignore
└── README.md
```

Inside `app.py`:

- **Helpers** — `normalize_text`, `clean_youtube_url`, `extract_drive_file_id`, `load_capcut_projects`, `extract_clean_video_names_from_draft`.
- **`SheetsClient`** — all Google Sheets reads/writes (matching, duplicate checks, appends, overwrites).
- **`CapCutTrackerApp`** — the CustomTkinter window and the three tabs.

---

## Roadmap

- Browse for a custom CapCut projects folder from the UI
- Show a per-clip match report (which clip matched which title, and the score)
- Configurable match threshold in the UI
- Write a log file for an audit trail
- Package as a single `.exe` with PyInstaller
