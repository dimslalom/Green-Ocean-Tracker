"""
CapCut → Google Sheets Tracker (Proof of Concept)

What this app does (three tabs):
1) Submit Video
   - Lists local CapCut projects from CapCut's root project index.
   - Parses the selected project's draft_content.json to find source video files.
   - Fuzzy-matches those filenames against titles in worksheet "All Videos".
   - Writes (or overwrites) a summary row in worksheet "Video Projects", optionally
     hyperlinked to the final render on Google Drive.
2) Download
   - Downloads YouTube videos with yt-dlp and inserts only new ones into "All Videos".
3) Add & Rate Video
   - Manually adds (or overwrites) a video in "All Videos" with per-area ratings.

Tech stack:
- GUI: customtkinter
- Google Sheets/Drive: gspread + google-auth service account (+ google-api-python-client)
- Matching: rapidfuzz
- Downloading: yt-dlp

IMPORTANT:
- Place credentials.json in the same folder as this app.py (or change CREDENTIALS_PATH below).
- Set GOOGLE_SHEET_NAME below (or the GOOGLE_SHEET_NAME env var) to your real spreadsheet name.
"""

from __future__ import annotations

import importlib
import json
import os
import re
import threading
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple
from urllib.parse import parse_qs, urlparse

import customtkinter as ctk
import gspread
try:
    from googleapiclient.discovery import build  # type: ignore[import-not-found]
    from googleapiclient.errors import HttpError  # type: ignore[import-not-found]
except Exception:  # pragma: no cover - optional dependency for Drive validation
    build = None
    HttpError = Exception
from google.oauth2.service_account import Credentials
from gspread.utils import ValueInputOption
from rapidfuzz import fuzz, process


# ==============================
# Configuration (edit as needed)
# ==============================

# Name of your Google Spreadsheet (the file name shown in Google Sheets).
# You can hardcode it here, or set environment variable GOOGLE_SHEET_NAME.
GOOGLE_SHEET_NAME = os.getenv("GOOGLE_SHEET_NAME", "Capcut Video Tracker Test")

# Service account credentials file. By default: same folder as app.py.
CREDENTIALS_PATH = Path(__file__).resolve().parent / "credentials.json"

# CapCut master project index. Override with env var CAPCUT_ROOT_META_PATH if CapCut
# stores projects somewhere else (e.g. a custom drafts folder).
CAPCUT_ROOT_META_PATH = Path(
    os.getenv(
        "CAPCUT_ROOT_META_PATH",
        str(Path.home() / "AppData" / "Local" / "CapCut" / "User Data" / "Projects" / "com.lveditor.draft" / "root_meta_info.json"),
    )
)

# Fuzzy match threshold. Match is accepted only if score > 85.
MATCH_THRESHOLD = 85

# Extensions to keep as video files.
VIDEO_EXTENSIONS = {
    ".mp4", ".mov", ".mkv", ".avi", ".wmv", ".flv", ".webm", ".m4v", ".mpeg", ".mpg"
}

# Rating areas shown in the "Add & Rate Video" tab, in "All Videos" column order (D-K).
RATING_AREAS = [
    "Pool",
    "Backyard",
    "Overall Indoor",
    "Kitchen",
    "Living room",
    "Bathrooms",
    "Full Home Video",
    "Outdoor Kitchen",
]

# Status label colors.
COLOR_INFO = "#A0A0A0"
COLOR_OK = "#7EE787"
COLOR_WARN = "#FFB86C"
COLOR_ERROR = "#FF6B6B"

# Google API scopes needed for gspread service-account access.
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ======================
# Data model definitions
# ======================

@dataclass
class CapCutProject:
    """Simple model for a CapCut project entry from root_meta_info.json."""

    name: str
    folder_path: Path


# ==========================
# Utility/helper core methods
# ==========================

def read_json_file(json_path: Path) -> dict:
    """
    Read and parse a JSON file safely.

    Raises:
        FileNotFoundError: if file does not exist.
        ValueError: if JSON is invalid.
        OSError: for low-level file read issues.
    """
    if not json_path.exists():
        raise FileNotFoundError(f"JSON file not found: {json_path}")

    try:
        with json_path.open("r", encoding="utf-8") as file:
            return json.load(file)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON in file: {json_path}") from exc


def normalize_text(text: str) -> str:
    """
    Normalize a string for fuzzy matching:
    - lower-case
    - replace underscores/dashes with spaces
    - strip extension if present
    - keep letters/numbers/spaces only
    - collapse multiple spaces

    Example:
        "My_Video-Clip.mp4" -> "my video clip"
    """
    if not text:
        return ""

    # Remove extension if present.
    no_ext = Path(text).stem

    # Replace common formatting separators.
    replaced = no_ext.replace("_", " ").replace("-", " ")

    # Keep only alphanumeric + spaces.
    cleaned = re.sub(r"[^a-zA-Z0-9\s]", " ", replaced)

    # Lowercase + collapse whitespace.
    return re.sub(r"\s+", " ", cleaned).strip().lower()


def clean_youtube_url(raw_url: str) -> str:
    """
    Strip tracking parameters and normalize YouTube URLs.

    Examples:
        https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=4s&list=WL
        -> https://www.youtube.com/watch?v=dQw4w9WgXcQ
    """
    if not raw_url:
        return ""

    url = raw_url.strip()
    parsed = urlparse(url)

    if not parsed.scheme or not parsed.netloc:
        return url

    host = parsed.netloc.lower()
    path = parsed.path or ""

    if "youtu.be" in host:
        video_id = path.strip("/")
        if video_id:
            return f"https://www.youtube.com/watch?v={video_id}"
        return "https://www.youtube.com"

    if "youtube.com" in host and path.startswith("/watch"):
        params = parse_qs(parsed.query)
        video_id = (params.get("v") or [""])[0]
        if video_id:
            return f"https://www.youtube.com/watch?v={video_id}"

    # Fallback: drop query + fragment but keep the base path.
    return f"{parsed.scheme}://{parsed.netloc}{path}"


def extract_drive_file_id(link: str) -> str:
    """
    Extract Google Drive file ID from common link formats.
    """
    if not link:
        return ""

    patterns = [
        r"/d/([a-zA-Z0-9_-]+)",
        r"id=([a-zA-Z0-9_-]+)",
    ]

    for pattern in patterns:
        match = re.search(pattern, link)
        if match:
            return match.group(1)

    return ""


def build_project_cell(final_video_name: str, gdrive_link: str) -> str:
    """
    Build the "Project Name" cell value: plain text, or a HYPERLINK formula to the Drive file.
    Double quotes are escaped so names/links containing quotes do not break the formula.
    """
    if not gdrive_link:
        return final_video_name

    safe_link = gdrive_link.replace('"', '""')
    safe_name = final_video_name.replace('"', '""')
    return f'=HYPERLINK("{safe_link}", "{safe_name}")'


def is_video_file(file_path: str) -> bool:
    """
    Return True if path has a known video extension.
    Empty or malformed path returns False.
    """
    if not file_path:
        return False

    extension = Path(file_path).suffix.lower()
    return extension in VIDEO_EXTENSIONS


def find_draft_entries_recursive(node: object, found: Optional[List[CapCutProject]] = None) -> List[CapCutProject]:
    """
    Recursively traverse JSON data to find objects containing:
    - draft_name
    - draft_fold_path

    This is intentionally defensive because CapCut JSON structure can vary by version.
    """
    if found is None:
        found = []

    if isinstance(node, dict):
        # If this dictionary looks like a project entry, capture it.
        draft_name = node.get("draft_name")
        draft_fold_path = node.get("draft_fold_path")
        if isinstance(draft_name, str) and isinstance(draft_fold_path, str):
            name = draft_name.strip()
            folder = draft_fold_path.strip()
            if name and folder:
                found.append(CapCutProject(name=name, folder_path=Path(folder)))

        # Continue recursion for all values.
        for value in node.values():
            find_draft_entries_recursive(value, found)

    elif isinstance(node, list):
        # Continue recursion for all list items.
        for item in node:
            find_draft_entries_recursive(item, found)

    return found


def load_capcut_projects(root_meta_path: Path = CAPCUT_ROOT_META_PATH) -> List[CapCutProject]:
    """
    Read CapCut root_meta_info.json and extract project names/folder paths.

    Returns a deduplicated list by (name, folder_path).
    """
    data = read_json_file(root_meta_path)
    raw_projects = find_draft_entries_recursive(data)

    # Deduplicate while preserving order.
    seen: set[Tuple[str, str]] = set()
    unique_projects: List[CapCutProject] = []
    for project in raw_projects:
        key = (project.name, str(project.folder_path))
        if key not in seen:
            seen.add(key)
            unique_projects.append(project)

    return unique_projects


def get_draft_content_path(project_folder: Path) -> Path:
    """
    Build path to a project's draft_content.json from draft folder path.
    """
    return project_folder / "draft_content.json"


def extract_clean_video_names_from_draft(draft_content_path: Path) -> List[str]:
    """
    Parse draft_content.json and extract cleaned video names from:
    ["materials"]["videos"][*]["path"]

    Filtering rules:
    - Skip empty/non-string paths
    - Skip non-video extensions
    - Skip values that normalize to empty text

    Returns:
        List of normalized filename tokens for fuzzy matching.
    """
    data = read_json_file(draft_content_path)

    materials = data.get("materials", {}) if isinstance(data, dict) else {}
    videos = materials.get("videos", []) if isinstance(materials, dict) else []

    if not isinstance(videos, list):
        return []

    cleaned_names: List[str] = []

    for video_item in videos:
        if not isinstance(video_item, dict):
            continue

        raw_path = video_item.get("path")
        if not isinstance(raw_path, str) or not raw_path.strip():
            continue

        # Keep only true video files.
        if not is_video_file(raw_path):
            continue

        base_name = Path(raw_path).name
        cleaned = normalize_text(base_name)
        if cleaned:
            cleaned_names.append(cleaned)

    return cleaned_names


# ======================
# Google Sheets handling
# ======================

class SheetsClient:
    """
    Small wrapper around gspread operations used by this app.

    Expected worksheet layout:
    1) "All Videos"      -> Title (A), YouTube Link (B), Rating (C), Location (D), Used In (E)
    2) "Video Projects"  -> Project Name (A), Date (B), Videos Used (C)
    """

    def __init__(self, credentials_path: Path, spreadsheet_name: str) -> None:
        self.credentials_path = credentials_path
        self.spreadsheet_name = spreadsheet_name
        self.client: Optional[gspread.Client] = None
        self.spreadsheet: Optional[gspread.Spreadsheet] = None
        self.ws_all_videos: Optional[gspread.Worksheet] = None
        self.ws_video_projects: Optional[gspread.Worksheet] = None

    def connect(self) -> None:
        """Authenticate and open required worksheets."""
        if not self.credentials_path.exists():
            raise FileNotFoundError(
                f"credentials.json not found at: {self.credentials_path}. "
                "Place credentials.json next to app.py or update CREDENTIALS_PATH."
            )

        if not self.spreadsheet_name or self.spreadsheet_name == "YOUR_GOOGLE_SHEET_NAME":
            raise ValueError(
                "GOOGLE_SHEET_NAME is not configured. "
                "Set it in code or as environment variable GOOGLE_SHEET_NAME."
            )

        credentials = Credentials.from_service_account_file(
            str(self.credentials_path),
            scopes=GOOGLE_SCOPES,
        )

        self.client = gspread.authorize(credentials)
        self.spreadsheet = self.client.open(self.spreadsheet_name)
        self.ws_all_videos = self.spreadsheet.worksheet("All Videos")
        self.ws_video_projects = self.spreadsheet.worksheet("Video Projects")

    def match_youtube_links(
        self,
        cleaned_video_names: Sequence[str],
        score_threshold: int = MATCH_THRESHOLD,
    ) -> List[str]:
        """
        Match CapCut cleaned filenames against column A (Title) in "All Videos".
        Only collects YouTube Link from column B for logging; "All Videos" is not modified.

        Returns:
            List of matched YouTube links (deduplicated, insertion order).
        """
        if self.ws_all_videos is None:
            raise RuntimeError("Worksheet 'All Videos' is not connected.")

        all_values = self.ws_all_videos.get_all_values()
        if not all_values or len(all_values) < 2:
            # Sheet has only header or is empty.
            return []

        # Build row models from sheet values.
        # Row indexing in gspread starts at 1; first data row starts at row 2.
        row_models = []
        normalized_titles = []

        for sheet_row_index, row in enumerate(all_values[1:], start=2):
            title = row[0].strip() if len(row) > 0 else ""
            link = row[1].strip() if len(row) > 1 else ""
            title_normalized = normalize_text(title)

            # Skip rows with empty title.
            if not title_normalized:
                continue

            row_models.append(
                {
                    "sheet_row": sheet_row_index,
                    "title": title,
                    "title_norm": title_normalized,
                    "link": link,
                    "used_in": "",
                }
            )
            normalized_titles.append(title_normalized)

        if not row_models:
            return []

        collected_links: List[str] = []

        for video_name in cleaned_video_names:
            if not video_name:
                continue

            # Use rapidfuzz to find best title match.
            # Returns tuple: (matched_string, score, matched_index)
            match = process.extractOne(
                query=video_name,
                choices=normalized_titles,
                scorer=fuzz.WRatio,
            )

            if not match:
                continue

            _, score, matched_index = match
            if score <= score_threshold:
                # Requirement says score must be > 85, so <= threshold is rejected.
                continue

            target_row = row_models[matched_index]

            # Collect link for final summary if available.
            link_value = clean_youtube_url(target_row["link"])
            if link_value and link_value not in collected_links:
                collected_links.append(link_value)

        return collected_links

    def append_video_project_log(self, final_video_name: str, youtube_links: Sequence[str], gdrive_link: str) -> None:
        """
        Append one new row to worksheet "Video Projects":
        [Project Name, Date, Videos Used]
        """
        if self.ws_video_projects is None:
            raise RuntimeError("Worksheet 'Video Projects' is not connected.")

        links_text = ", ".join(youtube_links)
        today = date.today().strftime("%Y-%m-%d")
        project_cell_value = build_project_cell(final_video_name, gdrive_link)

        self.ws_video_projects.append_row(
            [project_cell_value, today, links_text],
            value_input_option=ValueInputOption.user_entered,
        )

    def find_project_row(self, project_name: str) -> Optional[int]:
        """
        Find a project row by exact match in column A of "Video Projects".
        Returns 1-indexed row number, or None if not found.
        """
        if self.ws_video_projects is None:
            raise RuntimeError("Worksheet 'Video Projects' is not connected.")

        column_values = self.ws_video_projects.col_values(1)
        for index, value in enumerate(column_values, start=1):
            text_value = str(value).strip()
            if index == 1 and text_value.lower() == "project name":
                continue
            if text_value == project_name:
                return index

        return None

    def update_video_project_log(
        self,
        row_index: int,
        final_video_name: str,
        youtube_links: Sequence[str],
        gdrive_link: str,
    ) -> None:
        """
        Overwrite an existing row in "Video Projects" at row_index.
        """
        if self.ws_video_projects is None:
            raise RuntimeError("Worksheet 'Video Projects' is not connected.")

        links_text = ", ".join(youtube_links)
        today = date.today().strftime("%Y-%m-%d")
        project_cell_value = build_project_cell(final_video_name, gdrive_link)

        self.ws_video_projects.update(
            range_name=f"A{row_index}:C{row_index}",
            values=[[project_cell_value, today, links_text]],
            value_input_option=ValueInputOption.user_entered,
        )

    def find_video_row(self, title: str, link: str) -> Optional[int]:
        """
        Find a video row by normalized title (column A) OR cleaned link (column B).
        Returns 1-indexed row number, or None if not found.
        """
        if self.ws_all_videos is None:
            raise RuntimeError("Worksheet 'All Videos' is not connected.")

        title_norm = normalize_text(title)
        link_clean = clean_youtube_url(link)

        all_values = self.ws_all_videos.get_all_values()
        for index, row in enumerate(all_values[1:], start=2):
            row_title = row[0].strip() if len(row) > 0 else ""
            row_link = row[1].strip() if len(row) > 1 else ""

            if title_norm and normalize_text(row_title) == title_norm:
                return index
            if link_clean and clean_youtube_url(row_link) == link_clean:
                return index

        return None

    def update_manual_video_row(
        self,
        row_index: int,
        row_values: Sequence[str],
        area_similarity: str,
        who_found_it: str,
    ) -> None:
        """
        Overwrite an existing "All Videos" row: A-L and N-O (skip M).
        """
        if self.ws_all_videos is None:
            raise RuntimeError("Worksheet 'All Videos' is not connected.")

        if len(row_values) != 12:
            raise ValueError("Expected 12 columns (A-L) for All Videos row insertion.")

        self.ws_all_videos.update(
            range_name=f"A{row_index}:L{row_index}",
            values=[list(row_values)],
            value_input_option=ValueInputOption.raw,
        )

        self.ws_all_videos.update(
            range_name=f"N{row_index}:O{row_index}",
            values=[[area_similarity, who_found_it]],
            value_input_option=ValueInputOption.raw,
        )

    def append_new_videos_to_all_videos(self, videos: Sequence[Dict[str, str]]) -> Tuple[int, int]:
        """
        Insert only new videos into worksheet "All Videos".

        Duplicate rules:
        - Duplicate if YouTube Link already exists in column B, OR
        - Duplicate if normalized Title already exists in column A.

        New row format (first 12 columns only; column M is intentionally untouched):
        [
            Title,
            Link,
            Original Location,
            Pool,
            Backyard,
            Overall Indoor,
            Kitchen,
            Living room,
            Bathrooms,
            Full Home Video,
            Outdoor Kitchen,
            Comments,
        ]

        Returns:
            (inserted_count, skipped_duplicate_count)
        """
        if self.ws_all_videos is None:
            raise RuntimeError("Worksheet 'All Videos' is not connected.")

        all_values = self.ws_all_videos.get_all_values()

        existing_links: set[str] = set()
        existing_titles_normalized: set[str] = set()

        if all_values:
            for row in all_values[1:]:
                title = row[0].strip() if len(row) > 0 else ""
                link = clean_youtube_url(row[1].strip()) if len(row) > 1 else ""

                title_norm = normalize_text(title)
                if title_norm:
                    existing_titles_normalized.add(title_norm)

                if link:
                    existing_links.add(link.casefold())

        rows_to_append: List[List[str]] = []
        skipped_duplicates = 0

        for video in videos:
            title = (video.get("title") or "").strip()
            link = clean_youtube_url((video.get("youtube_link") or "").strip())

            title_norm = normalize_text(title)
            link_norm = link.casefold() if link else ""

            if not title and not link:
                skipped_duplicates += 1
                continue

            is_duplicate = (
                (bool(link_norm) and link_norm in existing_links)
                or (bool(title_norm) and title_norm in existing_titles_normalized)
            )

            if is_duplicate:
                skipped_duplicates += 1
                continue

            rows_to_append.append(
                [
                    title,
                    link,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                ]
            )

            if link_norm:
                existing_links.add(link_norm)
            if title_norm:
                existing_titles_normalized.add(title_norm)

        if rows_to_append:
            self.ws_all_videos.append_rows(
                rows_to_append,
                value_input_option=ValueInputOption.raw,
            )

        return len(rows_to_append), skipped_duplicates

    def append_manual_video_row(self, row_values: Sequence[str], area_similarity: str, who_found_it: str) -> None:
        """
        Append a single row into worksheet "All Videos".
        Column M is intentionally left untouched.
        """
        if self.ws_all_videos is None:
            raise RuntimeError("Worksheet 'All Videos' is not connected.")

        if len(row_values) != 12:
            raise ValueError("Expected 12 columns (A-L) for All Videos row insertion.")

        response = self.ws_all_videos.append_row(
            list(row_values),
            value_input_option=ValueInputOption.raw,
        )

        # Use the row the API actually wrote to (e.g. "'All Videos'!A42:L42"), rather than
        # counting rows, which breaks if the sheet has blank or trailing rows.
        updated_range = str((response or {}).get("updates", {}).get("updatedRange", ""))
        row_match = re.search(r"![A-Z]+(\d+)", updated_range)
        row_number = int(row_match.group(1)) if row_match else len(self.ws_all_videos.get_all_values())
        if row_number:
            self.ws_all_videos.update(
                range_name=f"N{row_number}:O{row_number}",
                values=[[area_similarity, who_found_it]],
                value_input_option=ValueInputOption.raw,
            )


# =================
# GUI implementation
# =================

class CapCutTrackerApp(ctk.CTk):
    """Main CustomTkinter window for the proof-of-concept app."""

    def __init__(self) -> None:
        super().__init__()

        # Window configuration.
        self.title("CapCut YouTube Tracker (POC)")
        self.geometry("920x760")
        self.resizable(False, False)

        # Internal state.
        self.projects: List[CapCutProject] = []
        self.project_lookup: Dict[str, CapCutProject] = {}
        self.credentials: Optional[Credentials] = None

        # Build UI and load project list.
        self._build_ui()
        self._load_projects_into_dropdown()

    def _build_ui(self) -> None:
        """Create all widgets with a clean, dark-mode layout and two functional tabs."""
        # Global appearance settings.
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Main wrapper frame for consistent spacing.
        container = ctk.CTkFrame(self, corner_radius=12)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        # Header.
        title_label = ctk.CTkLabel(
            container,
            text="CapCut Project to Google Sheets Tracker",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        title_label.pack(pady=(20, 6))

        subtitle_label = ctk.CTkLabel(
            container,
            text="Track CapCut usage and download new YouTube videos into your database.",
            font=ctk.CTkFont(size=13),
        )
        subtitle_label.pack(pady=(0, 18))

        # Tab container.
        self.tabview = ctk.CTkTabview(container, width=760, height=500)
        self.tabview.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        tracker_tab = self.tabview.add("Submit Video")
        download_tab = self.tabview.add("Download")
        add_rate_tab = self.tabview.add("Add & Rate Video")

        # ------------------------------
        # Submit Video tab
        # ------------------------------
        project_label = ctk.CTkLabel(tracker_tab, text="CapCut Project", anchor="w")
        project_label.pack(fill="x", padx=24)

        self.project_combobox = ctk.CTkComboBox(
            tracker_tab,
            values=["Loading projects..."],
            width=680,
            state="readonly",
        )
        self.project_combobox.pack(padx=24, pady=(6, 16))
        self.project_combobox.set("Loading projects...")

        self.gdrive_link_entry = ctk.CTkEntry(
            tracker_tab,
            width=680,
            placeholder_text="Paste Google Drive Link here",
        )
        self.gdrive_link_entry.pack(padx=24, pady=(0, 10))

        self.fetch_drive_button = ctk.CTkButton(
            tracker_tab,
            text="Validate & Fetch Title",
            width=680,
            height=32,
            command=self.on_fetch_drive_clicked,
        )
        self.fetch_drive_button.pack(padx=24, pady=(0, 10))

        # Final video name label + entry.
        final_name_label = ctk.CTkLabel(tracker_tab, text="Final Video Name", anchor="w")
        final_name_label.pack(fill="x", padx=24)

        self.final_name_entry = ctk.CTkEntry(
            tracker_tab,
            width=680,
            placeholder_text="Example: Final_Dimas",
        )
        self.final_name_entry.pack(padx=24, pady=(6, 20))

        # Primary action button.
        self.update_button = ctk.CTkButton(
            tracker_tab,
            text="Submit to Database",
            width=680,
            height=40,
            command=self.on_update_database_clicked,
        )
        self.update_button.pack(padx=24, pady=(0, 14))

        # Status label for success/error feedback.
        self.status_label = ctk.CTkLabel(
            tracker_tab,
            text="Ready.",
            text_color=COLOR_INFO,
            wraplength=680,
            justify="left",
            anchor="w",
        )
        self.status_label.pack(fill="x", padx=24, pady=(0, 16))

        # ------------------------------
        # Download tab (new behavior)
        # ------------------------------
        urls_label = ctk.CTkLabel(download_tab, text="YouTube URL(s)", anchor="w")
        urls_label.pack(fill="x", padx=24, pady=(6, 0))

        self.urls_textbox = ctk.CTkTextbox(download_tab, width=680, height=170)
        self.urls_textbox.pack(padx=24, pady=(6, 16))

        urls_hint = ctk.CTkLabel(
            download_tab,
            text="Enter one URL per line (or paste multiple URLs separated by commas/spaces).",
            text_color=COLOR_INFO,
            anchor="w",
        )
        urls_hint.pack(fill="x", padx=24, pady=(0, 10))

        folder_label = ctk.CTkLabel(download_tab, text="Download Folder", anchor="w")
        folder_label.pack(fill="x", padx=24)

        folder_row = ctk.CTkFrame(download_tab, fg_color="transparent")
        folder_row.pack(fill="x", padx=24, pady=(6, 14))

        default_download_dir = Path.home() / "Downloads" / "CapCutDownloads"
        self.download_folder_entry = ctk.CTkEntry(folder_row, width=560)
        self.download_folder_entry.pack(side="left", padx=(0, 10))
        self.download_folder_entry.insert(0, str(default_download_dir))

        browse_button = ctk.CTkButton(
            folder_row,
            text="Browse",
            width=110,
            command=self.on_browse_download_folder,
        )
        browse_button.pack(side="left")

        self.download_button = ctk.CTkButton(
            download_tab,
            text="Download + Add to Database",
            width=680,
            height=40,
            command=self.on_download_and_add_clicked,
        )
        self.download_button.pack(padx=24, pady=(0, 10))

        self.download_progress = ctk.CTkProgressBar(download_tab, width=680)
        self.download_progress.pack(padx=24, pady=(0, 12))
        self.download_progress.set(0)

        self.download_status_label = ctk.CTkLabel(
            download_tab,
            text="Ready.",
            text_color=COLOR_INFO,
            wraplength=680,
            justify="left",
            anchor="w",
        )
        self.download_status_label.pack(fill="x", padx=24, pady=(0, 16))

        # ------------------------------
        # Add & Rate Video tab
        # ------------------------------
        add_rate_frame = ctk.CTkScrollableFrame(add_rate_tab, width=720, height=460)
        add_rate_frame.pack(fill="both", expand=True, padx=24, pady=(6, 12))

        add_title_label = ctk.CTkLabel(add_rate_frame, text="Manually Add & Rate Video", anchor="w")
        add_title_label.pack(fill="x", padx=8, pady=(6, 8))

        url_row = ctk.CTkFrame(add_rate_frame, fg_color="transparent")
        url_row.pack(fill="x", padx=8, pady=(0, 10))

        self.manual_url_entry = ctk.CTkEntry(url_row, width=520, placeholder_text="YouTube URL")
        self.manual_url_entry.pack(side="left", padx=(0, 10))
        self.manual_url_entry.bind("<FocusOut>", lambda _event: self.on_fetch_title_clicked())

        fetch_title_button = ctk.CTkButton(
            url_row,
            text="Fetch Title",
            width=130,
            command=self.on_fetch_title_clicked,
        )
        fetch_title_button.pack(side="left")

        self.manual_title_entry = ctk.CTkEntry(add_rate_frame, width=680, placeholder_text="Title")
        self.manual_title_entry.pack(padx=8, pady=(0, 12))

        self.manual_location_entry = ctk.CTkEntry(add_rate_frame, width=680, placeholder_text="Original Location")
        self.manual_location_entry.pack(padx=8, pady=(0, 12))

        rating_options = [str(value) for value in range(0, 11)]

        rating_frame = ctk.CTkFrame(add_rate_frame, fg_color="transparent")
        rating_frame.pack(fill="x", padx=8, pady=(0, 12))
        rating_frame.grid_columnconfigure((0, 1), weight=1)

        # One 0-10 rating dropdown per area, in sheet column order D-K.
        self.rating_combos: List[ctk.CTkComboBox] = []
        for row_index, area_label in enumerate(RATING_AREAS):
            combo = ctk.CTkComboBox(rating_frame, values=rating_options, width=330)
            combo.set("0")
            combo.grid(row=row_index, column=0, padx=(0, 10), pady=(0, 10), sticky="w")
            ctk.CTkLabel(rating_frame, text=area_label, anchor="w").grid(row=row_index, column=1, sticky="w")
            self.rating_combos.append(combo)

        self.manual_comments_entry = ctk.CTkEntry(add_rate_frame, width=680, placeholder_text="Comments")
        self.manual_comments_entry.pack(padx=8, pady=(0, 10))

        self.manual_area_entry = ctk.CTkEntry(add_rate_frame, width=680, placeholder_text="Area Similarity")
        self.manual_area_entry.pack(padx=8, pady=(0, 10))

        self.manual_found_entry = ctk.CTkEntry(add_rate_frame, width=680, placeholder_text="Who Found it")
        self.manual_found_entry.pack(padx=8, pady=(0, 12))

        self.manual_submit_button = ctk.CTkButton(
            add_rate_frame,
            text="Submit to Database",
            width=680,
            height=40,
            command=self.on_submit_manual_video,
        )
        self.manual_submit_button.pack(padx=8, pady=(0, 10))

        self.manual_status_label = ctk.CTkLabel(
            add_rate_frame,
            text="Ready.",
            text_color=COLOR_INFO,
            wraplength=680,
            justify="left",
            anchor="w",
        )
        self.manual_status_label.pack(fill="x", padx=8, pady=(0, 16))

    def set_status(self, text: str, color: str = COLOR_INFO) -> None:
        """Update status for the Tracker tab."""
        self.status_label.configure(text=text, text_color=color)

    def set_download_status(self, text: str, color: str = COLOR_INFO) -> None:
        """Update status for the Download tab."""
        self.download_status_label.configure(text=text, text_color=color)

    def set_manual_status(self, text: str, color: str = COLOR_INFO) -> None:
        """Update status for the Add & Rate Video tab."""
        self.manual_status_label.configure(text=text, text_color=color)

    def set_download_progress(self, value: float) -> None:
        """Update the download progress bar."""
        clamped = max(0.0, min(1.0, value))
        self.download_progress.set(clamped)
        self.update_idletasks()

    def _run_in_background(self, work: Callable[[], Any], on_done: Callable[[Any, Optional[BaseException]], None]) -> None:
        """
        Run slow work (downloads, network calls) on a worker thread so the window stays responsive.
        on_done(result, error) is always called back on the Tk main thread.
        """

        def worker() -> None:
            try:
                result = work()
            except BaseException as exc:  # noqa: BLE001 - surfaced to the UI via on_done
                self.after(0, on_done, None, exc)
            else:
                self.after(0, on_done, result, None)

        threading.Thread(target=worker, daemon=True).start()

    def _get_service_credentials(self) -> Credentials:
        """Load and cache service account credentials for Google APIs."""
        if self.credentials is not None:
            return self.credentials

        if not CREDENTIALS_PATH.exists():
            raise FileNotFoundError(
                f"credentials.json not found at: {CREDENTIALS_PATH}. "
                "Place credentials.json next to app.py or update CREDENTIALS_PATH."
            )

        self.credentials = Credentials.from_service_account_file(
            str(CREDENTIALS_PATH),
            scopes=GOOGLE_SCOPES,
        )
        return self.credentials

    def on_fetch_drive_clicked(self) -> None:
        """
        Validate Drive link permissions and auto-fill the final video name.
        """
        link = self.gdrive_link_entry.get().strip()
        if not link:
            self.set_status("Please paste a Google Drive link.", color=COLOR_WARN)
            return

        file_id = extract_drive_file_id(link)
        if not file_id:
            self.set_status("Could not extract a Google Drive file ID.", color=COLOR_WARN)
            return

        try:
            if build is None:
                raise RuntimeError("google-api-python-client is not available.")

            credentials = self._get_service_credentials()
            service = build("drive", "v3", credentials=credentials)
            metadata = service.files().get(fileId=file_id, fields="name").execute()
            name = str(metadata.get("name") or "").strip()

            # Drop the file extension (e.g. "Final_Dimas.mp4" -> "Final_Dimas").
            if is_video_file(name):
                name = Path(name).stem

            if name:
                self.final_name_entry.delete(0, "end")
                self.final_name_entry.insert(0, name)
                self.set_status("Drive link validated. Title fetched.", color=COLOR_OK)
            else:
                self.set_status("Drive link validated, but title is empty.", color=COLOR_WARN)

        except HttpError:
            messagebox.showwarning(
                "Access Denied",
                "⚠️ ACCESS DENIED! This Google Drive link is not public. "
                "You MUST change the sharing settings to 'Anyone with the link' before submitting.",
            )
            self.set_status("Drive access denied. Update sharing settings.", color=COLOR_ERROR)
        except Exception as exc:
            self.set_status(f"Failed to validate Drive link: {exc}", color=COLOR_ERROR)

    def on_browse_download_folder(self) -> None:
        """Open directory picker and fill download target folder."""
        selected_folder = filedialog.askdirectory(title="Select download folder")
        if selected_folder:
            self.download_folder_entry.delete(0, "end")
            self.download_folder_entry.insert(0, selected_folder)

    def _parse_urls_from_text(self, raw_text: str) -> List[str]:
        """Extract URL tokens from user input text."""
        tokens = re.split(r"[\s,;]+", raw_text.strip())
        cleaned = [clean_youtube_url(token.strip()) for token in tokens if token.strip().startswith(("http://", "https://"))]
        return [url for url in cleaned if url]

    def _download_youtube_items(self, urls: Sequence[str], output_dir: Path) -> List[Dict[str, str]]:
        """
        Download URLs with yt-dlp and return metadata used for Google Sheets insertion.

        Returns list of dicts with keys:
        - title
        - youtube_link
        """
        try:
            yt_dlp_module = importlib.import_module("yt_dlp")
        except Exception as exc:
            raise RuntimeError("yt-dlp is not installed. Run: pip install yt-dlp") from exc

        YoutubeDL: Any = getattr(yt_dlp_module, "YoutubeDL")

        output_dir.mkdir(parents=True, exist_ok=True)

        # yt-dlp calls this from the worker thread, so hand UI updates to the Tk main thread.
        def progress_hook(data: Dict[str, Any]) -> None:
            status = data.get("status")
            if status == "downloading":
                total = data.get("total_bytes") or data.get("total_bytes_estimate")
                downloaded = data.get("downloaded_bytes") or 0
                fraction = downloaded / total if isinstance(total, (int, float)) and total > 0 else 0.0
                self.after(0, self.set_download_progress, fraction)
            elif status == "finished":
                self.after(0, self.set_download_progress, 1.0)

        ydl_opts: Dict[str, Any] = {
            "outtmpl": str(output_dir / "%(title)s.%(ext)s"),
            "quiet": True,
            "no_warnings": True,
            "noplaylist": False,
            "progress_hooks": [progress_hook],
        }

        results: List[Dict[str, str]] = []

        with YoutubeDL(ydl_opts) as ydl:
            for url in urls:
                info = ydl.extract_info(url, download=True)
                if not isinstance(info, dict):
                    continue

                entries = info.get("entries")
                if isinstance(entries, list):
                    candidates = [entry for entry in entries if isinstance(entry, dict)]
                else:
                    candidates = [info]

                for entry in candidates:
                    title = str(entry.get("title") or "").strip()
                    youtube_link = clean_youtube_url(
                        str(
                        entry.get("webpage_url")
                        or entry.get("original_url")
                        or url
                        ).strip()
                    )

                    results.append(
                        {
                            "title": title,
                            "youtube_link": youtube_link,
                        }
                    )

        return results

    def on_download_and_add_clicked(self) -> None:
        """
        Download YouTube videos and insert only NEW records into worksheet "All Videos".
        Duplicates are skipped automatically.
        """
        raw_urls = self.urls_textbox.get("1.0", "end")
        urls = self._parse_urls_from_text(raw_urls)
        download_folder_text = self.download_folder_entry.get().strip()

        if not urls:
            self.set_download_status("Please enter at least one valid YouTube URL.", color=COLOR_WARN)
            return

        if not download_folder_text:
            self.set_download_status("Please select a download folder.", color=COLOR_WARN)
            return

        output_dir = Path(download_folder_text)

        self.download_button.configure(state="disabled")
        self.set_download_status(
            f"Downloading {len(urls)} link(s) and syncing to database... Please wait.",
            color=COLOR_INFO,
        )
        self.set_download_progress(0.0)

        def work() -> Tuple[int, int, int]:
            downloaded_items = self._download_youtube_items(urls, output_dir)
            if not downloaded_items:
                return 0, 0, 0

            sheets = SheetsClient(
                credentials_path=CREDENTIALS_PATH,
                spreadsheet_name=GOOGLE_SHEET_NAME,
            )
            sheets.connect()

            inserted_count, skipped_count = sheets.append_new_videos_to_all_videos(downloaded_items)
            return len(downloaded_items), inserted_count, skipped_count

        def on_done(result: Optional[Tuple[int, int, int]], error: Optional[BaseException]) -> None:
            self.download_button.configure(state="normal")

            if error is None and result is not None:
                downloaded_count, inserted_count, skipped_count = result
                if downloaded_count == 0:
                    self.set_download_status("Download finished but no video metadata was produced.", color=COLOR_WARN)
                else:
                    self.set_download_status(
                        f"Done: downloaded {downloaded_count} item(s), inserted {inserted_count} new row(s), "
                        f"skipped {skipped_count} duplicate(s).",
                        color=COLOR_OK,
                    )
            elif isinstance(error, FileNotFoundError):
                self.set_download_status(f"File not found: {error}", color=COLOR_ERROR)
            elif isinstance(error, ValueError):
                self.set_download_status(f"Configuration error: {error}", color=COLOR_ERROR)
            elif isinstance(error, gspread.exceptions.SpreadsheetNotFound):
                self.set_download_status(
                    "Google Sheet not found. Check GOOGLE_SHEET_NAME and sharing permissions.",
                    color=COLOR_ERROR,
                )
            elif isinstance(error, gspread.exceptions.WorksheetNotFound):
                self.set_download_status(
                    f"Worksheet missing: {error}. Required worksheet: 'All Videos'.",
                    color=COLOR_ERROR,
                )
            else:
                self.set_download_status(f"Unexpected error: {error}", color=COLOR_ERROR)

        self._run_in_background(work, on_done)

    def _fetch_title_from_url(self, url: str) -> str:
        """
        Fetch a YouTube title without downloading the video.
        """
        try:
            yt_dlp_module = importlib.import_module("yt_dlp")
        except Exception as exc:
            raise RuntimeError("yt-dlp is not installed. Run: pip install yt-dlp") from exc

        YoutubeDL: Any = getattr(yt_dlp_module, "YoutubeDL")

        ydl_opts: Dict[str, Any] = {
            "quiet": True,
            "no_warnings": True,
            "noplaylist": True,
            "skip_download": True,
        }

        with YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
            if isinstance(info, dict):
                return str(info.get("title") or "").strip()
        return ""

    def on_fetch_title_clicked(self) -> None:
        """
        Attempt to auto-fill the Title field based on the YouTube URL.
        """
        raw_url = self.manual_url_entry.get().strip()
        if not raw_url:
            return

        cleaned_url = clean_youtube_url(raw_url)
        if not cleaned_url:
            self.set_manual_status("Please enter a valid YouTube URL.", color=COLOR_WARN)
            return

        if self.manual_title_entry.get().strip():
            return

        self.set_manual_status("Fetching title...", color=COLOR_INFO)

        def on_done(title: Optional[str], error: Optional[BaseException]) -> None:
            if error is not None:
                self.set_manual_status(f"Failed to fetch title: {error}", color=COLOR_ERROR)
            elif not title:
                self.set_manual_status("Could not fetch title. You can enter it manually.", color=COLOR_WARN)
            elif self.manual_title_entry.get().strip():
                # User typed a title while we were fetching; keep theirs.
                self.set_manual_status("Title already entered; fetched title not applied.", color=COLOR_INFO)
            else:
                self.manual_title_entry.insert(0, title)
                self.set_manual_status("Title fetched successfully.", color=COLOR_OK)

        self._run_in_background(lambda: self._fetch_title_from_url(cleaned_url), on_done)

    def on_submit_manual_video(self) -> None:
        """
        Collect form values and append a new row to "All Videos".
        """
        raw_url = self.manual_url_entry.get().strip()
        title = self.manual_title_entry.get().strip()

        if not raw_url:
            self.set_manual_status("Please enter a YouTube URL.", color=COLOR_WARN)
            return

        if not title:
            self.set_manual_status("Please enter a Title.", color=COLOR_WARN)
            return

        cleaned_url = clean_youtube_url(raw_url)
        if not cleaned_url:
            self.set_manual_status("Please enter a valid YouTube URL.", color=COLOR_WARN)
            return

        original_location = self.manual_location_entry.get().strip()
        comments = self.manual_comments_entry.get().strip()
        area_similarity = self.manual_area_entry.get().strip()
        who_found_it = self.manual_found_entry.get().strip()

        row_values = [
            title,
            cleaned_url,
            original_location,
            *(combo.get().strip() for combo in self.rating_combos),
            comments,
        ]

        self.manual_submit_button.configure(state="disabled")
        self.set_manual_status("Submitting to database...", color=COLOR_INFO)
        self.update_idletasks()

        try:
            sheets = SheetsClient(
                credentials_path=CREDENTIALS_PATH,
                spreadsheet_name=GOOGLE_SHEET_NAME,
            )
            sheets.connect()
            existing_row = sheets.find_video_row(title, cleaned_url)

            if existing_row is not None:
                overwrite = messagebox.askyesno(
                    "Duplicate Found",
                    "This video already exists in the database.\n\nDo you want to overwrite its data and ratings?",
                )
                if not overwrite:
                    self.set_manual_status("Manual addition cancelled.", color=COLOR_WARN)
                    return

                sheets.update_manual_video_row(existing_row, row_values, area_similarity, who_found_it)
                self.set_manual_status("Video data updated successfully.", color=COLOR_OK)
            else:
                sheets.append_manual_video_row(row_values, area_similarity, who_found_it)
                self.set_manual_status("Manual video added successfully.", color=COLOR_OK)

        except FileNotFoundError as exc:
            self.set_manual_status(f"File not found: {exc}", color=COLOR_ERROR)
        except ValueError as exc:
            self.set_manual_status(f"Configuration error: {exc}", color=COLOR_ERROR)
        except gspread.exceptions.SpreadsheetNotFound:
            self.set_manual_status(
                "Google Sheet not found. Check GOOGLE_SHEET_NAME and sharing permissions.",
                color=COLOR_ERROR,
            )
        except gspread.exceptions.WorksheetNotFound as exc:
            self.set_manual_status(
                f"Worksheet missing: {exc}. Required worksheet: 'All Videos'.",
                color=COLOR_ERROR,
            )
        except Exception as exc:
            self.set_manual_status(f"Unexpected error: {exc}", color=COLOR_ERROR)
        finally:
            self.manual_submit_button.configure(state="normal")

    def _load_projects_into_dropdown(self) -> None:
        """
        Load CapCut projects from root_meta_info.json and populate combobox.
        """
        try:
            self.projects = load_capcut_projects(CAPCUT_ROOT_META_PATH)
        except FileNotFoundError:
            self.projects = []
            self.project_combobox.configure(values=["No CapCut projects found"])
            self.project_combobox.set("No CapCut projects found")
            self.set_status(
                f"CapCut root_meta_info.json not found at: {CAPCUT_ROOT_META_PATH}",
                color=COLOR_ERROR,
            )
            return
        except Exception as exc:
            self.projects = []
            self.project_combobox.configure(values=["Failed to load projects"])
            self.project_combobox.set("Failed to load projects")
            self.set_status(f"Error loading CapCut projects: {exc}", color=COLOR_ERROR)
            return

        if not self.projects:
            self.project_combobox.configure(values=["No projects available"])
            self.project_combobox.set("No projects available")
            self.set_status("No CapCut projects were found in root_meta_info.json.", color=COLOR_WARN)
            return

        # Build lookup and set dropdown values.
        names = [project.name for project in self.projects]
        self.project_lookup = {project.name: project for project in self.projects}

        self.project_combobox.configure(values=names)
        self.project_combobox.set(names[0])
        self.set_status(f"Loaded {len(names)} CapCut project(s).", color=COLOR_OK)

    def on_update_database_clicked(self) -> None:
        """
        Main button handler:
        1) Validate user input.
        2) Parse selected CapCut project videos.
        3) Match/update "All Videos".
        4) Append summary row to "Video Projects".
        """
        selected_project_name = self.project_combobox.get().strip()
        gdrive_link = self.gdrive_link_entry.get().strip()
        final_video_name = self.final_name_entry.get().strip()

        # Input validation.
        if not selected_project_name or selected_project_name in {
            "Loading projects...",
            "No CapCut projects found",
            "Failed to load projects",
            "No projects available",
        }:
            self.set_status("Please select a valid CapCut project.", color=COLOR_WARN)
            return

        if gdrive_link and not final_video_name:
            self.set_status("Error: You must click 'Validate & Fetch Title' first.", color=COLOR_WARN)
            return

        if not final_video_name:
            self.set_status("Please enter Final Video Name.", color=COLOR_WARN)
            return

        project = self.project_lookup.get(selected_project_name)
        if project is None:
            self.set_status("Selected project is not available in memory.", color=COLOR_ERROR)
            return

        # Prevent double-click while processing.
        self.update_button.configure(state="disabled")
        self.set_status("Processing selected project...", color=COLOR_INFO)
        self.update_idletasks()

        try:
            # Parse selected project's draft_content.json.
            draft_content_path = get_draft_content_path(project.folder_path)
            cleaned_video_names = extract_clean_video_names_from_draft(draft_content_path)

            if not cleaned_video_names:
                self.set_status(
                    "No valid video paths found in selected project (materials/videos/path).",
                    color=COLOR_WARN,
                )
                return

            # Connect to Google Sheets.
            sheets = SheetsClient(
                credentials_path=CREDENTIALS_PATH,
                spreadsheet_name=GOOGLE_SHEET_NAME,
            )
            sheets.connect()

            # Match against worksheet 1 and update "Used In".
            youtube_links = sheets.match_youtube_links(
                cleaned_video_names=cleaned_video_names,
                score_threshold=MATCH_THRESHOLD,
            )

            if not youtube_links:
                self.set_status(
                    "Error: No matching videos found in the database. Did you use the Ingest tool to download them?",
                    color=COLOR_ERROR,
                )
                return

            # Append row in worksheet 2 regardless of link count (POC behavior).
            existing_row = sheets.find_project_row(final_video_name)
            if existing_row is not None:
                overwrite = messagebox.askyesno(
                    "Duplicate Found",
                    f"The project '{final_video_name}' already exists in the database.\n\nDo you want to overwrite it?",
                )
                if not overwrite:
                    self.set_status("Submission cancelled by user.", color=COLOR_WARN)
                    return

                sheets.update_video_project_log(
                    row_index=existing_row,
                    final_video_name=final_video_name,
                    youtube_links=youtube_links,
                    gdrive_link=gdrive_link,
                )
                self.set_status("Project updated successfully.", color=COLOR_OK)
            else:
                sheets.append_video_project_log(
                    final_video_name=final_video_name,
                    youtube_links=youtube_links,
                    gdrive_link=gdrive_link,
                )

                self.set_status(
                    f"Success: processed {len(cleaned_video_names)} video file(s), "
                    f"matched {len(youtube_links)} link(s), and updated the database.",
                    color=COLOR_OK,
                )

        except FileNotFoundError as exc:
            self.set_status(f"File not found: {exc}", color=COLOR_ERROR)
        except ValueError as exc:
            self.set_status(f"Configuration/JSON error: {exc}", color=COLOR_ERROR)
        except gspread.exceptions.SpreadsheetNotFound:
            self.set_status(
                "Google Sheet not found. Check GOOGLE_SHEET_NAME and sharing permissions.",
                color=COLOR_ERROR,
            )
        except gspread.exceptions.WorksheetNotFound as exc:
            self.set_status(
                f"Worksheet missing: {exc}. Required: 'All Videos' and 'Video Projects'.",
                color=COLOR_ERROR,
            )
        except Exception as exc:
            # Catch-all for unexpected runtime issues.
            self.set_status(f"Unexpected error: {exc}", color=COLOR_ERROR)
        finally:
            self.update_button.configure(state="normal")


# ===========
# Entry point
# ===========

if __name__ == "__main__":
    app = CapCutTrackerApp()
    app.mainloop()
