import os
import sys
import json
import random
import shutil
import logging
from pathlib import Path
from datetime import datetime, timedelta, timezone

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError


# ============================================================
# LEARN VERSE YOUTUBE AUTO UPLOADER
# ============================================================

BASE_DIR = Path(r"C:\excel-formula-shorts-github")
OUTPUT_DIR = BASE_DIR / "output"

SECRETS_DIR = Path(r"C:\LearnVerseSecrets")
CLIENT_SECRET_FILE = SECRETS_DIR / "client_secret.json"
TOKEN_FILE = SECRETS_DIR / "token.json"

STATE_FILE = SECRETS_DIR / "youtube_upload_state.json"
LOG_FILE = SECRETS_DIR / "youtube_uploader.log"

UPLOADED_DIR = OUTPUT_DIR / "uploaded"

SCOPES = [
    "https://www.googleapis.com/auth/youtube.upload"
]

# UAE = UTC+4
UAE = timezone(
    timedelta(hours=4),
    name="UAE"
)


# ============================================================
# UPLOAD WINDOWS
# ============================================================

WINDOWS = {
    "morning": {
        "start_hour": 9,
        "end_hour": 10,
        "video_number": 1,
    },
    "afternoon": {
        "start_hour": 13,
        "end_hour": 14,
        "video_number": 2,
    },
    "evening": {
        "start_hour": 20,
        "end_hour": 21,
        "video_number": 3,
    },
}


# ============================================================
# LOGGING
# ============================================================

SECRETS_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
UPLOADED_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)


def log(message):
    print(message)
    logging.info(message)


# ============================================================
# STATE
# ============================================================

def default_state():
    return {
        "days": {},
        "uploaded_files": [],
        "uploads": [],
    }


def load_state():
    if not STATE_FILE.exists():
        return default_state()

    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        if not isinstance(data, dict):
            return default_state()

        data.setdefault("days", {})
        data.setdefault("uploaded_files", [])
        data.setdefault("uploads", [])

        return data

    except Exception as e:
        log(f"WARNING: Could not read state file: {e}")
        return default_state()


def save_state(state):
    SECRETS_DIR.mkdir(parents=True, exist_ok=True)

    temp_file = STATE_FILE.with_suffix(".tmp")

    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(
            state,
            f,
            indent=2,
            ensure_ascii=False,
        )

    os.replace(temp_file, STATE_FILE)


# ============================================================
# DAILY SCHEDULE
# ============================================================

def create_daily_schedule(date_obj, state):
    date_key = date_obj.strftime("%Y-%m-%d")

    if date_key in state["days"]:
        return state["days"][date_key]

    schedule = {}

    for window_name, config in WINDOWS.items():
        minute = random.randint(0, 54)

        schedule[window_name] = (
            f"{config['start_hour']:02d}:{minute:02d}"
        )

    state["days"][date_key] = {
        "schedule": schedule,
        "completed_windows": {},
    }

    save_state(state)

    return state["days"][date_key]


def parse_schedule_time(date_key, time_string):
    hour, minute = map(int, time_string.split(":"))

    year, month, day = map(
        int,
        date_key.split("-")
    )

    return datetime(
        year,
        month,
        day,
        hour,
        minute,
        tzinfo=UAE,
    )


# ============================================================
# METADATA / FILE MAPPING
# ============================================================

def load_metadata(metadata_file):
    with open(
        metadata_file,
        "r",
        encoding="utf-8"
    ) as f:
        return json.load(f)


def find_video_for_slot(date_key, video_number):
    pattern = (
        f"learnverse_{date_key}_"
        f"{video_number:02d}_*.json"
    )

    metadata_files = sorted(
        OUTPUT_DIR.glob(pattern)
    )

    if not metadata_files:
        return None, None

    for metadata_file in metadata_files:

        try:
            metadata = load_metadata(metadata_file)
        except Exception as e:
            log(
                f"WARNING: Cannot read "
                f"{metadata_file.name}: {e}"
            )
            continue

        if str(metadata.get("date", "")) != date_key:
            continue

        if int(metadata.get("video_number", 0)) != video_number:
            continue

        video_file_name = metadata.get("video_file")

        if not video_file_name:
            continue

        video_path = OUTPUT_DIR / video_file_name

        if not video_path.exists():
            log(
                f"Video file missing: "
                f"{video_path.name}"
            )
            continue

        return metadata_file, metadata

    return None, None


def get_video_path(metadata):
    video_file = metadata.get("video_file")

    if not video_file:
        raise FileNotFoundError(
            "Metadata does not contain video_file."
        )

    video_path = OUTPUT_DIR / video_file

    if not video_path.exists():
        raise FileNotFoundError(
            f"Video not found: {video_path}"
        )

    return video_path


def get_workbook_path(metadata):
    workbook_file = metadata.get("workbook_file")

    if not workbook_file:
        return None

    workbook_path = OUTPUT_DIR / workbook_file

    if workbook_path.exists():
        return workbook_path

    return None


def get_youtube_title(metadata):
    title = metadata.get("youtube_title")

    if title:
        return str(title)

    formula_name = metadata.get(
        "formula_name",
        "Excel Formula"
    )

    return (
        f"{formula_name} Excel Formula "
        f"That Saves Time! #Excel #Shorts"
    )


def get_youtube_description(metadata):
    description = metadata.get(
        "youtube_description"
    )

    if description:
        return str(description)

    formula = metadata.get(
        "formula",
        ""
    )

    explanation = metadata.get(
        "explanation",
        ""
    )

    return (
        f"Learn this practical Excel formula.\n\n"
        f"Formula: {formula}\n\n"
        f"{explanation}\n\n"
        f"Follow Learn Verse for more practical "
        f"Excel tips.\n"
        f"@LearnVerse9556"
    )


def get_youtube_tags(metadata):
    hashtags = metadata.get(
        "hashtags",
        [
            "#Excel",
            "#ExcelTips",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnVerse",
            "#Shorts",
        ],
    )

    tags = []

    for item in hashtags:
        tag = str(item).strip()

        if tag.startswith("#"):
            tag = tag[1:]

        if tag:
            tags.append(tag)

    return tags


# ============================================================
# SLOT INFORMATION
# ============================================================

def print_slot_information(date_key, video_number):

    metadata_file, metadata = find_video_for_slot(
        date_key,
        video_number
    )

    print(f"Video {video_number}:")

    if metadata_file is None:
        print("  Metadata : NOT FOUND")
        print("  Video    : NOT FOUND")
        print("  Workbook : NOT FOUND")
        print("  Title    : NOT FOUND")
        return

    video_file = metadata.get(
        "video_file",
        "NOT FOUND"
    )

    workbook_file = metadata.get(
        "workbook_file",
        "NOT FOUND"
    )

    title = get_youtube_title(metadata)

    print(
        f"  Metadata : {metadata_file.name}"
    )

    print(
        f"  Video    : {video_file}"
    )

    print(
        f"  Workbook : {workbook_file}"
    )

    print(
        f"  Title    : {title}"
    )


# ============================================================
# YOUTUBE AUTHENTICATION
# ============================================================

def authenticate():
    credentials = None

    if TOKEN_FILE.exists():

        try:
            credentials = Credentials.from_authorized_user_file(
                str(TOKEN_FILE),
                SCOPES
            )
        except Exception as e:
            log(
                f"WARNING: Could not load token: {e}"
            )
            credentials = None

    if credentials and credentials.expired:

        if credentials.refresh_token:

            log(
                "Refreshing YouTube authorization..."
            )

            credentials.refresh(Request())

        else:
            credentials = None

    if not credentials or not credentials.valid:

        if not CLIENT_SECRET_FILE.exists():
            raise FileNotFoundError(
                f"Client secret not found: "
                f"{CLIENT_SECRET_FILE}"
            )

        log(
            "Starting Google OAuth authorization..."
        )

        flow = InstalledAppFlow.from_client_secrets_file(
            str(CLIENT_SECRET_FILE),
            SCOPES
        )

        credentials = flow.run_local_server(
            port=0,
            access_type="offline",
            prompt="consent",
        )

        with open(
            TOKEN_FILE,
            "w",
            encoding="utf-8"
        ) as f:
            f.write(credentials.to_json())

    granted_scopes = (
        credentials.scopes or []
    )

    log(
        f"Granted scopes: {granted_scopes}"
    )

    required_scope = (
        "https://www.googleapis.com/auth/"
        "youtube.upload"
    )

    if required_scope not in granted_scopes:
        raise RuntimeError(
            "YouTube upload permission is not "
            "available in the current OAuth token."
        )

    log(
        "Required youtube.upload scope confirmed."
    )

    return credentials


def create_youtube_service():
    credentials = authenticate()

    return build(
        "youtube",
        "v3",
        credentials=credentials,
        cache_discovery=False,
    )


# ============================================================
# YOUTUBE UPLOAD
# ============================================================

def upload_video(metadata):

    video_path = get_video_path(metadata)

    title = get_youtube_title(metadata)
    description = get_youtube_description(metadata)
    tags = get_youtube_tags(metadata)

    log("")
    log("==========================================")
    log("UPLOADING VIDEO")
    log("==========================================")

    log(
        f"File: {video_path.name}"
    )

    log(
        f"Title: {title}"
    )

    log(
        f"Size: {video_path.stat().st_size:,} bytes"
    )

    youtube = create_youtube_service()

    body = {
        "snippet": {
            "title": title,
            "description": description,
            "tags": tags,
            "categoryId": "27",
        },
        "status": {
            "privacyStatus": "public",
            "selfDeclaredMadeForKids": False,
        },
    }

    media = MediaFileUpload(
        str(video_path),
        mimetype="video/mp4",
        resumable=True,
        chunksize=1024 * 1024,
    )

    request = youtube.videos().insert(
        part="snippet,status",
        body=body,
        media_body=media,
    )

    response = None

    while response is None:

        try:

            status, response = request.next_chunk()

            if status:
                progress = int(
                    status.progress() * 100
                )

                log(
                    f"Upload progress: "
                    f"{progress}%"
                )

        except HttpError as e:

            log(
                f"YouTube HTTP error: {e}"
            )

            raise

    video_id = response["id"]

    video_url = (
        f"https://www.youtube.com/watch?v="
        f"{video_id}"
    )

    log("")
    log("UPLOAD SUCCESSFUL")
    log(
        f"Video ID: {video_id}"
    )
    log(
        f"Video URL: {video_url}"
    )

    return {
        "video_id": video_id,
        "video_url": video_url,
    }


# ============================================================
# UPLOAD STATE
# ============================================================

def is_window_completed(
    state,
    date_key,
    window_name
):
    day_state = state["days"].get(
        date_key,
        {}
    )

    completed = day_state.get(
        "completed_windows",
        {}
    )

    return bool(
        completed.get(window_name)
    )


def record_upload(
    state,
    date_key,
    window_name,
    video_number,
    metadata,
    upload_result
):

    state["days"].setdefault(
        date_key,
        {}
    )

    state["days"][date_key].setdefault(
        "completed_windows",
        {}
    )

    state["days"][date_key][
        "completed_windows"
    ][window_name] = True

    video_file = metadata.get(
        "video_file"
    )

    if video_file:
        if video_file not in state[
            "uploaded_files"
        ]:
            state[
                "uploaded_files"
            ].append(video_file)

    state["uploads"].append({
        "date": date_key,
        "window": window_name,
        "video_number": video_number,
        "video_file": video_file,
        "video_id": upload_result[
            "video_id"
        ],
        "video_url": upload_result[
            "video_url"
        ],
        "uploaded_at": datetime.now(
            UAE
        ).isoformat(),
    })

    save_state(state)

    log(
        f"Upload state saved: "
        f"{video_file}"
    )


# ============================================================
# MOVE SUCCESSFULLY UPLOADED FILES
# ============================================================

def move_uploaded_files(
    metadata_file,
    metadata,
    date_key
):

    destination = (
        UPLOADED_DIR / date_key
    )

    destination.mkdir(
        parents=True,
        exist_ok=True
    )

    files_to_move = [
        metadata_file
    ]

    video_file = metadata.get(
        "video_file"
    )

    workbook_file = metadata.get(
        "workbook_file"
    )

    if video_file:
        files_to_move.append(
            OUTPUT_DIR / video_file
        )

    if workbook_file:
        files_to_move.append(
            OUTPUT_DIR / workbook_file
        )

    for source in files_to_move:

        if not source.exists():
            continue

        target = destination / source.name

        if target.exists():
            target.unlink()

        shutil.move(
            str(source),
            str(target)
        )

        log(
            f"Moved uploaded file: "
            f"{source.name}"
        )


# ============================================================
# AUTHENTICATION TEST
# ============================================================

def authentication_test():

    print("==========================================")
    print("YOUTUBE AUTHENTICATION TEST")
    print("==========================================")

    try:
        credentials = authenticate()

        if credentials and credentials.valid:
            print(
                "=========================================="
            )
            print(
                "AUTHENTICATION TEST SUCCESSFUL"
            )
            print(
                "NO VIDEO WAS UPLOADED"
            )
            print(
                "=========================================="
            )
            return True

    except Exception as e:

        print(
            "AUTHENTICATION TEST FAILED"
        )

        print(
            f"Error: {e}"
        )

        return False

    return False


# ============================================================
# SAFE DRY RUN
# ============================================================

def dry_run():

    print("==========================================")
    print("LEARN VERSE YOUTUBE UPLOADER DRY RUN")
    print("==========================================")

    now = datetime.now(UAE)

    print(
        f"Current UAE time: "
        f"{now.strftime('%Y-%m-%d %H:%M:%S')}"
    )

    date_key = now.strftime(
        "%Y-%m-%d"
    )

    state = load_state()

    daily = create_daily_schedule(
        now.date(),
        state
    )

    print("")
    print("Today's target minutes:")

    for window_name, config in WINDOWS.items():

        target = daily[
            "schedule"
        ][window_name]

        print(
            f"{window_name:<9} {target}"
        )

    print("")

    for window_name, config in WINDOWS.items():

        print(
            f"Checking {window_name}..."
        )

        print_slot_information(
            date_key,
            config["video_number"]
        )

    print("")
    print(
        "DRY RUN COMPLETE"
    )
    print(
        "NO VIDEO WAS UPLOADED"
    )
    print(
        "=========================================="
    )


# ============================================================
# SCHEDULED UPLOAD
# ============================================================

def scheduled_upload():

    now = datetime.now(UAE)

    log("")
    log("==========================================")
    log("LEARN VERSE YOUTUBE AUTO UPLOADER")
    log("==========================================")

    log(
        f"Current UAE time: "
        f"{now.strftime('%Y-%m-%d %H:%M:%S')}"
    )

    date_key = now.strftime(
        "%Y-%m-%d"
    )

    state = load_state()

    daily = create_daily_schedule(
        now.date(),
        state
    )

    log("Today's target minutes:")

    for window_name in WINDOWS:

        target = daily[
            "schedule"
        ][window_name]

        log(
            f"{window_name:<9} {target}"
        )

    active_window = None

    for window_name, config in WINDOWS.items():

        if (
            config["start_hour"]
            <= now.hour
            < config["end_hour"]
        ):
            active_window = window_name
            break

    if active_window is None:

        log(
            "No upload window is active."
        )

        return

    config = WINDOWS[
        active_window
    ]

    if is_window_completed(
        state,
        date_key,
        active_window
    ):

        log(
            f"{active_window.capitalize()} "
            f"upload already completed."
        )

        return

    target_time = parse_schedule_time(
        date_key,
        daily["schedule"][
            active_window
        ]
    )

    if now < target_time:

        log(
            f"{active_window.capitalize()} "
            f"upload target has not been reached yet."
        )

        log(
            f"Target time: "
            f"{target_time.strftime('%H:%M')}"
        )

        return

    video_number = config[
        "video_number"
    ]

    log(
        f"Active window: "
        f"{active_window}"
    )

    log(
        f"Selected video number: "
        f"{video_number}"
    )

    metadata_file, metadata = (
        find_video_for_slot(
            date_key,
            video_number
        )
    )

    if metadata_file is None:

        log(
            f"Video {video_number} "
            f"for {date_key} was not found."
        )

        return

    log(
        f"Metadata selected: "
        f"{metadata_file.name}"
    )

    video_path = get_video_path(
        metadata
    )

    log(
        f"Video selected: "
        f"{video_path.name}"
    )

    # --------------------------------------------------------
    # SAFE DRY-RUN SWITCH
    # --------------------------------------------------------

    if os.environ.get(
        "YOUTUBE_DRY_RUN",
        ""
    ) == "1":

        log("")
        log(
            "YOUTUBE_DRY_RUN=1"
        )
        log(
            "Upload would happen now."
        )
        log(
            "NO VIDEO WAS UPLOADED."
        )

        return

    # --------------------------------------------------------
    # REAL UPLOAD
    # --------------------------------------------------------

    try:

        upload_result = upload_video(
            metadata
        )

    except HttpError as e:

        log(
            f"Upload failed: {e}"
        )

        if getattr(
            e,
            "content",
            None
        ):
            log(
                f"Server response: "
                f"{e.content}"
            )

        return

    except Exception as e:

        log(
            f"Upload failed: "
            f"{e}"
        )

        return

    record_upload(
        state=state,
        date_key=date_key,
        window_name=active_window,
        video_number=video_number,
        metadata=metadata,
        upload_result=upload_result,
    )

    move_uploaded_files(
        metadata_file,
        metadata,
        date_key
    )

    log("")
    log(
        "=========================================="
    )
    log(
        "SCHEDULED UPLOAD FINISHED"
    )
    log(
        "=========================================="
    )


# ============================================================
# MAIN
# ============================================================

def main():

    if os.environ.get(
        "YOUTUBE_AUTH_TEST",
        ""
    ) == "1":

        success = authentication_test()

        sys.exit(
            0 if success else 1
        )

    if os.environ.get(
        "YOUTUBE_DRY_RUN",
        ""
    ) == "1":

        dry_run()
        return

    scheduled_upload()


if __name__ == "__main__":
    main()