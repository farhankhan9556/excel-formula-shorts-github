from __future__ import annotations

import json
import random
import shutil
from datetime import datetime, timedelta, timezone
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"
UPLOADED = OUT / "uploaded"
SECRETS = Path(r"C:\LearnVerseSecrets")
CLIENT_SECRET = SECRETS / "client_secret.json"
TOKEN = SECRETS / "token.json"
STATE = SECRETS / "youtube_upload_state.json"
LOG = SECRETS / "uploader.log"
SCOPES = ["https://www.googleapis.com/auth/youtube.upload"]
UAE = timezone(timedelta(hours=4), name="UAE")

WINDOWS = {
    "morning": (9, 10, 1),
    "afternoon": (13, 14, 2),
    "evening": (20, 21, 3),
}


def log(message):
    message = str(message)
    print(message)
    SECRETS.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.now(UAE):%Y-%m-%d %H:%M:%S} UAE | {message}\n")


def load_state():
    if not STATE.exists():
        return {"uploaded_files": [], "uploaded_video_ids": {}, "daily_schedule": {}}
    try:
        d = json.loads(STATE.read_text(encoding="utf-8"))
        if not isinstance(d, dict):
            raise ValueError
    except Exception:
        d = {}
    d.setdefault("uploaded_files", [])
    d.setdefault("uploaded_video_ids", {})
    d.setdefault("daily_schedule", {})
    return d


def save_state(state):
    STATE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")


def daily_schedule(state, date_key):
    schedule = state["daily_schedule"].get(date_key)
    if schedule:
        return schedule
    schedule = {
        name: random.randint(start_minute, 54)
        for name, (_, _, start_minute) in WINDOWS.items()
    }
    # Store actual minute targets, not exact upload timestamps.
    schedule = {name: random.randint(0, 54) for name in WINDOWS}
    state["daily_schedule"][date_key] = schedule
    # Keep state small.
    old = sorted(state["daily_schedule"].keys())[:-14]
    for key in old:
        state["daily_schedule"].pop(key, None)
    save_state(state)
    return schedule


def authenticate():
    SECRETS.mkdir(parents=True, exist_ok=True)
    creds = None
    if TOKEN.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN), SCOPES)
        except Exception:
            creds = None

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        TOKEN.write_text(creds.to_json(), encoding="utf-8")

    if not creds or not creds.valid:
        if not CLIENT_SECRET.exists():
            raise FileNotFoundError(f"Missing Google OAuth file: {CLIENT_SECRET}")
        flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRET), SCOPES)
        creds = flow.run_local_server(port=0)
        TOKEN.write_text(creds.to_json(), encoding="utf-8")

    return creds


def get_youtube():
    return build("youtube", "v3", credentials=authenticate())


def metadata_files():
    return sorted(
        p for p in OUT.glob("*.json")
        if p.name != "manifest.json"
    )


def upload_video(youtube, metadata):
    video_path = OUT / metadata["video"]
    body = {
        "snippet": {
            "title": metadata.get("title", "Excel Tip | Learn Verse")[:100],
            "description": metadata.get("description", ""),
            "tags": [x.lstrip("#") for x in metadata.get("hashtags", [])],
            "categoryId": "27",
        },
        "status": {
            "privacyStatus": metadata.get("privacy", "public"),
            "selfDeclaredMadeForKids": False,
        },
    }
    media = MediaFileUpload(str(video_path), mimetype="video/mp4", resumable=True, chunksize=8 * 1024 * 1024)
    request = youtube.videos().insert(part="snippet,status", body=body, media_body=media)
    response = None
    while response is None:
        _, response = request.next_chunk()
    return response["id"]


def move_uploaded(metadata_path, metadata):
    date_folder = UPLOADED / metadata["date"]
    date_folder.mkdir(parents=True, exist_ok=True)
    for name in (metadata["video"], metadata["workbook"], metadata_path.name):
        src = OUT / name
        if src.exists():
            shutil.move(str(src), str(date_folder / name))


def main():
    now = datetime.now(UAE)
    state = load_state()
    date_key = now.strftime("%Y-%m-%d")
    schedule = daily_schedule(state, date_key)

    log("==========================================")
    log("LEARN VERSE YOUTUBE AUTO UPLOADER")
    log(f"Current UAE time: {now:%Y-%m-%d %H:%M:%S}")
    log(f"Today's target minutes: morning {schedule['morning']}, afternoon {schedule['afternoon']}, evening {schedule['evening']}")

    # Pick the earliest due slot that has not already been uploaded.
    due = []
    for name, (start_hour, end_hour, slot_number) in WINDOWS.items():
        target = schedule[name]
        target_time = now.replace(hour=start_hour, minute=target, second=0, microsecond=0)
        if now >= target_time:
            due.append((target_time, name, slot_number))

    if not due:
        log("No upload slot is due yet.")
        return

    due.sort(key=lambda x: x[0])
    _, window_name, slot_number = due[0]
    completed_key = f"{date_key}:{window_name}"

    if completed_key in state.get("uploaded_files", []):
        log(f"{window_name} upload already completed.")
        return

    candidates = []
    for p in metadata_files():
        try:
            m = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        if m.get("date") == date_key and int(m.get("slot_number", 0)) == slot_number:
            if m.get("video") and (OUT / m["video"]).exists():
                candidates.append((p, m))

    if not candidates:
        log(f"No generated video is ready for {window_name}.")
        return

    metadata_path, metadata = candidates[0]
    youtube = get_youtube()
    log(f"Uploading: {metadata['video']}")

    try:
        video_id = upload_video(youtube, metadata)
    except HttpError as exc:
        text = str(exc)
        if "quotaExceeded" in text or "uploadLimitExceeded" in text:
            log("YouTube daily quota/upload limit reached. Stopping without retry loop.")
            return
        raise

    log(f"UPLOAD SUCCESSFUL: https://www.youtube.com/watch?v={video_id}")
    move_uploaded(metadata_path, metadata)
    state.setdefault("uploaded_video_ids", {})[metadata["video"]] = video_id
    state.setdefault("uploaded_files", []).append(completed_key)
    save_state(state)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log(f"ERROR: {exc}")
        raise
