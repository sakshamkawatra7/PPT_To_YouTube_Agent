from __future__ import annotations
import argparse, json, logging, re, shutil, subprocess, sys, time, zipfile
from pathlib import Path
import os
from xml.etree import ElementTree as ET

import pymupdf
import edge_tts
import asyncio
from pptx import Presentation
from pptx.exc import PackageNotFoundError
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from concurrent.futures import ThreadPoolExecutor, as_completed
import uuid

import job_registry

BASE = Path(__file__).resolve().parent
CONFIG_PATH = BASE / "config.json"
SCOPES = ["https://www.googleapis.com/auth/youtube.upload"]
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

def load_config():
    cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    for key in ("input_dir", "work_dir", "output_dir"):
        cfg[key] = str((BASE / cfg[key]).resolve())
    return cfg

def setup_logging():
    log_dir = BASE / "logs"
    log_dir.mkdir(exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_dir / "agent.log", encoding="utf-8"), logging.StreamHandler(sys.stdout)],
        force=True,
    )

def clean_text(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()

def extract_notes_from_pptx(pptx_path: Path) -> dict[int, str]:
    """Return {slide_number: speaker_notes}. Uses PPTX relationships, not filename ordering."""
    notes = {}
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = set(z.namelist())
            for name in names:
                if not re.fullmatch(r"ppt/notesSlides/notesSlide\d+\.xml", name):
                    continue
                rel_name = name.replace("notesSlides/", "notesSlides/_rels/") + ".rels"
                slide_number = None
                if rel_name in names:
                    rel_root = ET.fromstring(z.read(rel_name))
                    for rel in rel_root.findall("rel:Relationship", NS):
                        if rel.attrib.get("Type", "").endswith("/slide"):
                            target = rel.attrib.get("Target", "")
                            m = re.search(r"slide(\d+)\.xml$", target)
                            if m:
                                slide_number = int(m.group(1))
                                break
                if slide_number is None:
                    m = re.search(r"notesSlide(\d+)\.xml$", name)
                    slide_number = int(m.group(1)) if m else None
                root = ET.fromstring(z.read(name))
                chunks = []
                for shape in root.findall(".//p:sp", NS):
                    ph = shape.find(".//p:ph", NS)
                    ph_type = ph.attrib.get("type", "") if ph is not None else ""
                    if ph_type in {"sldNum", "dt", "hdr", "ftr"}:
                        continue
                    vals = [t.text or "" for t in shape.findall(".//a:t", NS)]
                    txt = clean_text(" ".join(vals))
                    if txt:
                        chunks.append(txt)
                # Notes XML often contains slide-image placeholders; body text is the useful content.
                result = clean_text(" ".join(chunks))
                if slide_number and result:
                    notes[slide_number] = result
    except Exception as exc:
        logging.warning("Could not parse speaker notes (%s). Visible slide text will be used as fallback.", exc)
    return notes

def extract_slide_scripts(pptx_path: Path, cfg: dict) -> list[str]:
    prs = Presentation(str(pptx_path))
    notes = extract_notes_from_pptx(pptx_path)
    scripts = []
    for idx, slide in enumerate(prs.slides, start=1):
        script = notes.get(idx, "")
        if not script:
            bits = []
            for shape in slide.shapes:
                if getattr(shape, "has_text_frame", False):
                    t = clean_text(shape.text)
                    if t:
                        bits.append(t)
            script = clean_text(" ".join(bits))
        if not script:
            script = f"Slide {idx}."
        limit = int(cfg.get("max_script_chars_per_slide", 5000))
        if len(script) > limit:
            logging.warning("Slide %s script clipped to %s characters.", idx, limit)
            script = script[:limit]
        scripts.append(script)
    return scripts


def find_soffice() -> str:
    candidates = [
        shutil.which("soffice"),

        # macOS
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",

        # Windows
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",

        # Linux
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
    ]

    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(candidate)

    raise FileNotFoundError(
        "LibreOffice not found. Install LibreOffice and ensure "
        "its program folder is available in PATH."
    )

def find_ffmpeg() -> str:
    exe = shutil.which("ffmpeg")
    if not exe:
        raise FileNotFoundError("FFmpeg not found. Install FFmpeg and add its bin folder to PATH.")
    return exe


def render_slides(pptx_path: Path, work: Path) -> list[Path]:
    pdf_dir = work / "pdf"
    img_dir = work / "slides"
    profile_dir = work / "libreoffice_profile"

    pdf_dir.mkdir(parents=True, exist_ok=True)
    img_dir.mkdir(parents=True, exist_ok=True)
    profile_dir.mkdir(parents=True, exist_ok=True)

    soffice = find_soffice()

    # Give every presentation its own LibreOffice user profile.
    profile_uri = profile_dir.resolve().as_uri()

    cmd = [
        soffice,
        "--headless",
        f"-env:UserInstallation={profile_uri}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(pdf_dir),
        str(pptx_path),
    ]

    logging.info("Rendering PowerPoint to PDF: %s", pptx_path.name)

    subprocess.run(
        cmd,
        check=True,
        capture_output=True,
        text=True,
    )

    pdf_path = pdf_dir / (pptx_path.stem + ".pdf")

    if not pdf_path.exists():
        raise RuntimeError(
            "LibreOffice did not create the expected PDF. "
            "Check whether the PPTX opens correctly."
        )

    doc = pymupdf.open(pdf_path)

    images = []

    try:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(
                matrix=pymupdf.Matrix(1.8, 1.8),
                alpha=False,
            )

            out = img_dir / f"slide_{i+1:03d}.jpg"
            pix.save(out, jpg_quality=90)
            images.append(out)

    finally:
        doc.close()

    return images


async def make_tts(text: str, out_file: Path, cfg: dict):
    voice = cfg.get("voice", "en-IN-NeerjaNeural")
    communicate = edge_tts.Communicate(
        text=text, voice=voice,
        rate=cfg.get("rate", "+0%"),
        volume=cfg.get("volume", "+0%"),
    )
    await communicate.save(str(out_file))


def probe_duration(ffmpeg: str, audio_path: Path) -> float:
    ffmpeg_path = Path(ffmpeg)

    # Look for ffprobe alongside ffmpeg first.
    ffprobe_name = "ffprobe.exe" if os.name == "nt" else "ffprobe"
    ffprobe_candidate = ffmpeg_path.with_name(ffprobe_name)

    if ffprobe_candidate.exists():
        ffprobe = str(ffprobe_candidate)
    else:
        # Fall back to searching the system PATH.
        ffprobe = shutil.which("ffprobe")

    if not ffprobe:
        raise FileNotFoundError(
            "FFprobe not found. Install FFmpeg and ensure FFprobe is available in PATH."
        )

    result = subprocess.run(
        [
            ffprobe,
            "-v", "error",
            "-show_entries", "format=duration",
            "-of", "default=noprint_wrappers=1:nokey=1",
            str(audio_path),
        ],
        check=True,
        capture_output=True,
        text=True,
    )

    return max(0.5, float(result.stdout.strip()))

def srt_time(seconds: float) -> str:
    ms = int(round(seconds * 1000))
    h, rem = divmod(ms, 3600000)
    m, rem = divmod(rem, 60000)
    s, ms = divmod(rem, 1000)
    return f"{h:02}:{m:02}:{s:02},{ms:03}"

def build_video(pptx_path: Path, scripts: list[str], slide_images: list[Path], work: Path, out_mp4: Path, cfg: dict):
    ffmpeg = find_ffmpeg()
    audio_dir = work / "audio"
    segment_dir = work / "segments"
    audio_dir.mkdir(parents=True, exist_ok=True)
    segment_dir.mkdir(parents=True, exist_ok=True)

    durations = []
    audio_files = []
    for i, script in enumerate(scripts, start=1):
        audio = audio_dir / f"slide_{i:03d}.mp3"
        logging.info("Generating Indian-English narration for slide %s/%s...", i, len(scripts))
        asyncio.run(make_tts(script, audio, cfg))
        audio_files.append(audio)
        durations.append(probe_duration(ffmpeg, audio))

    width, height, fps = int(cfg.get("video_width", 1920)), int(cfg.get("video_height", 1080)), int(cfg.get("fps", 25))
    segment_paths = []
    for i, (image, audio, duration) in enumerate(zip(slide_images, audio_files, durations), start=1):
        segment = segment_dir / f"segment_{i:03d}.mp4"
        vf = f"scale={width}:{height}:force_original_aspect_ratio=decrease,pad={width}:{height}:(ow-iw)/2:(oh-ih)/2,setsar=1"
        cmd = [
            ffmpeg, "-y", "-loop", "1", "-framerate", str(fps), "-i", str(image), "-i", str(audio),
            "-vf", vf, "-t", f"{duration:.3f}", "-r", str(fps),
            "-c:v", "libx264", "-preset", "veryfast", "-crf", "23",
            "-c:a", "aac", "-b:a", "160k", "-pix_fmt", "yuv420p",
            "-shortest", str(segment)
        ]
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        # raise RuntimeError("TEST: Simulated video generation failure")
        segment_paths.append(segment)

    concat_file = work / "concat.txt"
    concat_file.write_text("".join(f"file '{p.as_posix()}'\n" for p in segment_paths), encoding="utf-8")
    out_mp4.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [ffmpeg, "-y", "-f", "concat", "-safe", "0", "-i", str(concat_file), "-c", "copy", "-movflags", "+faststart", str(out_mp4)],
        check=True, capture_output=True, text=True
    )

    # Simple slide-level subtitle file, aligned to each slide's narration duration.
    elapsed, entries = 0.0, []
    for i, (script, duration) in enumerate(zip(scripts, durations), start=1):
        entries.append(f"{i}\n{srt_time(elapsed)} --> {srt_time(elapsed + duration)}\n{script}\n")
        elapsed += duration
    out_mp4.with_suffix(".srt").write_text("\n".join(entries), encoding="utf-8")

    # Use the first slide as a simple 16:9 thumbnail.
    thumb = out_mp4.with_name(out_mp4.stem + "_thumbnail.jpg")
    subprocess.run([ffmpeg, "-y", "-i", str(slide_images[0]), "-vf",
                    f"scale=1280:720:force_original_aspect_ratio=decrease,pad=1280:720:(ow-iw)/2:(oh-ih)/2",
                    "-frames:v", "1", str(thumb)], check=True, capture_output=True, text=True)
    return out_mp4

def youtube_service():
    creds_path = BASE / "credentials.json"
    token_path = BASE / "token.json"
    if not creds_path.exists():
        raise FileNotFoundError("credentials.json missing. Follow README OAuth setup and run --auth.")
    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")
    return build("youtube", "v3", credentials=creds)

def upload_to_youtube(video_path: Path, cfg: dict, title: str | None = None):
    youtube = youtube_service()
    clean_title = (title or video_path.stem.replace("_", " ").replace("-", " "))[:100]
    description = (
        f"{clean_title}\n\n"
        "In this video, the presentation slides are narrated in English.\n\n"
        + cfg.get("channel_description_suffix", "")
    )[:5000]
    body = {
        "snippet": {
            "title": clean_title,
            "description": description,
            "tags": cfg.get("youtube_tags", []),
            "categoryId": str(cfg.get("youtube_category_id", "27")),
            "defaultLanguage": cfg.get("youtube_language", "en")
        },
        "status": {
            "privacyStatus": cfg.get("youtube_privacy", "public"),
            "selfDeclaredMadeForKids": bool(cfg.get("made_for_kids", False)),
            "containsSyntheticMedia": bool(cfg.get("contains_synthetic_media", False))
        }
    }
    media = MediaFileUpload(str(video_path), mimetype="video/mp4", chunksize=8 * 1024 * 1024, resumable=True)
    request = youtube.videos().insert(part="snippet,status", body=body, media_body=media)
    response = None
    while response is None:
        status, response = request.next_chunk()
        if status:
            logging.info("YouTube upload: %.1f%%", status.progress() * 100)
    video_id = response["id"]
    logging.info("Uploaded: https://www.youtube.com/watch?v=%s", video_id)
    return video_id

def validate_pptx(pptx_path: Path) -> None:
    """
    Validate that the supplied file is a readable PowerPoint presentation.
    Raises ValueError with a user-friendly message if invalid.
    """

    if not pptx_path.exists():
        raise FileNotFoundError(
            f"PowerPoint file does not exist: {pptx_path}"
        )

    if not pptx_path.is_file():
        raise ValueError(
            f"Expected a file, but received: {pptx_path}"
        )

    if pptx_path.suffix.lower() != ".pptx":
        raise ValueError(
            f"Unsupported file format: {pptx_path.name}. "
            "Only .pptx files are supported."
        )

    if pptx_path.stat().st_size == 0:
        raise ValueError(
            f"{pptx_path.name} is empty."
        )

    try:
        Presentation(str(pptx_path))

    except Exception as exc:
        raise ValueError(
            f"{pptx_path.name} is not a valid or readable "
            "PowerPoint presentation. The file may be corrupted."
        ) from exc


def process_pptx(pptx_path: Path, cfg: dict):
    pptx_path = pptx_path.resolve()

    validate_pptx(pptx_path)

    work = (
        Path(cfg["work_dir"])
        / f"{pptx_path.stem}_{int(time.time())}_{uuid.uuid4().hex[:8]}"
    )
    work.mkdir(parents=True, exist_ok=True)

    scripts = extract_slide_scripts(pptx_path, cfg)

    (work / "scripts.json").write_text(
        json.dumps(scripts, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    images = render_slides(pptx_path, work)

    if len(images) != len(scripts):
        raise RuntimeError(
            f"Slide/image count mismatch: {len(scripts)} scripts "
            f"vs {len(images)} rendered slides."
        )

    
    # Generate the video and supporting files in the temporary work directory first.
    temp_out = work / f"{pptx_path.stem}.mp4"

    build_video(pptx_path, scripts, images, work, temp_out, cfg)

    # Only create the final output folder after video generation succeeds.
    presentation_output_dir = Path(cfg["output_dir"]) / pptx_path.stem
    presentation_output_dir.mkdir(parents=True, exist_ok=True)

    out = presentation_output_dir / f"{pptx_path.stem}.mp4"

    # Move the completed video and its supporting files to the final folder.
    shutil.move(str(temp_out), str(out))
    shutil.move(
        str(temp_out.with_suffix(".srt")),
        str(out.with_suffix(".srt"))
    )

    temp_thumbnail = temp_out.with_name(temp_out.stem + "_thumbnail.jpg")
    final_thumbnail = out.with_name(out.stem + "_thumbnail.jpg")

    shutil.move(str(temp_thumbnail), str(final_thumbnail))

    logging.info("Video created: %s", out)

    return {
        "video": str(out),
        "youtube_url": None,
    }

def authenticate():
    svc = youtube_service()
    channels = svc.channels().list(part="snippet", mine=True).execute()
    items = channels.get("items", [])
    if not items:
        raise RuntimeError("OAuth succeeded, but no YouTube channel was returned for this Google account.")
    for item in items:
        logging.info("Authenticated YouTube channel: %s (%s)", item["snippet"]["title"], item["id"])
    logging.info("OAuth token saved. You can now upload automatically.")

def process_next(upload=False, max_workers=2):
    cfg = load_config()
    setup_logging()

    job_registry.initialize_database()

    input_dir = Path(cfg["input_dir"])
    input_dir.mkdir(parents=True, exist_ok=True)

    candidates = sorted(
        p for p in input_dir.glob("*.pptx")
        if not p.name.startswith("~$")
    )

    if not candidates:
        logging.info("No .pptx files waiting in input_ppt.")
        return {
            "status": "idle",
            "message": "No PowerPoint files found."
        }

    max_workers = max(1, int(max_workers))

    logging.info(
        "Starting batch processing: %s presentations, %s workers.",
        len(candidates),
        max_workers,
    )
    
    jobs_to_process = []
    results = []
    videos_to_upload = []

    for pptx in candidates:
        file_hash = job_registry.calculate_file_hash(pptx)

        is_new, job_info = job_registry.register_job(
            filename=pptx.name,
            file_hash=file_hash,
            stored_path=pptx,
        )

        if is_new:
            jobs_to_process.append({
                "pptx": pptx,
                "job_id": job_info,
            })

            logging.info(
                "Registered new job %s for %s",
                job_info,
                pptx.name,
            )

        else:
            logging.warning(
                "Duplicate detected: %s. Existing job: %s",
                pptx.name,
                job_info["job_id"],
            )

            duplicates_dir = input_dir / "duplicates"

            duplicate_path = job_registry.move_job_file(
                pptx,
                duplicates_dir,
                uuid.uuid4().hex,
            )

            logging.info(
                "Moved duplicate file to: %s",
                duplicate_path,
            )

    with ThreadPoolExecutor(max_workers=max_workers) as executor:

        future_to_file = {
            executor.submit(
                process_pptx,
                job["pptx"],
                cfg,
            ): job
            for job in jobs_to_process
        }

        for future in as_completed(future_to_file):
            job = future_to_file[future]
            pptx = job["pptx"]
            job_id = job["job_id"]

            try:
                result = future.result()

                job_registry.update_job_status(
                    job_id,
                    "success",
                    output_path=Path(result["video"]),
                )

                logging.info("Successfully processed: %s", pptx.name)

                results.append({
                    "file": pptx.name,
                    "status": "success",
                    **result,
                })

                if cfg.get("delete_input_after_success", False):
                    pptx.unlink(missing_ok=True)

                    job_registry.update_job_stored_path(
                        job_id,
                        None,
                    )
                else:
                    done_dir = input_dir / "processed"

                    stored_path = job_registry.move_job_file(
                        pptx,
                        done_dir,
                        job_id,
                    )

                    job_registry.update_job_stored_path(
                        job_id,
                        stored_path,
                    )
                if upload:
                    job_registry.update_upload_status(
                        job_id,
                        "pending",
                    )
                    videos_to_upload.append({
                        "file": pptx.name,
                        "video": Path(result["video"]),
                        "job_id": job_id,
                    })               

            except Exception as exc:
                job_registry.update_job_status(
                    job_id,
                    "failed",
                )
                
                logging.error(
                    "Failed processing %s: %s",
                    pptx.name,
                    exc,
                )

                failed_dir = input_dir / "failed"

                try:
                    stored_path = job_registry.move_job_file(
                        pptx,
                        failed_dir,
                        job_id,
                    )

                    job_registry.update_job_stored_path(
                        job_id,
                        stored_path,
                    )

                except Exception:
                    logging.exception(
                        "Could not move failed file: %s",
                        pptx.name,
                    )

                results.append({
                    "file": pptx.name,
                    "status": "failed",
                    "error": str(exc),
                })
    
    
    if upload:
        for item in videos_to_upload:
            try:
                logging.info(
                    "Uploading to YouTube: %s",
                    item["file"],
                )

                job_registry.update_upload_status(
                    item["job_id"],
                    "uploading",
                )

                video_id = upload_to_youtube(item["video"], cfg)

            except Exception as exc:
                try:
                    job_registry.update_upload_status(
                        item["job_id"],
                        "failed",
                    )
                except Exception:
                    logging.exception(
                        "Could not update upload status for %s",
                        item["file"],
                    )

                logging.exception(
                    "YouTube upload failed for %s: %s",
                    item["file"],
                    exc,
                )

                for result in results:
                    if result["file"] == item["file"]:
                        result["upload_error"] = str(exc)
                        break

                continue

            youtube_url = f"https://www.youtube.com/watch?v={video_id}"

            # The YouTube upload has succeeded.
            try:
                job_registry.update_upload_status(
                    item["job_id"],
                    "success",
                )
            except Exception:
                logging.exception(
                    "Video uploaded, but database status update failed: %s",
                    item["file"],
                )

            # Record the URL in the processing result.
            for result in results:
                if result["file"] == item["file"]:
                    result["youtube_url"] = youtube_url
                    break

            # Save the URL locally as a separate operation.
            url_file = (
                item["video"].parent
                / f"{item['video'].stem}_youtube_url.txt"
            )

            try:
                url_file.write_text(
                    f"{youtube_url}\n",
                    encoding="utf-8",
                )
            except OSError:
                logging.exception(
                    "Video uploaded successfully, but URL file could not be saved: %s",
                    url_file,
                )

            logging.info(
                "Successfully uploaded %s: %s",
                item["file"],
                youtube_url,
            )

    successful = sum(
        1 for item in results if item["status"] == "success"
    )
    failed = sum(
        1 for item in results if item["status"] == "failed"
    )

    logging.info(
        "Batch completed. Successful: %s, Failed: %s",
        successful,
        failed,
    )

    return {
        "status": "completed",
        "total": len(candidates),
        "successful": successful,
        "failed": failed,
        "results": results,
    }


def main():
    parser = argparse.ArgumentParser(description="PowerPoint to narrated YouTube video agent")
    parser.add_argument("--input", help="Process one .pptx file")
    upload_group = parser.add_mutually_exclusive_group()

    upload_group.add_argument(
        "--upload",
        action="store_true",
        help="Upload generated videos to YouTube",
    )

    upload_group.add_argument(
        "--no-upload",
        action="store_true",
        help="Do not upload videos (default)",
    )
    parser.add_argument("--auth", action="store_true", help="Authenticate the YouTube channel")
    parser.add_argument("--process-next", action="store_true", help="Process all .pptx files in input_ppt")
    args = parser.parse_args()
    setup_logging()
    cfg = load_config()
    if args.auth:
        authenticate()
    elif args.input:
        result = process_pptx(Path(args.input), cfg)
        print(json.dumps(result, indent=2))
    elif args.process_next:
        max_workers = int(cfg.get("max_workers", 2))

        result = process_next(
            upload=args.upload,
            max_workers=max_workers,
        )

        print(json.dumps(result, indent=2))
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
