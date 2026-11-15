from __future__ import annotations
import argparse, json, logging, re, shutil, subprocess, sys, time, zipfile
from pathlib import Path
import os
from xml.etree import ElementTree as ET

import fitz
import edge_tts
import asyncio
from pptx import Presentation
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

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
    pdf_dir.mkdir(parents=True, exist_ok=True)
    img_dir.mkdir(parents=True, exist_ok=True)
    soffice = find_soffice()
    cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(pptx_path)]
    logging.info("Rendering PowerPoint to PDF...")
    subprocess.run(cmd, check=True, capture_output=True, text=True)
    pdf_path = pdf_dir / (pptx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError("LibreOffice did not create the expected PDF. Check whether the PPTX opens correctly.")
    doc = fitz.open(pdf_path)
    images = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(1.8, 1.8), alpha=False)
        out = img_dir / f"slide_{i+1:03d}.jpg"
        pix.save(out, jpg_quality=90)
        images.append(out)
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

def process_pptx(pptx_path: Path, cfg: dict, upload: bool = True):
    pptx_path = pptx_path.resolve()
    if not pptx_path.exists() or pptx_path.suffix.lower() != ".pptx":
        raise FileNotFoundError(f"Expected an existing .pptx file: {pptx_path}")
    work = Path(cfg["work_dir"]) / f"{pptx_path.stem}_{int(time.time())}"
    work.mkdir(parents=True, exist_ok=True)
    scripts = extract_slide_scripts(pptx_path, cfg)
    (work / "scripts.json").write_text(json.dumps(scripts, ensure_ascii=False, indent=2), encoding="utf-8")
    images = render_slides(pptx_path, work)
    if len(images) != len(scripts):
        raise RuntimeError(f"Slide/image count mismatch: {len(scripts)} scripts vs {len(images)} rendered slides.")
    out = Path(cfg["output_dir"]) / f"{pptx_path.stem}.mp4"
    build_video(pptx_path, scripts, images, work, out, cfg)
    logging.info("Video created: %s", out)
    if upload:
        video_id = upload_to_youtube(out, cfg)
        (Path(cfg["output_dir"]) / f"{pptx_path.stem}_youtube_url.txt").write_text(
            f"https://www.youtube.com/watch?v={video_id}\n", encoding="utf-8"
        )
        return {"video": str(out), "youtube_url": f"https://www.youtube.com/watch?v={video_id}"}
    return {"video": str(out), "youtube_url": None}

def authenticate():
    svc = youtube_service()
    channels = svc.channels().list(part="snippet", mine=True).execute()
    items = channels.get("items", [])
    if not items:
        raise RuntimeError("OAuth succeeded, but no YouTube channel was returned for this Google account.")
    for item in items:
        logging.info("Authenticated YouTube channel: %s (%s)", item["snippet"]["title"], item["id"])
    logging.info("OAuth token saved. You can now upload automatically.")

def process_next():
    cfg = load_config()
    setup_logging()
    input_dir = Path(cfg["input_dir"])
    input_dir.mkdir(parents=True, exist_ok=True)
    candidates = sorted(p for p in input_dir.glob("*.pptx") if not p.name.startswith("~$"))
    if not candidates:
        logging.info("No .pptx files waiting in input_ppt.")
        return {"status": "idle", "message": "No PowerPoint file found."}
    results = []
    for pptx in candidates:
        try:
            logging.info("Starting: %s", pptx.name)
            result = process_pptx(pptx, cfg, upload=True)
            results.append({"file": pptx.name, **result})
            if cfg.get("delete_input_after_success", False):
                pptx.unlink(missing_ok=True)
            else:
                done_dir = input_dir / "processed"
                done_dir.mkdir(exist_ok=True)
                shutil.move(str(pptx), str(done_dir / pptx.name))
        except Exception:
            logging.exception("Failed processing %s", pptx.name)
            failed_dir = input_dir / "failed"
            failed_dir.mkdir(exist_ok=True)
            try:
                shutil.move(str(pptx), str(failed_dir / pptx.name))
            except Exception:
                pass
            raise
    return {"status": "completed", "results": results}

def main():
    parser = argparse.ArgumentParser(description="PowerPoint to narrated YouTube video agent")
    parser.add_argument("--input", help="Process one .pptx file")
    parser.add_argument("--no-upload", action="store_true", help="Render MP4 but do not upload")
    parser.add_argument("--auth", action="store_true", help="Authenticate the YouTube channel")
    parser.add_argument("--process-next", action="store_true", help="Process all .pptx files in input_ppt")
    args = parser.parse_args()
    setup_logging()
    cfg = load_config()
    if args.auth:
        authenticate()
    elif args.input:
        result = process_pptx(Path(args.input), cfg, upload=not args.no_upload)
        print(json.dumps(result, indent=2))
    elif args.process_next:
        print(json.dumps(process_next(), indent=2))
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
