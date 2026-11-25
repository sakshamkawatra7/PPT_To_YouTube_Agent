from pathlib import Path
import shutil, uuid
from fastapi import FastAPI, UploadFile, File, HTTPException
from ppt_youtube_agent import load_config, process_next, process_pptx, setup_logging

app = FastAPI(title="PPT-to-YouTube Agent", version="1.0")
cfg = load_config()
setup_logging()

@app.get("/health")
def health():
    return {"status": "ok", "service": "ppt-to-youtube-agent"}

@app.post("/process-next")
def run_queue():
    try:
        return process_next()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))

@app.post("/upload-and-process")
def upload_and_publish(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Upload a .pptx PowerPoint file.")
    input_dir = Path(cfg["input_dir"])
    input_dir.mkdir(parents=True, exist_ok=True)
    safe_name = Path(file.filename).name
    destination = input_dir / f"{uuid.uuid4().hex[:8]}_{safe_name}"
    with destination.open("wb") as out:
        shutil.copyfileobj(file.file, out)
    try:
        result = process_pptx(destination, cfg)
        processed = input_dir / "processed"
        processed.mkdir(exist_ok=True)
        shutil.move(str(destination), str(processed / destination.name))
        return {"status": "processed", "result": result}
    except Exception as exc:
        failed = input_dir / "failed"
        failed.mkdir(exist_ok=True)
        if destination.exists():
            shutil.move(str(destination), str(failed / destination.name))
        raise HTTPException(status_code=500, detail=str(exc))
