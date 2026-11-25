import sqlite3
from pathlib import Path
import hashlib
from datetime import datetime, timedelta, timezone
import uuid
from contextlib import contextmanager

DB_PATH = Path(__file__).resolve().parent / "jobs.db"


def calculate_file_hash(file_path: Path) -> str:
    """
    Calculate the SHA-256 hash of a file.

    Reads the file in chunks to avoid loading
    the entire presentation into memory.
    """
    sha256 = hashlib.sha256()

    with open(file_path, "rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            sha256.update(chunk)

    return sha256.hexdigest()



def move_job_file(
    source_path: Path,
    destination_dir: Path,
    job_id: str,
) -> Path:
    """
    Move a job's input file to a destination directory
    without overwriting an existing file.

    The job ID provides a unique filename.
    """
    import shutil

    source_path = source_path.resolve()
    destination_dir = destination_dir.resolve()

    destination_dir.mkdir(parents=True, exist_ok=True)

    destination_path = (
        destination_dir
        / f"{source_path.stem}_{job_id[:8]}{source_path.suffix}"
    )

    if destination_path.exists():
        raise FileExistsError(
            f"Destination already exists: {destination_path}"
        )

    shutil.move(str(source_path), str(destination_path))

    return destination_path


@contextmanager
def get_connection():
    """
    Create a SQLite connection and ensure it is
    committed/rolled back and closed properly.
    """
    connection = sqlite3.connect(DB_PATH, timeout=30)
    connection.row_factory = sqlite3.Row

    try:
        yield connection
        connection.commit()

    except Exception:
        connection.rollback()
        raise

    finally:
        connection.close()



def initialize_database():
    """
    Create the jobs table if it does not already exist.
    Add new columns when upgrading an existing database.
    """
    with get_connection() as connection:
        connection.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                job_id TEXT PRIMARY KEY,
                filename TEXT NOT NULL,
                file_hash TEXT NOT NULL,
                status TEXT NOT NULL,
                created_at TEXT NOT NULL,
                output_path TEXT,
                stored_path TEXT
            )
        """)

        # Check which columns already exist.
        columns = {
            row["name"]
            for row in connection.execute(
                "PRAGMA table_info(jobs)"
            ).fetchall()
        }

        # Upgrade existing databases without deleting records.
        if "stored_path" not in columns:
            connection.execute("""
                ALTER TABLE jobs
                ADD COLUMN stored_path TEXT
            """)
        if "upload_status" not in columns:
            connection.execute("""
            ALTER TABLE jobs
            ADD COLUMN upload_status TEXT NOT NULL DEFAULT 'not_requested'
        """)

        connection.execute("""
            CREATE INDEX IF NOT EXISTS
            idx_jobs_hash_created
            ON jobs (file_hash, created_at)
        """)


def register_job(
    filename: str,
    file_hash: str,
    stored_path: Path,
):
    """
    Atomically check for a recent duplicate and
    register a new processing job if none exists.

    Returns:
        (True, job_id) if registered.
        (False, existing_job) if duplicate found.
    """
    cutoff_time = (
        datetime.now(timezone.utc) - timedelta(hours=24)
    ).isoformat()

    job_id = str(uuid.uuid4())

    with get_connection() as connection:
        # Acquire a write lock before checking for duplicates.
        connection.execute("BEGIN IMMEDIATE")

        existing_job = connection.execute(
            """
            SELECT *
            FROM jobs
            WHERE file_hash = ?
              AND created_at >= ?
              AND status IN ('processing', 'success')
            ORDER BY created_at DESC
            LIMIT 1
            """,
            (file_hash, cutoff_time),
        ).fetchone()

        if existing_job:
            return False, dict(existing_job)

        connection.execute(
            """
            INSERT INTO jobs (
                job_id,
                filename,
                file_hash,
                status,
                created_at,
                stored_path
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                job_id,
                filename,
                file_hash,
                "processing",
                datetime.now(timezone.utc).isoformat(),
                str(stored_path.resolve()),
            ),
        )

    return True, job_id



def update_job_status(
    job_id: str,
    status: str,
    output_path: str | None = None,
):
    """
    Update the status of a registered job.
    Optionally store the generated video path.
    """
    allowed_statuses = {"processing", "success", "failed"}

    if status not in allowed_statuses:
        raise ValueError(f"Invalid job status: {status}")

    with get_connection() as connection:
        cursor = connection.execute(
            """
            UPDATE jobs
            SET status = ?,
                output_path = ?
            WHERE job_id = ?
            """,
            (status, output_path, job_id),
        )

        if cursor.rowcount == 0:
            raise ValueError(f"Job not found: {job_id}")



def find_recent_failed_job(file_hash: str):
    """
    Find the most recent failed job with identical
    file contents within the previous 24 hours.
    """
    cutoff_time = (
        datetime.now(timezone.utc) - timedelta(hours=24)
    ).isoformat()

    with get_connection() as connection:
        job = connection.execute(
            """
            SELECT *
            FROM jobs
            WHERE file_hash = ?
              AND created_at >= ?
              AND status = 'failed'
            ORDER BY created_at DESC
            LIMIT 1
            """,
            (file_hash, cutoff_time),
        ).fetchone()

    return dict(job) if job else None

def find_recent_duplicate(file_hash: str):
    """
    Find an identical presentation submitted within
    the previous 24 hours.

    Failed jobs are excluded because they can be retried.
    """
    cutoff_time = (
        datetime.now(timezone.utc) - timedelta(hours=24)
    ).isoformat()

    with get_connection() as connection:
        job = connection.execute(
            """
            SELECT *
            FROM jobs
            WHERE file_hash = ?
              AND created_at >= ?
              AND status IN ('processing', 'success')
            ORDER BY created_at DESC
            LIMIT 1
            """,
            (file_hash, cutoff_time),
        ).fetchone()

    return dict(job) if job else None


def update_job_stored_path(
    job_id: str,
    stored_path: Path | None,
):
    """
    Update the current location of a job's input file.
    Use None if the input file has been deleted.
    """
    path_value = (
        str(stored_path.resolve())
        if stored_path is not None
        else None
    )

    with get_connection() as connection:
        cursor = connection.execute(
            """
            UPDATE jobs
            SET stored_path = ?
            WHERE job_id = ?
            """,
            (path_value, job_id),
        )

        if cursor.rowcount == 0:
            raise ValueError(f"Job not found: {job_id}")


def update_upload_status(job_id: str, upload_status: str):
    """
    Update the YouTube upload status independently
    of the video processing status.
    """
    allowed_statuses = {
        "not_requested",
        "pending",
        "uploading",
        "success",
        "failed",
    }

    if upload_status not in allowed_statuses:
        raise ValueError(
            f"Invalid upload status: {upload_status}"
        )

    with get_connection() as connection:
        cursor = connection.execute(
            """
            UPDATE jobs
            SET upload_status = ?
            WHERE job_id = ?
            """,
            (upload_status, job_id),
        )

        if cursor.rowcount == 0:
            raise ValueError(f"Job not found: {job_id}")
