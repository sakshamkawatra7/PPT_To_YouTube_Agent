\
# PPT-to-YouTube Agent (Windows, local-first)

This project takes `.pptx` files from `input_ppt/`, reads speaker notes (or falls back to visible slide text), generates Indian-English narration, renders each slide to a video segment, joins the segments into an MP4, and uploads the video to your YouTube channel.

## What it costs

- Python, LibreOffice, FFmpeg and n8n self-hosted: free.
- YouTube Data API: no per-upload fee, subject to Google's quota and API policies.
- `edge-tts`: uses Microsoft's online speech service; no API key is configured here, but it requires internet and is not a guaranteed SLA or an offline/local TTS engine. If the service changes, replace the TTS function with a locally licensed voice engine.
- Your Windows PC and internet connection must be on while processing.

## 1. Install prerequisites

1. Install Python 3.11 (64-bit): https://www.python.org/downloads/windows/
   - During setup tick **Add python.exe to PATH**.
2. Install LibreOffice: https://www.libreoffice.org/download/download/
3. Install FFmpeg. One simple Windows route is:
   ```powershell
   winget install Gyan.FFmpeg
   ```
   Close and reopen PowerShell, then verify:
   ```powershell
   ffmpeg -version
   ffprobe -version
   ```
4. Install Node.js LTS if you want n8n: https://nodejs.org/en/download
5. Optional but recommended: install Git.

## 2. Set up the Python agent

Open PowerShell in this folder:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

If PowerShell blocks activation, use:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```
Then activate again.

## 3. Configure YouTube OAuth

1. Open https://console.cloud.google.com/
2. Create a Google Cloud project (or select one).
3. Go to **APIs & Services → Library**, search for **YouTube Data API v3**, and enable it.
4. Go to **Google Auth Platform → Audience**. Configure the consent screen. For a personal-only tool, choose **External** if required and add your own Google account as a test user while the app is in Testing. **Important:** Google says refresh tokens for External apps left in Testing expire after 7 days when non-basic scopes such as YouTube upload are requested. For persistent unattended publishing, review the consent-screen publishing/verification requirements and move the app out of Testing where eligible; this may involve Google review. Until then, expect to re-authorize periodically.
5. Go to **APIs & Services → Credentials → Create Credentials → OAuth client ID**.
6. Choose **Desktop app**. Download the JSON file.
7. Rename the downloaded file to `credentials.json` and put it in this project folder.
8. Run:
   ```powershell
   .\authenticate_youtube.bat
   ```
9. A browser opens. Sign in with the Google account that owns the YouTube channel and grant the `youtube.upload` permission. The app stores `token.json` locally.

Keep `credentials.json` and `token.json` private. Never upload them to GitHub or share them. If a token is exposed, revoke it from your Google account's third-party access page.

**API verification note:** Google can impose restrictions on uploads from unverified API projects. If uploads are rejected or forced to private, review the current YouTube API compliance/audit requirements in the official documentation. Do not assume changing the code can bypass Google's project verification.

## 4. Prepare your PowerPoint

Best results:
- Save as `.pptx` (not legacy `.ppt`).
- Put one complete narration script in each slide's **Notes** area (PowerPoint → View → Notes, or the Notes pane).
- Use natural spoken English, not only short slide bullets.
- Avoid relying on animations: the renderer exports slide appearance, not timed PowerPoint animations.
- Avoid password-protected files.

If a slide has no speaker notes, the agent narrates visible slide text. If both are empty, it says "Slide N."

## 5. Test without publishing

Set `youtube_privacy` to `unlisted` in `config.json` for your first test. Then:

```powershell
python ppt_youtube_agent.py --input "C:\full\path\to\your.pptx" --no-upload
```

Inspect the MP4 in `output_videos/`. Confirm slide order, narration, timings, text pronunciation and audio quality.

## 6. Enable automatic publishing

After the test, set `youtube_privacy` in `config.json` to `public` (already set in the supplied starter configuration). Confirm `youtube_tags`, `youtube_category_id`, `made_for_kids` and `contains_synthetic_media` are accurate for your channel/content.

Put one or more `.pptx` files into `input_ppt/`, then double-click `run_once.bat`. Successful files move to `input_ppt/processed/`; failed files move to `input_ppt/failed/`. Videos are saved in `output_videos/`.

The default is public publishing. This is intentionally automatic. A wrong title, copyright issue, incorrect audience declaration or inaccurate synthetic-media setting can therefore go live; review the first few videos carefully.

## 7. Start the local HTTP service

Double-click `start_agent.bat`. The local API listens only on `127.0.0.1:8000`.

Test:
```powershell
Invoke-RestMethod http://127.0.0.1:8000/health
```

To process files already placed in `input_ppt/`:
```powershell
Invoke-RestMethod -Method Post http://127.0.0.1:8000/process-next
```

To upload one PPTX through the API:
```powershell
curl.exe -X POST -F "file=@C:\path\to\presentation.pptx" http://127.0.0.1:8000/upload-and-publish
```

Do not expose this service directly to the public internet. It has no user authentication; keep it on localhost.

## 8. n8n workflow (optional)

This package includes `n8n_workflow.json`.

1. Keep `start_agent.bat` running.
2. Install n8n locally:
   ```powershell
   npm install -g n8n@2
   ```
3. Start it:
   ```powershell
   n8n
   ```
4. Open http://localhost:5678
5. Choose **Workflows → Import from File**, then select `n8n_workflow.json`.
6. Open the workflow and check that the HTTP Request node points to `http://127.0.0.1:8000/process-next`.
7. Save and activate the workflow.

Every 10 minutes n8n asks the local Python service to process any waiting PPTX files. You still place PPTX files in `input_ppt/`. This approach avoids n8n's Execute Command node and avoids Docker path-mapping problems.

Important: n8n must run on the same Windows computer as the Python service. The included workflow targets n8n 2.x. The n8n project documentation notes that npm installs are deprecated starting with n8n 3.0, so use its current Docker-based self-hosting instructions if you install n8n 3.x or later. Do not use n8n Cloud with `127.0.0.1`; that address would refer to n8n's remote machine, not your PC. If your n8n version requires a different Schedule Trigger interval UI, set the same 10-minute cadence manually.

## 9. Voice options

Default: `en-IN-NeerjaNeural` (Indian English female voice).
Alternative: `en-IN-PrabhatNeural` (Indian English male voice).

Change `voice` in `config.json`. TTS is generated from the slide notes; it is not a recording of your own voice.

## 10. What this starter does not do yet

- It does not write original scripts or research topics from your Excel sheet.
- It does not create Shorts, animated graphics, B-roll or branded intros.
- It does not upload to Instagram/Facebook/LinkedIn.
- It does not monitor comments or channel analytics.
- It does not guarantee monetization or channel growth.
- It does not reliably preserve complex PowerPoint animations, embedded video or slide transitions.

Those are separate stages. First validate this basic PPT-to-video-to-YouTube pipeline.

## Troubleshooting

- **`LibreOffice not found`:** install LibreOffice; ensure `soffice.exe` is in PATH or at the standard Program Files location.
- **`FFmpeg not found`:** install FFmpeg and reopen PowerShell so PATH refreshes.
- **TTS error:** check internet access; try the alternate voice in `config.json`.
- **No script in narration:** add text to the slide Notes pane. The parser uses notes when present, otherwise visible text.
- **OAuth error:** verify that YouTube Data API v3 is enabled, the OAuth client type is Desktop app, and the signed-in account owns/has access to the channel.
- **Upload quota / forbidden:** check Google Cloud quota, channel restrictions and API verification status. Retry transient errors; do not repeatedly submit the same file blindly.

## Official references

- YouTube Data API upload method: https://developers.google.com/youtube/data/reference/rest/v3/videos/insert
- YouTube API Python quickstart: https://developers.google.com/youtube/data/quickstart/python
- Uploading a video guide: https://developers.google.com/youtube/v3/guides/uploading_a_video
- n8n documentation: https://docs.n8n.io/
- LibreOffice: https://www.libreoffice.org/
- FFmpeg: https://ffmpeg.org/
