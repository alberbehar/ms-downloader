# MS Downloader

Chrome / Edge extension that downloads Microsoft Teams, SharePoint and Stream recordings – video, audio and transcripts – that you can already watch, even when the download button is turned off.

> **Version 2 – rewritten from scratch.**

## Features
- A **Download** button on SharePoint / Stream recording pages and on recordings opened inside Teams
- **Video + audio** (MP4), **Audio only** (M4A) or **Video only** (MP4), in the best quality the player offers
- Downloads in the browser with parallel requests and automatic retry when Microsoft throttles
- Handles AES-128 encrypted streams (keys the player itself receives); copy-protected (DRM) recordings are detected and refused
- **Transcript** as readable text (.txt), WebVTT (.vtt), SubRip (.srt) or the original JSON, with optional speaker names and timestamps and a live preview
- Remembers your last choices
- Everything runs in your browser – nothing is sent anywhere else

## Install (developer mode)
1. Download this repository (**Code → Download ZIP**) and unzip it.
2. Open `chrome://extensions` and turn on **Developer mode**.
3. Click **Load unpacked** and choose the `extension` folder.

## Use
1. Open a meeting recording (Teams, SharePoint or Stream).
2. Press play for a few seconds so the player loads the video.
3. Click the blue **Download** button in the bottom-right corner and pick **Video + audio**, **Audio only**, **Video only** or **Transcript**.
4. Press **Download** in the dialog. The file goes to your Downloads folder.

## Limits
- Only works for recordings you are allowed to view.
- DRM-protected recordings cannot be saved.
- Transcripts exist only if transcription was on during the meeting.

## Project layout
```
extension/
├── manifest.json      Manifest V3
├── page-hook.js       Reads the file id, video manifest and tokens from the page
├── panel.js           Button, menu, dialogs and downloads
├── lib/dash.js        Reads the DASH manifest
├── lib/mp4.js         Joins video and audio into one MP4
├── lib/transcript.js  Transcript → TXT / VTT / SRT
├── popup.html/.js     Toolbar popup
└── icons/
```

## License
MIT © 2026 alberbehar
