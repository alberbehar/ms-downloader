# MS Downloader

Chrome / Edge extension that saves transcripts of Microsoft Teams, SharePoint and Stream recordings you can already watch.

> **Version 2 – rewritten from scratch.** Step 1: transcripts. Video and audio download are coming next.

## Features
- One-click **Transcript** button on SharePoint / Stream recording pages and on recordings opened inside Teams
- Formats: readable text (.txt), WebVTT (.vtt), SubRip (.srt) and the original JSON
- Optional speaker names and timestamps, live preview, editable file name
- Remembers your last choices
- Everything runs in your browser – nothing is sent anywhere else

## Install (developer mode)
1. Download this repository (**Code → Download ZIP**) and unzip it.
2. Open `chrome://extensions` and turn on **Developer mode**.
3. Click **Load unpacked** and choose the `extension` folder.

## Use
1. Open a meeting recording (Teams, SharePoint or Stream).
2. Click the blue **Transcript** button in the bottom-right corner.
3. Pick a format and press **Download**.

## Limits
- Only works for recordings you are allowed to view.
- The meeting must have been transcribed.

## Project layout
```
extension/
├── manifest.json      Manifest V3
├── page-hook.js       Reads the file id and API token from the page
├── panel.js           Button, dialog and downloads
├── lib/transcript.js  Transcript → TXT / VTT / SRT
├── popup.html/.js     Toolbar popup
└── icons/
```

## License
MIT © 2026 alberbehar
