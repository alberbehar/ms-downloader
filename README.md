# MS Downloader

Chrome extension to download Microsoft Teams / SharePoint / Stream meeting recordings (video, audio) and transcripts — even when the built-in download button is disabled.

Works on `teams.microsoft.com`, `teams.cloud.microsoft` and `*.sharepoint.com` (including the Stream-on-SharePoint player).

## Features
- Download **Video + audio (MP4)**, **Audio only (M4A)** or **Video only**, entirely in the browser
- Download transcripts as **JSON**, **WebVTT** or **Grouped VTT** (with optional timestamps and speaker names)
- Optional subtitles embedded in video downloads

## Installation (load unpacked)
1. Clone or download this repository.
2. Open `chrome://extensions/` and turn on **Developer mode**.
3. Click **Load unpacked** and select the `src/` folder.

## Usage
Open a recording in Teams / SharePoint / Stream, click **Download** in the command bar and choose a format.

## Notes
- Only works for content you can already view; it does not bypass access restrictions or DRM.
- All processing happens locally in the browser.

## Credits & License
Based on [ms-teams-sharepoint-downloader](https://github.com/brendangooden/ms-teams-sharepoint-downloader) by Brendan Gooden, released under the MIT License. See [LICENSE](LICENSE).

## Website
The project website lives in [`website/`](website/) (Astro). Run `npm install` and `npm run dev` inside it to preview, `npm run build` to build.
