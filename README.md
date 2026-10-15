# MicCamGlow

A small Windows 11 tool that turns your taskbar into an "on air" light.

- **Red, pulsing** while any app has your microphone or camera open.
- **Green, steady** while you are in a Microsoft Teams meeting and muted.
- **Off** when nothing is using the mic or camera.

A bright label on the taskbar says what is going on, for example `MIC LIVE (Teams)`, `CAMERA ON (Zoom)` or
`MUTED (Teams)`. The overlay is click-through, never steals focus, and hides while a full-screen app or a
presentation is running.

It is a single PowerShell script with some embedded C#. No installer, no admin rights, no network access.

## Requirements

- Windows 11 (Windows 10 should work but is untested)
- Windows PowerShell 5.1 (built in)
- Microsoft Teams (new Teams) for the muted/live distinction; other apps get the red "in use" signal only

## Quick start

```powershell
# try it in the foreground (Ctrl+C to stop); forces the red glow for 8 seconds so you can see it
powershell -ExecutionPolicy Bypass -File .\MicCamGlow.ps1 -DemoSeconds 8

# install: adds a hidden Startup shortcut for the current user and starts the watcher
powershell -ExecutionPolicy Bypass -File .\MicCamGlow.ps1 -Install
```

Keep the script somewhere permanent before running `-Install`, because the Startup shortcut points at its path.

## Commands

| Switch | What it does |
|---|---|
| `-Status` | Shows whether the watcher runs, who holds the mic or camera, and recent log lines |
| `-Install` | Writes a Startup shortcut (hidden launch via `conhost --headless`) and starts the watcher |
| `-Uninstall` | Removes the Startup shortcut and stops the watcher |
| `-Stop` | Stops the watcher |
| `-DemoSeconds N` | Forces the red glow for N seconds (look test) |

Options (pass them together with `-Install` to keep them in the shortcut):

| Option | Effect |
|---|---|
| `-Ignore 'App1','App2'` | Ignore apps whose registry key contains these strings, for example a dictation tool |
| `-LabelColor '#FFFFFF'` | Label text color (default bright yellow `#FFE600`) |
| `-Color`, `-MutedColor` | Glow colors (default red `#E61E1E`, green `#1E9E4A`) |
| `-NoLabel` | Glow only, no text |
| `-NoMutedIndicator` | No green strip while muted |
| `-NoUia` | Do not read the Teams window (mic "in use" only) |

## How it works

1. **Is the device open?** Windows records, per app, when it starts and stops using the microphone or webcam,
   under `HKCU\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore`. This is the same data
   behind the mic icon in the system tray. An app with a start time and no later stop time is holding the device.

2. **Are you muted in Teams?** Teams keeps the mic open for the whole meeting, even while you are muted, so step 1
   alone cannot tell muted from live. While Teams holds the mic, the script reads the meeting window through
   Windows UI Automation (the accessibility API). The mic button is named `Mute mic` while you are live and
   `Unmute mic` while you are muted.

The mute reading can only turn red into green. If the button can't be read (a non-English Teams UI, a renamed
button), you get red for the whole meeting with `MIC OPEN (Teams, mute unknown)` rather than a false "muted".

### About the Teams third-party API

Teams used to offer a local WebSocket API (Settings > Privacy > Third-party app API) that reported mute state.
Microsoft retired it in December 2025. The client code is still in the script behind an opt-in `-TeamsApi` switch in
case a replacement appears; by default it is off.

## Files it creates

| Path | Purpose |
|---|---|
| `%LOCALAPPDATA%\MicCamGlow\miccamglow.log` | State changes and diagnostics (trimmed at 1 MB) |
| `shell:startup\MicCamGlow.lnk` | Only after `-Install` |

## Troubleshooting

- **Glow stays red while muted in Teams:** run `-Status` and check the log. If it shows
  `meeting window found but no mic button matched; buttons: ...`, your Teams uses different button names; adjust
  `MicRx` / `LeaveRx` in the script.
- **Vertical taskbar:** supported; the label runs down the bar.
- **CPU use:** about 3% of one core while glowing, near zero otherwise.

## License

MIT, see [LICENSE](LICENSE).
