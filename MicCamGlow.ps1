<#
.SYNOPSIS
  Pulses a red glow over the Windows 11 taskbar whenever your microphone or
  camera is live, and (with the Teams API linked) shows a calm green strip
  while you are muted in a Teams meeting.

.DESCRIPTION
  Two signals are combined:

  1. Windows' own per-app "device in use" record under
     HKCU\...\CapabilityAccessManager\ConsentStore (the data behind the mic
     icon in the system tray). Any app holding the mic or webcam = RED.
     Limitation: Teams keeps the mic OPEN for the whole meeting even while you
     are muted inside Teams, so this signal alone cannot tell muted from live.

  2. UI Automation: the Mute/Unmute button of the Teams meeting window (the
     same "interface discovery" surviving tools such as MuteDeck use). Read
     only while Teams holds the mic. English Teams UI assumed.
     "Unmute..." button => you are muted => GREEN strip; "Mute..." => RED.

  3. (Opt-in, -TeamsApi) The Teams third-party app API websocket on port 8124.
     Microsoft RETIRED it in December 2025 (Message Center MC1266901), so it is
     off by default and kept only in case a replacement appears.

  Signals 2 and 3 are only trusted to CANCEL the red; signal 1 always wins
  otherwise ("MIC OPEN (Teams, mute unknown)").

  The overlay is a click-through, never-activated window sized to each
  taskbar; clicks pass straight through to the taskbar.

.PARAMETER Install        Add a Startup shortcut (hidden launch) and start now.
.PARAMETER Uninstall      Remove the Startup shortcut and stop the watcher.
.PARAMETER Stop           Stop the running watcher.
.PARAMETER Status         Show watcher / Teams-link / device state and exit.
.PARAMETER TeamsApi       Also try the (retired) Teams third-party app API on port 8124.
.PARAMETER TeamsToken     Save a token for that API and restart the watcher.
.PARAMETER NoUia          Do not read the Teams meeting window via UI Automation.
.PARAMETER NoMutedIndicator  Do not show the green "MUTED" strip.
.PARAMETER NoLabel        Do not draw text on the strip.
.PARAMETER Ignore         Substrings of app keys to ignore (e.g. 'Wispr Flow').
.PARAMETER DemoSeconds    Force the red glow on for N seconds after start (look test).
#>
[CmdletBinding()]
param(
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$Stop,
    [switch]$Status,
    [string]$TeamsToken,
    [switch]$TeamsApi,
    [switch]$NoUia,
    [switch]$NoMutedIndicator,
    [switch]$NoLabel,
    [string[]]$Ignore = @(),
    [int]$DemoSeconds = 0,
    [int]$PollMs = 500,
    [int]$TeamsPort = 8124,
    [double]$MinOpacity = 0.22,
    [double]$MaxOpacity = 0.62,
    [double]$MutedOpacity = 0.20,
    [double]$PulseSeconds = 1.4,
    [string]$Color = '#E61E1E',
    [string]$MutedColor = '#1E9E4A',
    [string]$LabelColor = '#FFE600'
)

$ErrorActionPreference = 'Stop'
$ScriptPath   = $MyInvocation.MyCommand.Path
$ShortcutPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'MicCamGlow.lnk'
$DataDir      = Join-Path $env:LOCALAPPDATA 'MicCamGlow'
$TokenPath    = Join-Path $DataDir 'teams-token.txt'
$LogPath      = Join-Path $DataDir 'miccamglow.log'
$PsExe        = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$ConhostExe   = Join-Path $env:SystemRoot 'System32\conhost.exe'
if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Path $DataDir -Force | Out-Null }

function Write-Log([string]$Message) {
    # .NET append on purpose: Add-Content refuses the file while a reader such as tail -f holds it open.
    try { [System.IO.File]::AppendAllText($LogPath, ("{0} [glow] {1}`r`n" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)) } catch { }
}

function Get-WatcherProcesses {
    Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe' OR Name = 'pwsh.exe'" |
        Where-Object { $_.CommandLine -and $_.CommandLine -like '*MicCamGlow.ps1*' -and $_.ProcessId -ne $PID } |
        Where-Object { $_.CommandLine -notmatch '-(Install|Uninstall|Stop|Status|TeamsToken)\b' }
}

function Get-LaunchArguments {
    $a = "--headless `"$PsExe`" -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
    if ($NoLabel)          { $a += ' -NoLabel' }
    if ($NoMutedIndicator) { $a += ' -NoMutedIndicator' }
    if ($TeamsApi)         { $a += ' -TeamsApi' }
    if ($NoUia)            { $a += ' -NoUia' }
    if ($Ignore.Count)     { $a += ' -Ignore ' + (($Ignore | ForEach-Object { "'$_'" }) -join ',') }
    return $a
}

function Start-Watcher {
    if (@(Get-WatcherProcesses).Count) { return 'Watcher already running.' }
    Start-Process -FilePath $ConhostExe -ArgumentList (Get-LaunchArguments) -WindowStyle Hidden
    Start-Sleep -Seconds 3
    if (@(Get-WatcherProcesses).Count) { 'Watcher started.' } else { 'Watcher did NOT start; run the script with no switches to see errors.' }
}

function Stop-Watcher {
    $procs = @(Get-WatcherProcesses)
    foreach ($p in $procs) { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue }
    "Stopped $($procs.Count) watcher process(es)."
}

function ConvertTo-FriendlyAppName([string]$Key) {
    if ($Key -match '#') { return [IO.Path]::GetFileNameWithoutExtension(($Key -split '#')[-1]) }   # C:#Program Files#Zoom#Zoom.exe
    $base = ($Key -split '_')[0]
    switch -Regex ($base) {
        '^MSTeams$'                       { return 'Teams' }
        '^MicrosoftWindows\.Client\.CBS$' { return 'Windows' }
        '^windows\.immersivecontrolpanel' { return 'Settings' }
        default                           { return ($base -split '\.')[-1] }
    }
}

# ---------------------------------------------------------------------------
# Compiled helpers: registry probe, overlay window + pulse engine, Teams link
# ---------------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, UIAutomationClient, UIAutomationTypes, WindowsBase
$refs = @('System.Windows.Forms', 'System.Drawing', 'System.Web.Extensions',
          [System.Windows.Automation.AutomationElement].Assembly.Location,
          [System.Windows.Automation.ControlType].Assembly.Location,
          [System.Windows.Rect].Assembly.Location)
Add-Type -ReferencedAssemblies $refs -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Net.WebSockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using Microsoft.Win32;

namespace MicCamGlow {
    // Admin-free fallback: read the Mute/Unmute button of the Teams meeting window via UI Automation.
    // Only scans while enabled (Teams holds the mic and the API is not trusted). Background MTA thread.
    public class TeamsUiaProbe {
        public volatile bool Enabled, Found, Muted;
        public volatile string MicButton = "";
        public long LastSeenTicks;
        readonly string logPath; readonly object logLock = new object();
        Thread thread; volatile bool stop; string lastLogged; DateTime lastDump = DateTime.MinValue;
        static readonly Regex LeaveRx = new Regex(@"^(Leave|Hang up|End meeting)(\b|$)", RegexOptions.IgnoreCase);
        static readonly Regex MicRx   = new Regex(@"^(Mute|Unmute)(?!\s+(all|everyone|notifications|chat|channel))|Ctrl\+Shift\+M", RegexOptions.IgnoreCase);
        public TeamsUiaProbe(string logPath) { this.logPath = logPath; }
        public void Start() { thread = new Thread(Loop); thread.IsBackground = true; thread.SetApartmentState(ApartmentState.MTA); thread.Start(); }
        public void Stop() { stop = true; }
        void Log(string msg) {
            try { lock (logLock) File.AppendAllText(logPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " [uia] " + msg + Environment.NewLine); } catch { }
        }
        void Loop() {
            var winCond = new PropertyCondition(AutomationElement.ClassNameProperty, "TeamsWebView");
            var btnCond = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
            while (!stop) {
                if (!Enabled) { Found = false; Thread.Sleep(500); continue; }
                try {
                    bool found = false, muted = false; string mic = "";
                    AutomationElementCollection wins = AutomationElement.RootElement.FindAll(TreeScope.Children, winCond);
                    foreach (AutomationElement w in wins) {
                        AutomationElementCollection btns = w.FindAll(TreeScope.Descendants, btnCond);
                        bool leave = false; string micName = null; var all = new StringBuilder();
                        foreach (AutomationElement b in btns) {
                            string n = b.Current.Name ?? "";
                            if (all.Length < 1500) all.Append(n).Append(" | ");
                            if (LeaveRx.IsMatch(n)) leave = true;
                            if (micName == null && MicRx.IsMatch(n)) micName = n;
                        }
                        if (leave && micName != null) { found = true; mic = micName; muted = Regex.IsMatch(micName, @"^Unmute", RegexOptions.IgnoreCase); break; }
                        if (leave && micName == null && (DateTime.UtcNow - lastDump).TotalSeconds > 60) {
                            lastDump = DateTime.UtcNow; Log("meeting window found but no mic button matched; buttons: " + all);
                        }
                    }
                    Found = found; Muted = muted; MicButton = mic;
                    if (found) LastSeenTicks = DateTime.UtcNow.Ticks;
                    string state = found ? ("mic button '" + mic + "' => muted=" + muted) : "no Teams meeting window with Leave + Mute/Unmute buttons";
                    if (state != lastLogged) { Log(state); lastLogged = state; }
                } catch (Exception ex) { Log("error: " + ex.Message); }
                Thread.Sleep(1500);
            }
        }
    }

    public static class Native {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string title);
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
        [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr ctx);
        [DllImport("shell32.dll")] public static extern int SHQueryUserNotificationState(out int state);
        [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        public const uint SWP_NOACTIVATE = 0x0010, SWP_SHOWWINDOW = 0x0040;
        public const int SW_HIDE = 0;

        // Primary Shell_TrayWnd plus one Shell_SecondaryTrayWnd per extra monitor that shows a taskbar.
        public static List<IntPtr> GetTaskbars() {
            var list = new List<IntPtr>();
            IntPtr h = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "Shell_TrayWnd", null);
            if (h != IntPtr.Zero) list.Add(h);
            IntPtr s = IntPtr.Zero;
            while ((s = FindWindowEx(IntPtr.Zero, s, "Shell_SecondaryTrayWnd", null)) != IntPtr.Zero) list.Add(s);
            return list;
        }
        // QUNS_BUSY=2, QUNS_RUNNING_D3D_FULL_SCREEN=3, QUNS_PRESENTATION_MODE=4 -> a full-screen app owns the screen.
        public static bool IsFullScreenAppActive() {
            int q; if (SHQueryUserNotificationState(out q) != 0) return false;
            return q == 2 || q == 3 || q == 4;
        }
    }

    // Who holds the microphone / webcam right now, straight from the registry (fast: ~1 ms for 130 keys).
    public static class CaptureProbe {
        const string Sub = @"Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\";
        public static List<string> GetUsers(string device, string[] ignore) {
            var users = new List<string>();
            foreach (RegistryKey hive in new RegistryKey[] { Registry.CurrentUser, Registry.LocalMachine }) {
                RegistryKey root = null;
                try { root = hive.OpenSubKey(Sub + device, false); } catch { }
                if (root == null) continue;
                using (root) {
                    Scan(root, ignore, users);
                    RegistryKey np = root.OpenSubKey("NonPackaged", false);
                    if (np != null) using (np) Scan(np, ignore, users);
                }
            }
            return users;
        }
        static long L(object v) { if (v is long) return (long)v; if (v is int) return (int)v; return 0; }
        static void Scan(RegistryKey parent, string[] ignore, List<string> users) {
            foreach (string name in parent.GetSubKeyNames()) {
                if (name == "NonPackaged") continue;
                using (RegistryKey k = parent.OpenSubKey(name, false)) {
                    if (k == null) continue;
                    long start = L(k.GetValue("LastUsedTimeStart")), stop = L(k.GetValue("LastUsedTimeStop"));
                    if (start == 0) continue;                 // never used
                    if (stop != 0 && stop >= start) continue; // released
                    bool skip = false;
                    if (ignore != null) foreach (string pat in ignore)
                        if (!string.IsNullOrEmpty(pat) && name.IndexOf(pat, StringComparison.OrdinalIgnoreCase) >= 0) { skip = true; break; }
                    if (!skip) users.Add(name);
                }
            }
        }
    }

    // Shared window style: TOPMOST | TRANSPARENT (click-through) | TOOLWINDOW (no Alt-Tab) | LAYERED | NOACTIVATE
    public class OverlayBase : Form {
        public OverlayBase() { FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; StartPosition = FormStartPosition.Manual; TopMost = true; }
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override CreateParams CreateParams {
            get { CreateParams cp = base.CreateParams; cp.ExStyle |= 0x8 | 0x20 | 0x80 | 0x80000 | 0x08000000; return cp; }
        }
    }

    // The glow: whole-window alpha (pulsed by the controller) over a vertical gradient.
    public class GlowForm : OverlayBase {
        private Color glow = Color.FromArgb(230, 30, 30);
        public Color Glow { get { return glow; } }
        public GlowForm() {
            Opacity = 0.0; BackColor = glow;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }
        public void SetGlowColor(Color c) { glow = c; BackColor = c; Invalidate(); }
        protected override void OnPaint(PaintEventArgs e) {
            Rectangle r = ClientRectangle;
            if (r.Width <= 0 || r.Height <= 0) return;
            using (var b = new LinearGradientBrush(r, ControlPaint.Light(glow, 0.5f), glow, LinearGradientMode.Vertical))
                e.Graphics.FillRectangle(b, r);
        }
    }

    // The label: its own per-pixel-alpha window (UpdateLayeredWindow), fully opaque text on a dark pill,
    // so it stays readable no matter how faint the glow is at that moment.
    public class LabelForm : OverlayBase {
        [DllImport("user32.dll")] static extern IntPtr GetDC(IntPtr h);
        [DllImport("user32.dll")] static extern int ReleaseDC(IntPtr h, IntPtr dc);
        [DllImport("gdi32.dll")] static extern IntPtr CreateCompatibleDC(IntPtr dc);
        [DllImport("gdi32.dll")] static extern bool DeleteDC(IntPtr dc);
        [DllImport("gdi32.dll")] static extern IntPtr SelectObject(IntPtr dc, IntPtr obj);
        [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr obj);
        [DllImport("user32.dll")] static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr dcDst, ref POINT dst, ref SIZE size, IntPtr dcSrc, ref POINT src, int key, ref BLENDFUNCTION blend, int flags);
        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; public POINT(int x, int y) { X = x; Y = y; } }
        [StructLayout(LayoutKind.Sequential)] struct SIZE { public int W, H; public SIZE(int w, int h) { W = w; H = h; } }
        [StructLayout(LayoutKind.Sequential)] struct BLENDFUNCTION { public byte Op, Flags, Alpha, Format; }
        public string Text2 = ""; public Color Fore = Color.Yellow; public int X = -1, Y = -1, H = -1; public bool V = false;

        // bar = the taskbar rectangle. A horizontal bar gets the pill at its left edge; a vertical bar
        // (docked left or right) gets the same pill rotated 90 degrees, running down the bar from the top.
        public void Render(string text, Color fore, Native.RECT bar) {
            int barW = bar.Right - bar.Left, barH = bar.Bottom - bar.Top;
            bool vertical = barH > barW;
            int thickness = vertical ? barW : barH;
            int height = vertical ? Math.Max(24, Math.Min(thickness - 8, 48)) : thickness;
            int x = vertical ? bar.Left + (thickness - height) / 2 : bar.Left + 14;
            int y = vertical ? bar.Top + 14 : bar.Top;
            if (text == Text2 && fore == Fore && x == X && y == Y && height == H && vertical == V) return;
            Text2 = text; Fore = fore; X = x; Y = y; H = height; V = vertical;
            float pt = Math.Max(9f, Math.Min(12.5f, height * 0.26f));
            using (Font font = new Font("Segoe UI", pt, FontStyle.Bold, GraphicsUnit.Point)) {
                SizeF measured;
                using (Bitmap probe = new Bitmap(1, 1)) using (Graphics g0 = Graphics.FromImage(probe)) {
                    g0.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit; measured = g0.MeasureString(text, font);
                }
                int padX = 12, padY = Math.Max(4, height / 8);
                int w = (int)Math.Ceiling(measured.Width) + 2 * padX, h = height;
                using (Bitmap bmp = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb)) {
                    using (Graphics g = Graphics.FromImage(bmp)) {
                        g.Clear(Color.Transparent);
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                        Rectangle pill = new Rectangle(0, padY, w - 1, h - 2 * padY - 1);
                        int rad = Math.Max(1, Math.Min(pill.Height, pill.Width) / 2);
                        using (GraphicsPath p = new GraphicsPath()) {
                            p.AddArc(pill.X, pill.Y, 2 * rad, 2 * rad, 180, 90);
                            p.AddArc(pill.Right - 2 * rad, pill.Y, 2 * rad, 2 * rad, 270, 90);
                            p.AddArc(pill.Right - 2 * rad, pill.Bottom - 2 * rad, 2 * rad, 2 * rad, 0, 90);
                            p.AddArc(pill.X, pill.Bottom - 2 * rad, 2 * rad, 2 * rad, 90, 90);
                            p.CloseFigure();
                            using (SolidBrush bg = new SolidBrush(Color.FromArgb(170, 0, 0, 0))) g.FillPath(bg, p);
                        }
                        using (SolidBrush fg = new SolidBrush(fore)) g.DrawString(text, font, fg, padX, (h - measured.Height) / 2f);
                    }
                    if (vertical) bmp.RotateFlip(RotateFlipType.Rotate90FlipNone);   // text reads top-to-bottom along the bar
                    Push(bmp, x, y);
                }
            }
        }
        void Push(Bitmap bmp, int x, int y) {
            IntPtr screenDc = GetDC(IntPtr.Zero), memDc = CreateCompatibleDC(screenDc), hBmp = IntPtr.Zero, old = IntPtr.Zero;
            try {
                hBmp = bmp.GetHbitmap(Color.FromArgb(0)); old = SelectObject(memDc, hBmp);
                SIZE size = new SIZE(bmp.Width, bmp.Height); POINT src = new POINT(0, 0), dst = new POINT(x, y);
                BLENDFUNCTION blend = new BLENDFUNCTION(); blend.Op = 0; blend.Flags = 0; blend.Alpha = 255; blend.Format = 1;   // AC_SRC_OVER, AC_SRC_ALPHA
                UpdateLayeredWindow(Handle, screenDc, ref dst, ref size, memDc, ref src, 0, ref blend, 2);                     // ULW_ALPHA
            } finally {
                if (old != IntPtr.Zero) SelectObject(memDc, old);
                if (hBmp != IntPtr.Zero) DeleteObject(hBmp);
                DeleteDC(memDc); ReleaseDC(IntPtr.Zero, screenDc);
            }
        }
    }

    public enum GlowMode { Off, Pulse, Steady }

    // Owns one glow + one label per taskbar, pulses the glow at 20 fps and keeps both glued on top.
    public class GlowController {
        readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        readonly Dictionary<IntPtr, GlowForm> forms = new Dictionary<IntPtr, GlowForm>();
        readonly Dictionary<IntPtr, LabelForm> labels = new Dictionary<IntPtr, LabelForm>();
        readonly Dictionary<IntPtr, Native.RECT> rects = new Dictionary<IntPtr, Native.RECT>();
        GlowMode mode = GlowMode.Off; string label = ""; Color color = Color.Red, labelColor = Color.Yellow;
        readonly double minOp, maxOp, period, steadyOp; DateTime t0; int ticks;
        const uint SWP_NOSIZE = 0x0001, SWP_NOMOVE = 0x0002;
        public GlowController(double minOpacity, double maxOpacity, double periodSeconds, double steadyOpacity) {
            minOp = minOpacity; maxOp = maxOpacity; period = periodSeconds; steadyOp = Math.Min(steadyOpacity, 0.98);
            RefreshTaskbars();
            timer.Interval = 50; timer.Tick += Tick; timer.Start();
        }
        public void Apply(GlowMode m, string lbl, Color c, Color lblColor) {
            if (m != GlowMode.Off && mode == GlowMode.Off) t0 = DateTime.UtcNow;   // fade in from 0 on every new alert
            mode = m; label = lbl ?? ""; color = c; labelColor = lblColor;
            if (mode == GlowMode.Off) HideAll();
        }
        public void HideAll() {
            foreach (GlowForm f in forms.Values) Native.ShowWindow(f.Handle, Native.SW_HIDE);
            foreach (LabelForm l in labels.Values) Native.ShowWindow(l.Handle, Native.SW_HIDE);
        }
        public void Dispose() {
            timer.Stop(); HideAll();
            foreach (GlowForm f in forms.Values) f.Dispose(); forms.Clear();
            foreach (LabelForm l in labels.Values) l.Dispose(); labels.Clear();
        }
        void RefreshTaskbars() {
            rects.Clear();
            foreach (IntPtr tb in Native.GetTaskbars()) {
                if (!Native.IsWindowVisible(tb)) continue;
                Native.RECT r; if (Native.GetWindowRect(tb, out r)) rects[tb] = r;
            }
        }
        void Tick(object sender, EventArgs e) {
            if (++ticks % 10 == 0) RefreshTaskbars();            // every 500 ms: taskbar moved / auto-hid / monitor change
            if (mode == GlowMode.Off) return;
            double op = steadyOp;
            if (mode == GlowMode.Pulse) {
                double t = (DateTime.UtcNow - t0).TotalSeconds;
                op = minOp + (maxOp - minOp) * (0.5 - 0.5 * Math.Cos(2 * Math.PI * t / period));
            }
            foreach (KeyValuePair<IntPtr, Native.RECT> kv in rects) {
                Native.RECT r = kv.Value;
                GlowForm f;
                if (!forms.TryGetValue(kv.Key, out f)) { f = new GlowForm(); IntPtr h = f.Handle; forms[kv.Key] = f; }
                if (f.Glow != color) f.SetGlowColor(color);
                f.Opacity = op;
                Native.SetWindowPos(f.Handle, Native.HWND_TOPMOST, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, Native.SWP_NOACTIVATE | Native.SWP_SHOWWINDOW);

                LabelForm l;
                if (!labels.TryGetValue(kv.Key, out l)) { l = new LabelForm(); IntPtr h2 = l.Handle; labels[kv.Key] = l; }
                if (label.Length == 0) { Native.ShowWindow(l.Handle, Native.SW_HIDE); continue; }
                l.Render(label, labelColor, r);
                // placed after the glow so it lands above it in the topmost band
                Native.SetWindowPos(l.Handle, Native.HWND_TOPMOST, 0, 0, 0, 0, Native.SWP_NOACTIVATE | Native.SWP_SHOWWINDOW | SWP_NOSIZE | SWP_NOMOVE);
            }
        }
    }

    // Teams third-party app API client (ws://localhost:8124). Background thread, auto-reconnect, token persistence.
    public class TeamsLink {
        public volatile bool Connected, Paired, InMeeting, Muted, VideoOn, CanToggleMute, CanPair;
        public volatile string Status = "not started";
        public volatile string Token;
        public long LastUpdateTicks;
        readonly string tokenPath, logPath; readonly int port; readonly object logLock = new object();
        readonly JavaScriptSerializer js = new JavaScriptSerializer();
        Thread thread; volatile bool stop;
        public TeamsLink(int port, string tokenPath, string logPath) { this.port = port; this.tokenPath = tokenPath; this.logPath = logPath; LoadToken(); }
        public void Start() { thread = new Thread(Loop); thread.IsBackground = true; thread.Start(); }
        public void Stop() { stop = true; }
        public bool HasToken { get { return !string.IsNullOrEmpty(Token); } }
        void LoadToken() { try { Token = File.Exists(tokenPath) ? File.ReadAllText(tokenPath).Trim() : null; } catch { } }
        public void Log(string msg) {
            try { lock (logLock) File.AppendAllText(logPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " [teams] " + msg + Environment.NewLine); } catch { }
        }
        static Exception Root(Exception ex) { while (ex.InnerException != null) ex = ex.InnerException; return ex; }
        void SleepUnlessStopped(int ms) { for (int i = 0; i < ms && !stop; i += 200) Thread.Sleep(200); }
        static bool B(Dictionary<string, object> d, string k) { object v; return d.TryGetValue(k, out v) && v is bool && (bool)v; }

        void Loop() {
            string lastOffline = null;
            while (!stop) {
                LoadToken();
                string url = "ws://localhost:" + port + "?" + (HasToken ? "token=" + Uri.EscapeDataString(Token) + "&" : "")
                           + "protocol-version=2.0.0&manufacturer=MicCamGlow&device=MicCamGlow&app=MicCamGlow&app-version=1.0";
                using (ClientWebSocket ws = new ClientWebSocket()) {
                    try { ws.ConnectAsync(new Uri(url), CancellationToken.None).Wait(); }
                    catch (Exception ex) {
                        string m = "offline (" + Root(ex).Message.Trim() + ")";
                        Status = m; Connected = false; Paired = false;
                        if (m != lastOffline) { Log(m); lastOffline = m; }   // log each distinct reason once, not every retry
                        SleepUnlessStopped(10000); continue;
                    }
                    lastOffline = null;
                    Connected = true; Paired = false; CanToggleMute = false; CanPair = false;
                    Status = HasToken ? "connected (token), waiting for state" : "connected (no token), waiting for Allow in Teams";
                    Log("connected " + (HasToken ? "with token" : "without token"));
                    try { Send(ws, "{\"action\":\"query-state\",\"parameters\":{},\"requestId\":1}"); } catch { }
                    byte[] buf = new byte[65536]; StringBuilder sb = new StringBuilder();
                    try {
                        while (!stop && ws.State == WebSocketState.Open) {
                            WebSocketReceiveResult res = ws.ReceiveAsync(new ArraySegment<byte>(buf), CancellationToken.None).Result;
                            if (res.MessageType == WebSocketMessageType.Close) break;
                            sb.Append(Encoding.UTF8.GetString(buf, 0, res.Count));
                            if (!res.EndOfMessage) continue;
                            string msg = sb.ToString(); sb.Length = 0;
                            Handle(msg);
                        }
                    } catch (Exception ex) { Log("receive error: " + Root(ex).Message.Trim()); }
                    Connected = false; Paired = false; CanToggleMute = false; Status = "disconnected";
                    Log("disconnected");
                }
                SleepUnlessStopped(3000);
            }
        }
        void Send(ClientWebSocket ws, string json) {
            byte[] b = Encoding.UTF8.GetBytes(json);
            ws.SendAsync(new ArraySegment<byte>(b), WebSocketMessageType.Text, true, CancellationToken.None).Wait();
        }
        void Handle(string msg) {
            try {
                Dictionary<string, object> d = js.DeserializeObject(msg) as Dictionary<string, object>;
                if (d == null) return;
                object tr;
                if (d.TryGetValue("tokenRefresh", out tr) && tr is string && !string.IsNullOrEmpty((string)tr)) {
                    Token = (string)tr;
                    try { Directory.CreateDirectory(Path.GetDirectoryName(tokenPath)); File.WriteAllText(tokenPath, Token); } catch { }
                    Log("token received from Teams and saved");
                    return;
                }
                object mu;
                if (d.TryGetValue("meetingUpdate", out mu) && mu is Dictionary<string, object>) {
                    Dictionary<string, object> m = (Dictionary<string, object>)mu; object st, pm;
                    if (m.TryGetValue("meetingPermissions", out pm) && pm is Dictionary<string, object>) {
                        Dictionary<string, object> p = (Dictionary<string, object>)pm;
                        CanToggleMute = B(p, "canToggleMute"); CanPair = B(p, "canPair");
                    }
                    if (m.TryGetValue("meetingState", out st) && st is Dictionary<string, object>) {
                        Dictionary<string, object> s = (Dictionary<string, object>)st;
                        InMeeting = B(s, "isInMeeting"); Muted = B(s, "isMuted"); VideoOn = B(s, "isVideoOn") || B(s, "isCameraOn");
                        Paired = true; LastUpdateTicks = DateTime.UtcNow.Ticks; Status = "live";
                    }
                    Log("state inMeeting=" + InMeeting + " muted=" + Muted + " video=" + VideoOn + " canToggleMute=" + CanToggleMute + " canPair=" + CanPair);
                    return;
                }
                object err;
                if (d.TryGetValue("errorMsg", out err) && err != null) { Status = "error: " + err; Log("error from Teams: " + err); }
            } catch (Exception ex) { Log("parse error: " + ex.Message + " msg=" + msg); }
        }
    }
}
'@

# ---------------------------------------------------------------------------
# Management switches
# ---------------------------------------------------------------------------
if ($Status) {
    $procs = @(Get-WatcherProcesses)
    if ($procs.Count) { "Watcher:          running, PID " + (($procs | ForEach-Object ProcessId) -join ', ') } else { 'Watcher:          NOT running' }
    "Startup shortcut: " + $(if (Test-Path $ShortcutPath) { 'installed' } else { 'not installed' })
    $listen = Get-NetTCPConnection -LocalPort $TeamsPort -State Listen -ErrorAction SilentlyContinue
    "Teams mute source: UI Automation of the meeting window (Teams' websocket API was retired by Microsoft, Dec 2025)"
    if ($listen) { "Teams API port:   listening on $TeamsPort (pid $($listen[0].OwningProcess)); start with -TeamsApi to use it" }
    if (Test-Path $TokenPath) { "Teams token:      saved at $TokenPath" }
    $mic = [MicCamGlow.CaptureProbe]::GetUsers('microphone', [string[]]$Ignore)
    $cam = [MicCamGlow.CaptureProbe]::GetUsers('webcam', [string[]]$Ignore)
    "Mic held by:      " + $(if ($mic.Count) { ($mic | ForEach-Object { ConvertTo-FriendlyAppName $_ }) -join ', ' } else { '(nobody)' })
    "Camera held by:   " + $(if ($cam.Count) { ($cam | ForEach-Object { ConvertTo-FriendlyAppName $_ }) -join ', ' } else { '(nobody)' })
    if (Test-Path $LogPath) { "Last log lines:"; Get-Content $LogPath -Tail 6 | ForEach-Object { "  $_" } }
    return
}
if ($Stop -or $Uninstall) {
    Stop-Watcher
    if ($Uninstall) { if (Test-Path $ShortcutPath) { Remove-Item $ShortcutPath -Force; "Removed $ShortcutPath" } else { 'No Startup shortcut to remove.' } }
    return
}
if ($TeamsToken) {
    Set-Content -LiteralPath $TokenPath -Value $TeamsToken.Trim() -NoNewline
    "Teams API token saved to $TokenPath"
    if (@(Get-WatcherProcesses).Count) { Stop-Watcher; Start-Watcher }
    return
}
if ($Install) {
    $wsh = New-Object -ComObject WScript.Shell
    $lnk = $wsh.CreateShortcut($ShortcutPath)
    $lnk.TargetPath = $ConhostExe
    $lnk.Arguments = Get-LaunchArguments
    $lnk.WorkingDirectory = Split-Path $ScriptPath
    $lnk.WindowStyle = 7
    $lnk.IconLocation = "$env:SystemRoot\System32\SndVol.exe,0"
    $lnk.Description = 'Red taskbar glow while mic or camera is live'
    $lnk.Save()
    "Startup shortcut written: $ShortcutPath"
    Start-Watcher
    return
}

# ---------------------------------------------------------------------------
# Watcher
# ---------------------------------------------------------------------------
$created = $false
$mutex = New-Object System.Threading.Mutex($true, 'Local\MicCamGlow', [ref]$created)
if (-not $created) { Write-Warning 'MicCamGlow is already running in this session.'; return }

try { if ((Get-Item $LogPath -ErrorAction SilentlyContinue).Length -gt 1MB) { Remove-Item $LogPath -Force } } catch { }
try { [void][MicCamGlow.Native]::SetProcessDpiAwarenessContext([IntPtr]-4) } catch { }   # per-monitor DPI; harmless if it fails

$redColor   = [System.Drawing.ColorTranslator]::FromHtml($Color)
$greenColor = [System.Drawing.ColorTranslator]::FromHtml($MutedColor)
$textColor  = [System.Drawing.ColorTranslator]::FromHtml($LabelColor)
$ctl   = New-Object MicCamGlow.GlowController($MinOpacity, $MaxOpacity, $PulseSeconds, $MutedOpacity)
$teams = $null
if ($TeamsApi) { $teams = New-Object MicCamGlow.TeamsLink($TeamsPort, $TokenPath, $LogPath); $teams.Start() }
$uia = $null
if (-not $NoUia) { $uia = New-Object MicCamGlow.TeamsUiaProbe($LogPath); $uia.Start() }

$script:demoEnd   = if ($DemoSeconds -gt 0) { [DateTime]::UtcNow.AddSeconds($DemoSeconds) } else { [DateTime]::MinValue }
$script:lastState = ''
Write-Log "started pid $PID (uia: $(if ($uia) { 'on' } else { 'off' }), teams api: $(if ($teams) { 'on' } else { 'off' }), poll $PollMs ms)"

$poll = New-Object System.Windows.Forms.Timer
$poll.Interval = $PollMs
$poll.Add_Tick({
    try {
        $mic = [MicCamGlow.CaptureProbe]::GetUsers('microphone', [string[]]$Ignore)
        $cam = [MicCamGlow.CaptureProbe]::GetUsers('webcam', [string[]]$Ignore)

        # Teams API is trusted only when it affirmatively says: paired, in a meeting, and allowed to toggle mute.
        $tl = $teams
        $apiKnown = ($null -ne $tl) -and $tl.Paired -and $tl.CanToggleMute -and $tl.InMeeting
        # UI Automation fallback: Mute/Unmute button of the meeting window, seen within the last 5 s.
        $teamsHoldsMic = [bool]($mic -like 'MSTeams_*')
        if ($uia) { $uia.Enabled = $teamsHoldsMic -and -not $apiKnown }
        $uiaKnown = ($null -ne $uia) -and $uia.Found -and (([DateTime]::UtcNow.Ticks - $uia.LastSeenTicks) -lt 50000000)
        $teamsKnown = $apiKnown -or $uiaKnown
        $teamsMuted = ($apiKnown -and $tl.Muted) -or ((-not $apiKnown) -and $uiaKnown -and $uia.Muted)
        $teamsLive  = $teamsKnown -and -not $teamsMuted

        $micList = @($mic | Where-Object { -not ($teamsMuted -and $_ -like 'MSTeams_*') })
        $camList = @($cam)
        if ($tl -and $tl.Paired -and $tl.VideoOn -and -not ($camList -like 'MSTeams_*')) { $camList += 'MSTeams_api' }

        $demo = [DateTime]::UtcNow -lt $script:demoEnd
        $fullscreen = [MicCamGlow.Native]::IsFullScreenAppActive()
        $red   = (-not $fullscreen) -and (($micList.Count -gt 0) -or ($camList.Count -gt 0) -or $demo)
        $green = (-not $fullscreen) -and (-not $red) -and $teamsMuted -and (-not $NoMutedIndicator)

        $label = ''
        if ($red) {
            $parts = @()
            if ($micList.Count) {
                $names = @($micList | ForEach-Object { ConvertTo-FriendlyAppName $_ } | Sort-Object -Unique)
                $teamsHolds = [bool]($micList -like 'MSTeams_*')
                $word = if ($teamsHolds -and $teamsLive) { 'MIC LIVE' } elseif ($teamsHolds -and -not $teamsKnown) { 'MIC OPEN' } else { 'MIC ON' }
                $hint = if ($teamsHolds -and -not $teamsKnown) { ', mute unknown' } else { '' }
                $parts += "$word  ($($names -join ', ')$hint)"
            }
            if ($camList.Count) {
                $names = @($camList | ForEach-Object { ConvertTo-FriendlyAppName $_ } | Sort-Object -Unique)
                $parts += "CAMERA ON  ($($names -join ', '))"
            }
            if (-not $parts) { $parts += 'DEMO' }
            $label = $parts -join '   +   '
        } elseif ($green) { $label = 'MUTED  (Teams)' }
        if ($NoLabel) { $label = '' }

        if ($red)        { $ctl.Apply([MicCamGlow.GlowMode]::Pulse,  $label, $redColor, $textColor) }
        elseif ($green)  { $ctl.Apply([MicCamGlow.GlowMode]::Steady, $label, $greenColor, $textColor) }
        else             { $ctl.Apply([MicCamGlow.GlowMode]::Off, '', $redColor, $textColor) }

        $state = "$red|$green|$label|$fullscreen"
        if ($state -ne $script:lastState) {
            Write-Log ("{0} label='{1}' fullscreen={2} mic=[{3}] cam=[{4}] api={5} uia={6}" -f $(if ($red) { 'RED' } elseif ($green) { 'GREEN' } else { 'off' }), $label, $fullscreen, ($mic -join ';'), ($cam -join ';'), $(if ($tl) { $tl.Status } else { 'disabled' }), $(if (-not $uia) { 'disabled' } elseif ($uiaKnown) { "muted=$($uia.Muted)" } else { 'n/a' }))
            $script:lastState = $state
        }
    } catch { Write-Log "poll error: $_" }
})

$poll.Start()
Write-Host "MicCamGlow watching (poll $PollMs ms, UIA $(if ($uia) { 'on' } else { 'off' }), Teams API $(if ($teams) { 'on' } else { 'off' })). Ctrl+C to stop. Log: $LogPath"
try { [System.Windows.Forms.Application]::Run() }
finally {
    $poll.Stop(); $ctl.Dispose(); if ($teams) { $teams.Stop() }; if ($uia) { $uia.Stop() }
    Write-Log 'stopped'
    $mutex.ReleaseMutex() | Out-Null; $mutex.Dispose()
}
