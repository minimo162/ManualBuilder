using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;

namespace ManualBuilder.RecorderCompanion
{
    public static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            return Run(args);
        }

        public static int Run(string[] args)
        {
            try
            {
                try { SetProcessDpiAwarenessContext(new IntPtr(-4)); } catch { }
                try { SetThreadDpiAwarenessContext(new IntPtr(-4)); } catch { }
                CompanionOptions options = CompanionOptions.Parse(args);
                Application app = new Application();
                app.ShutdownMode = ShutdownMode.OnMainWindowClose;
                CompanionWindow window = new CompanionWindow(options);
                app.Run(window);
                return 0;
            }
            catch (Exception ex)
            {
                try
                {
                    string path = Path.Combine(Path.GetTempPath(), "ManualBuilder-RecorderCompanion-startup.log");
                    File.AppendAllText(path, ex + Environment.NewLine, new UTF8Encoding(false));
                }
                catch { }
                return 1;
            }
        }

        [DllImport("user32.dll")]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr value);

        [DllImport("user32.dll")]
        private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr value);
    }

    internal sealed class CompanionOptions
    {
        public string StatusPath;
        public string EventsDirectory;
        public string PausePath;
        public string UndoPath;
        public string ManualResultPath;
        public string StopPath;
        public string JobId;
        public string WebRoot;
        public bool TestMode;

        public static CompanionOptions Parse(string[] args)
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < args.Length; i++)
            {
                string key = args[i];
                if (!key.StartsWith("--", StringComparison.Ordinal) || i + 1 >= args.Length)
                    throw new ArgumentException("記録モニターの起動引数が正しくありません。");
                values[key.Substring(2)] = args[++i];
            }
            CompanionOptions options = new CompanionOptions();
            options.StatusPath = RequiredPath(values, "status");
            options.EventsDirectory = RequiredPath(values, "events");
            options.PausePath = RequiredPath(values, "pause");
            options.UndoPath = RequiredPath(values, "undo");
            options.ManualResultPath = RequiredPath(values, "result");
            options.StopPath = RequiredPath(values, "stop");
            options.JobId = RequiredValue(values, "job");
            options.WebRoot = RequiredPath(values, "web-root");
            string testMode;
            options.TestMode = values.TryGetValue("test-mode", out testMode) &&
                String.Equals(testMode, "true", StringComparison.OrdinalIgnoreCase);
            if (!Directory.Exists(options.WebRoot)) throw new DirectoryNotFoundException(options.WebRoot);
            return options;
        }

        private static string RequiredValue(Dictionary<string, string> values, string name)
        {
            string value;
            if (!values.TryGetValue(name, out value) || String.IsNullOrWhiteSpace(value))
                throw new ArgumentException("--" + name + " が必要です。");
            return value;
        }

        private static string RequiredPath(Dictionary<string, string> values, string name)
        {
            return Path.GetFullPath(RequiredValue(values, name));
        }
    }

    internal sealed class CompanionWindow : Window
    {
        private const int GaRoot = 2;
        private const int WmNcLButtonDown = 0x00A1;
        private const int HtCaption = 2;
        private const uint MonitorDefaultToNearest = 2;
        private const uint WdaExcludeFromCapture = 0x00000011;
        private const int DwmwaWindowCornerPreference = 33;
        private const int DwmwcpRound = 2;

        private readonly CompanionOptions options;
        private readonly WebView2 webView;
        private readonly DispatcherTimer timer;
        private readonly JavaScriptSerializer json;
        private IntPtr windowHandle;
        private IntPtr lastExternalWindow;
        private bool webReady;
        private bool compact = true;
        private bool closingRequested;
        private bool closingFromStatus;
        private string lastKnownState = "starting";
        private int currentCount;
        private string pendingUndoId = String.Empty;
        private DateTime pendingUndoAtUtc;
        private DateTime undoNoticeUntilUtc;
        private string pendingResultId = String.Empty;
        private string pendingResultPayload = String.Empty;
        private DateTime pendingResultAtUtc;
        private DateTime resultNoticeUntilUtc;
        private bool resultRetryReady;
        private DateTime closeAtUtc;
        private string helpText = "対象アプリでいつもどおり操作してください。クリックや入力は自動で記録されます。";
        private string tracedState = String.Empty;

        public CompanionWindow(CompanionOptions options)
        {
            this.options = options;
            json = new JavaScriptSerializer();
            Title = "ManualBuilder Recorder";
            Width = 760;
            Height = 300;
            MinWidth = 720;
            MinHeight = 270;
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            Topmost = true;
            ShowInTaskbar = true;
            ShowActivated = options.TestMode;
            Background = Brushes.White;

            lastExternalWindow = GetForegroundWindow();
            webView = new WebView2();
            Content = webView;
            Loaded += OnLoaded;
            SourceInitialized += OnSourceInitialized;
            Closing += OnClosing;
            Closed += OnClosed;

            // WebView2の連続描画でBackground優先度が飢餓状態にならないよう、状態監視はNormalで実行する。
            timer = new DispatcherTimer(DispatcherPriority.Normal);
            timer.Interval = TimeSpan.FromMilliseconds(150);
            timer.Tick += OnTimerTick;
        }

        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            PositionAtWorkAreaEdge();
            try
            {
                string userData = options.TestMode
                    ? Path.Combine(Path.GetDirectoryName(options.WebRoot), ".webview-test-" + Guid.NewGuid().ToString("N"))
                    : Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "ManualBuilder", "WebView2", "RecorderCompanion");
                Directory.CreateDirectory(userData);
                CoreWebView2Environment environment = await CoreWebView2Environment.CreateAsync(null, userData);
                await webView.EnsureCoreWebView2Async(environment);
                ConfigureWebView();
                timer.Start();
            }
            catch (Exception ex)
            {
                Log(ex);
                FailStartup("記録モニターを開けませんでした。WebView2 Runtimeを確認してください。");
            }
        }

        private void ConfigureWebView()
        {
            CoreWebView2 core = webView.CoreWebView2;
            string jobDirectory = Path.GetDirectoryName(options.StatusPath);
            core.SetVirtualHostNameToFolderMapping(
                "recording.manualbuilder.local", jobDirectory,
                CoreWebView2HostResourceAccessKind.Allow);
            core.Settings.AreDevToolsEnabled = false;
            core.Settings.AreDefaultContextMenusEnabled = false;
            core.Settings.AreBrowserAcceleratorKeysEnabled = false;
            core.Settings.IsStatusBarEnabled = false;
            core.Settings.IsZoomControlEnabled = false;
            core.Settings.AreDefaultScriptDialogsEnabled = false;
            string pageUri = new Uri(Path.Combine(options.WebRoot, "index.html")).AbsoluteUri;
            core.NavigationStarting += delegate(object sender, CoreWebView2NavigationStartingEventArgs e)
            {
                TraceTest("navigation-start=" + e.Uri);
                if (String.Equals(e.Uri, "about:blank", StringComparison.OrdinalIgnoreCase)) return;
                if (!String.Equals(e.Uri, pageUri, StringComparison.OrdinalIgnoreCase)) e.Cancel = true;
            };
            core.NewWindowRequested += delegate(object sender, CoreWebView2NewWindowRequestedEventArgs e) { e.Handled = true; };
            core.PermissionRequested += delegate(object sender, CoreWebView2PermissionRequestedEventArgs e) { e.State = CoreWebView2PermissionState.Deny; };
            core.WebMessageReceived += OnWebMessageReceived;
            core.NavigationCompleted += delegate(object sender, CoreWebView2NavigationCompletedEventArgs e)
            {
                if (!e.IsSuccess) TraceTest("navigation-error=" + e.WebErrorStatus);
            };
            core.Navigate(pageUri);
        }

        private void OnSourceInitialized(object sender, EventArgs e)
        {
            windowHandle = new WindowInteropHelper(this).Handle;
            try
            {
                int awareness = GetAwarenessFromDpiAwarenessContext(GetThreadDpiAwarenessContext());
                WriteText(options.StatusPath + ".companion.start.log", options.JobId + Environment.NewLine + "dpiAwareness=" + awareness);
            }
            catch { }
            if (!options.TestMode)
            {
                try { SetWindowDisplayAffinity(windowHandle, WdaExcludeFromCapture); } catch { }
            }
            try
            {
                int preference = DwmwcpRound;
                DwmSetWindowAttribute(windowHandle, DwmwaWindowCornerPreference, ref preference, Marshal.SizeOf(typeof(int)));
            }
            catch { }
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            string pageUri = new Uri(Path.Combine(options.WebRoot, "index.html")).AbsoluteUri;
            if (!String.Equals(e.Source, pageUri, StringComparison.OrdinalIgnoreCase)) return;
            string raw;
            try { raw = e.TryGetWebMessageAsString(); }
            catch { return; }
            Dictionary<string, object> message;
            try { message = json.Deserialize<Dictionary<string, object>>(raw); }
            catch { return; }
            string type = StringValue(message, "type");
            TraceTest("message=" + type + " command=" + StringValue(message, "command"));
            if (type == "client-error")
            {
                TraceTest("client-error=" + StringValue(message, "message") + " line=" + StringValue(message, "line"));
                return;
            }
            if (type == "ready")
            {
                webReady = true;
                TryWrite(options.StatusPath + ".companion.ready", options.JobId);
                PublishState();
                return;
            }
            if (type == "drag") { BeginWindowDrag(); return; }
            if (type == "minimize") { WindowState = WindowState.Minimized; return; }
            if (type == "close") { RequestClose(); return; }
            if (type == "cancel-close") { Publish(new Dictionary<string, object> { { "type", "hide-close-confirm" } }); return; }
            if (type == "toggle-compact") { SetCompact(!compact); RestoreTargetFocus(); return; }
            if (type != "command") return;
            string command = StringValue(message, "command");
            if (command == "pause") TogglePause();
            else if (command == "undo") RequestUndo();
            else if (command == "result") RequestResult();
            else if (command == "finish") RequestFinish();
        }

        private void OnTimerTick(object sender, EventArgs e)
        {
            try
            {
                TrackExternalWindow();
                Dictionary<string, object> status = ReadStatus();
                if (status == null || StringValue(status, "jobId") != options.JobId) return;
                lastKnownState = StringValue(status, "state");
                currentCount = IntValue(status, "count");
                if (tracedState != lastKnownState)
                {
                    tracedState = lastKnownState;
                    TraceTest("state=" + lastKnownState + " count=" + currentCount);
                }
                ProcessAcknowledgements(status);
                PublishState(status);

                if (lastKnownState == "failed")
                {
                    closingRequested = false;
                    if (!String.IsNullOrWhiteSpace(StringValue(status, "message"))) helpText = StringValue(status, "message");
                    if (compact) SetCompact(false);
                    PublishState(status);
                }
                else if (lastKnownState == "completed" || lastKnownState == "idle")
                {
                    closingFromStatus = true;
                    if (closeAtUtc == DateTime.MinValue) closeAtUtc = DateTime.UtcNow.AddMilliseconds(2600);
                }
                if (closeAtUtc != DateTime.MinValue && DateTime.UtcNow >= closeAtUtc)
                {
                    timer.Stop();
                    closingFromStatus = true;
                    try { webView.Dispose(); } catch { }
                    Application.Current.Shutdown();
                }
            }
            catch (Exception ex) { Log(ex); }
        }

        private void ProcessAcknowledgements(Dictionary<string, object> status)
        {
            string undoRequestId = StringValue(status, "undoRequestId");
            if (!String.IsNullOrEmpty(pendingUndoId) && undoRequestId == pendingUndoId)
            {
                pendingUndoId = String.Empty;
                pendingUndoAtUtc = DateTime.MinValue;
                undoNoticeUntilUtc = DateTime.MinValue;
                helpText = DefaultHelp();
            }
            else if (!String.IsNullOrEmpty(pendingUndoId) && DateTime.UtcNow >= pendingUndoAtUtc.AddSeconds(6))
            {
                pendingUndoAtUtc = DateTime.UtcNow;
                undoNoticeUntilUtc = DateTime.UtcNow.AddSeconds(6);
                TryWrite(options.UndoPath, pendingUndoId);
                helpText = "削除結果を確認しています。対象アプリで行った操作には影響しません。";
            }

            string resultRequestId = StringValue(status, "resultRequestId");
            if (!String.IsNullOrEmpty(pendingResultId) && resultRequestId == pendingResultId)
            {
                pendingResultId = String.Empty;
                pendingResultPayload = String.Empty;
                pendingResultAtUtc = DateTime.MinValue;
                resultRetryReady = false;
                bool failed = StringValue(status, "message").Contains("できませんでした");
                resultNoticeUntilUtc = failed ? DateTime.UtcNow.AddSeconds(6) : DateTime.MinValue;
                helpText = failed
                    ? "結果画面は追加されていません。対象アプリの画面を確認して、もう一度お試しください。"
                    : DefaultHelp();
                if (!IsVisible) Show();
                RestoreTargetFocus();
            }
            else if (!String.IsNullOrEmpty(pendingResultId) && DateTime.UtcNow >= pendingResultAtUtc.AddSeconds(6))
            {
                pendingResultAtUtc = DateTime.UtcNow;
                resultNoticeUntilUtc = DateTime.UtcNow.AddSeconds(6);
                resultRetryReady = true;
                helpText = "追加結果を受け取れませんでした。［追加を再確認］を押すか、そのまま終了できます。";
                if (!IsVisible) Show();
                RestoreTargetFocus();
            }
            if (resultNoticeUntilUtc != DateTime.MinValue && DateTime.UtcNow >= resultNoticeUntilUtc && String.IsNullOrEmpty(pendingResultId))
                helpText = DefaultHelp();
            if (undoNoticeUntilUtc != DateTime.MinValue && DateTime.UtcNow >= undoNoticeUntilUtc && String.IsNullOrEmpty(pendingUndoId))
                helpText = DefaultHelp();
        }

        private void TogglePause()
        {
            if (closingRequested || !String.IsNullOrEmpty(pendingResultId)) return;
            try
            {
                if (File.Exists(options.PausePath)) File.Delete(options.PausePath);
                else WriteText(options.PausePath, "pause");
            }
            catch (Exception ex) { Log(ex); }
            RestoreTargetFocus();
        }

        private void RequestUndo()
        {
            TraceTest("undo requested closing=" + closingRequested + " count=" + currentCount + " pendingUndo=" + pendingUndoId + " pendingResult=" + pendingResultId);
            if (closingRequested || currentCount <= 0 || !String.IsNullOrEmpty(pendingUndoId) || !String.IsNullOrEmpty(pendingResultId)) return;
            pendingUndoId = Guid.NewGuid().ToString("N");
            pendingUndoAtUtc = DateTime.UtcNow;
            if (!TryWrite(options.UndoPath, pendingUndoId))
            {
                pendingUndoId = String.Empty;
                pendingUndoAtUtc = DateTime.MinValue;
            }
            TraceTest("undo written id=" + pendingUndoId);
            PublishState();
            RestoreTargetFocus();
        }

        private async void RequestResult()
        {
            if (closingRequested || lastExternalWindow == IntPtr.Zero || !String.IsNullOrEmpty(pendingUndoId)) return;
            if (!String.IsNullOrEmpty(pendingResultId))
            {
                if (!resultRetryReady || String.IsNullOrEmpty(pendingResultPayload)) return;
                resultRetryReady = false;
            }
            else
            {
                pendingResultId = Guid.NewGuid().ToString("N");
                Dictionary<string, object> request = new Dictionary<string, object>();
                request["requestId"] = pendingResultId;
                request["windowHandle"] = lastExternalWindow.ToInt64();
                pendingResultPayload = json.Serialize(request);
            }
            pendingResultAtUtc = DateTime.UtcNow;
            PublishState();
            Hide();
            SetForegroundWindow(lastExternalWindow);
            await Task.Delay(180);
            if (!TryWrite(options.ManualResultPath, pendingResultPayload))
            {
                resultRetryReady = true;
                Show();
            }
        }

        private void RequestFinish()
        {
            if (lastKnownState == "failed")
            {
                closingFromStatus = true;
                Close();
                return;
            }
            if (closingRequested) return;
            closingRequested = true;
            if (!TryWrite(options.StopPath, "stop")) closingRequested = false;
            PublishState();
        }

        private void RequestClose()
        {
            if (closingFromStatus || lastKnownState == "failed")
            {
                closingFromStatus = true;
                Close();
                return;
            }
            if (closingRequested) return;
            Publish(new Dictionary<string, object> { { "type", "show-close-confirm" } });
        }

        private void OnClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (closingFromStatus) return;
            e.Cancel = true;
            RequestClose();
        }

        private void OnClosed(object sender, EventArgs e)
        {
            timer.Stop();
            try { webView.Dispose(); } catch { }
        }

        private void SetCompact(bool value)
        {
            compact = value;
            if (compact)
            {
                MinWidth = 720;
                MinHeight = 270;
                Width = 760;
                Height = 300;
                ResizeMode = ResizeMode.NoResize;
            }
            else
            {
                ResizeMode = ResizeMode.CanResize;
                MinWidth = 760;
                MinHeight = 560;
                Width = 980;
                Height = 760;
            }
            KeepInsideWorkArea();
            PublishState();
        }

        private void BeginWindowDrag()
        {
            if (windowHandle == IntPtr.Zero) return;
            ReleaseCapture();
            SendMessage(windowHandle, WmNcLButtonDown, new IntPtr(HtCaption), IntPtr.Zero);
        }

        private void TrackExternalWindow()
        {
            IntPtr foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero) return;
            IntPtr root = GetAncestor(foreground, GaRoot);
            if (root == IntPtr.Zero) root = foreground;
            uint processId;
            GetWindowThreadProcessId(root, out processId);
            if (processId == (uint)Process.GetCurrentProcess().Id) return;
            string title = WindowTitle(root);
            if (title.StartsWith("ManualBuilder", StringComparison.OrdinalIgnoreCase)) return;
            lastExternalWindow = root;
        }

        private void RestoreTargetFocus()
        {
            if (lastExternalWindow != IntPtr.Zero) SetForegroundWindow(lastExternalWindow);
        }

        private Dictionary<string, object> ReadStatus()
        {
            if (!File.Exists(options.StatusPath)) return null;
            try
            {
                string raw;
                using (FileStream stream = new FileStream(options.StatusPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8, true)) raw = reader.ReadToEnd();
                if (String.IsNullOrWhiteSpace(raw)) return null;
                return json.Deserialize<Dictionary<string, object>>(raw);
            }
            catch { return null; }
        }

        private void PublishState()
        {
            Dictionary<string, object> status = ReadStatus();
            if (status != null && StringValue(status, "jobId") == options.JobId)
            {
                lastKnownState = StringValue(status, "state");
                currentCount = IntValue(status, "count");
            }
            PublishState(status);
        }

        private void PublishState(Dictionary<string, object> status)
        {
            if (!webReady) return;
            if (status == null) status = new Dictionary<string, object>();
            string state = StringValue(status, "state");
            if (String.IsNullOrEmpty(state)) state = lastKnownState;
            int count = status.ContainsKey("count") ? IntValue(status, "count") : currentCount;
            string target = StringValue(status, "lastTarget");
            string message = StringValue(status, "message");
            string stateLabel = StateLabel(state);
            if (!String.IsNullOrEmpty(pendingResultId) && resultRetryReady) stateLabel = "結果画面の追加結果を確認中";
            else if (!String.IsNullOrEmpty(pendingUndoId) && undoNoticeUntilUtc > DateTime.UtcNow) stateLabel = "直前の記録の削除結果を確認中";

            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["type"] = "state";
            payload["state"] = state;
            payload["stateLabel"] = stateLabel;
            payload["count"] = count;
            payload["target"] = !String.IsNullOrWhiteSpace(target) ? target : (count > 0 ? "直前の操作を記録しました" : "直前の操作はまだありません");
            payload["message"] = message;
            payload["help"] = helpText;
            payload["compact"] = compact;
            payload["closing"] = closingRequested;
            payload["pauseLabel"] = state == "paused" ? "記録を再開" : "一時停止";
            payload["undoLabel"] = String.IsNullOrEmpty(pendingUndoId) ? "直前の記録を削除" : "削除中…";
            payload["resultLabel"] = resultRetryReady ? "追加を再確認" : "結果画面を追加";
            payload["canPause"] = !closingRequested && (state == "recording" || state == "paused") && String.IsNullOrEmpty(pendingResultId);
            payload["canUndo"] = !closingRequested && count > 0 && (state == "recording" || state == "paused") && String.IsNullOrEmpty(pendingUndoId) && String.IsNullOrEmpty(pendingResultId);
            payload["canResult"] = !closingRequested && count > 0 && state == "recording" && lastExternalWindow != IntPtr.Zero && String.IsNullOrEmpty(pendingUndoId) && (String.IsNullOrEmpty(pendingResultId) || resultRetryReady);
            payload["canFinish"] = !closingRequested && (state == "recording" || state == "paused" || state == "failed");
            payload["finishLabel"] = state == "failed" ? "閉じる" : "終了して確認";
            payload["beforeImage"] = ImageUrl(count, false);
            payload["afterImage"] = ImageUrl(count, true);
            payload["recentOperations"] = RecentOperations();
            Publish(payload);
        }

        private object[] RecentOperations()
        {
            List<object> result = new List<object>();
            string jobDirectory = Path.GetDirectoryName(options.EventsDirectory);
            string eventsPath = Path.Combine(jobDirectory, "events.jsonl");
            if (!File.Exists(eventsPath)) return result.ToArray();
            try
            {
                string raw;
                using (FileStream stream = new FileStream(eventsPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8, true)) raw = reader.ReadToEnd();
                string[] lines = raw.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                int first = Math.Max(0, lines.Length - 3);
                for (int i = first; i < lines.Length; i++)
                {
                    Dictionary<string, object> item = json.Deserialize<Dictionary<string, object>>(lines[i]);
                    string kind = StringValue(item, "kind");
                    string target = StringValue(item, "targetName");
                    string label = !String.IsNullOrWhiteSpace(target)
                        ? target
                        : (kind == "input" ? "入力を記録しました" : "画面上の操作を記録しました");
                    Dictionary<string, object> receipt = new Dictionary<string, object>();
                    receipt["index"] = IntValue(item, "index");
                    receipt["label"] = label;
                    receipt["kindLabel"] = kind == "input" ? "入力" : (kind == "right-click" ? "右クリック" : "クリック");
                    result.Add(receipt);
                }
            }
            catch { }
            return result.ToArray();
        }

        private string ImageUrl(int count, bool result)
        {
            if (count <= 0) return String.Empty;
            string name = String.Format(result ? "event-{0:000}-result.jpg" : "event-{0:000}.jpg", count);
            string path = Path.Combine(options.EventsDirectory, name);
            if (!File.Exists(path)) return String.Empty;
            try
            {
                FileInfo info = new FileInfo(path);
                string stamp = info.Length + "-" + info.LastWriteTimeUtc.Ticks;
                return "https://recording.manualbuilder.local/events/" + Uri.EscapeDataString(name) + "?v=" + stamp;
            }
            catch { return String.Empty; }
        }

        private void Publish(Dictionary<string, object> payload)
        {
            if (!webReady || webView.CoreWebView2 == null) return;
            try { webView.CoreWebView2.PostWebMessageAsJson(json.Serialize(payload)); }
            catch (Exception ex) { Log(ex); }
        }

        private string StateLabel(string state)
        {
            if (state == "starting") return "記録の準備中";
            if (state == "recording") return "記録中";
            if (state == "paused") return "一時停止中";
            if (state == "completed") return "記録完了・ManualBuilderで確認できます";
            if (state == "failed") return "記録を続けられませんでした";
            return "記録状態を確認中";
        }

        private string DefaultHelp()
        {
            return "対象アプリでいつもどおり操作してください。クリックや入力は自動で記録されます。";
        }

        private void PositionAtWorkAreaEdge()
        {
            NativeRect area = GetWorkArea(lastExternalWindow);
            Left = area.Right / GetDpiScaleX() - Width - 18;
            Top = area.Bottom / GetDpiScaleY() - Height - 18;
            KeepInsideWorkArea();
        }

        private void KeepInsideWorkArea()
        {
            IntPtr reference = windowHandle != IntPtr.Zero ? windowHandle : lastExternalWindow;
            NativeRect area = GetWorkArea(reference);
            double sx = GetDpiScaleX();
            double sy = GetDpiScaleY();
            double left = area.Left / sx;
            double top = area.Top / sy;
            double right = area.Right / sx;
            double bottom = area.Bottom / sy;
            Left = Math.Max(left, Math.Min(Left, right - ActualWidth));
            Top = Math.Max(top, Math.Min(Top, bottom - ActualHeight));
        }

        private static NativeRect GetWorkArea(IntPtr reference)
        {
            IntPtr monitor = MonitorFromWindow(reference, MonitorDefaultToNearest);
            MonitorInfo info = new MonitorInfo();
            info.Size = Marshal.SizeOf(typeof(MonitorInfo));
            if (monitor == IntPtr.Zero || !GetMonitorInfo(monitor, ref info))
                throw new InvalidOperationException("記録モニターを表示する画面を特定できません。");
            return info.WorkArea;
        }

        private double GetDpiScaleX()
        {
            PresentationSource source = PresentationSource.FromVisual(this);
            return source == null ? 1.0 : source.CompositionTarget.TransformToDevice.M11;
        }

        private double GetDpiScaleY()
        {
            PresentationSource source = PresentationSource.FromVisual(this);
            return source == null ? 1.0 : source.CompositionTarget.TransformToDevice.M22;
        }

        private void FailStartup(string message)
        {
            try
            {
                Dictionary<string, object> status = ReadStatus() ?? new Dictionary<string, object>();
                status["jobId"] = options.JobId;
                status["state"] = "failed";
                status["message"] = message;
                WriteText(options.StatusPath, json.Serialize(status));
                WriteText(options.StopPath, "stop");
            }
            catch { }
            closingFromStatus = true;
            Close();
        }

        private bool TryWrite(string path, string value)
        {
            try { WriteText(path, value); return true; }
            catch (Exception ex) { Log(ex); return false; }
        }

        private static void WriteText(string path, string value)
        {
            File.WriteAllText(path, value, new UTF8Encoding(false));
        }

        private void Log(Exception ex)
        {
            try { File.AppendAllText(options.StatusPath + ".companion.log", ex + Environment.NewLine, new UTF8Encoding(false)); }
            catch { }
        }

        private void TraceTest(string value)
        {
            if (!options.TestMode) return;
            try { File.AppendAllText(options.StatusPath + ".companion.test.log", DateTime.UtcNow.ToString("o") + " " + value + Environment.NewLine, new UTF8Encoding(false)); }
            catch { }
        }

        private static string StringValue(Dictionary<string, object> values, string name)
        {
            object value;
            return values != null && values.TryGetValue(name, out value) && value != null ? Convert.ToString(value) : String.Empty;
        }

        private static int IntValue(Dictionary<string, object> values, string name)
        {
            object value;
            if (values == null || !values.TryGetValue(name, out value) || value == null) return 0;
            try { return Convert.ToInt32(value); } catch { return 0; }
        }

        private static string WindowTitle(IntPtr handle)
        {
            StringBuilder builder = new StringBuilder(512);
            GetWindowText(handle, builder, builder.Capacity);
            return builder.ToString();
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MonitorInfo
        {
            public int Size;
            public NativeRect MonitorArea;
            public NativeRect WorkArea;
            public uint Flags;
        }

        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hWnd, int flags);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool SetWindowDisplayAffinity(IntPtr hWnd, uint affinity);
        [DllImport("user32.dll")] private static extern IntPtr GetThreadDpiAwarenessContext();
        [DllImport("user32.dll")] private static extern int GetAwarenessFromDpiAwarenessContext(IntPtr value);
        [DllImport("user32.dll")] private static extern IntPtr MonitorFromWindow(IntPtr hWnd, uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr hWnd, int attribute, ref int value, int size);
    }
}
