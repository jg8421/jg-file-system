using System;
using System.Drawing;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

internal sealed class NativePdfPreview : Form
{
    private readonly IntPtr parentHandle;
    private readonly string documentPath;
    private readonly string pipeName;
    private readonly Timer resizeTimer = new Timer { Interval = 150 };
    private readonly WebView2 browser = new WebView2 { Dock = DockStyle.Fill };

    [DllImport("user32.dll")] private static extern IntPtr SetParent(IntPtr child, IntPtr parent);
    [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr handle, out RECT rect);
    [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr handle, IntPtr after, int x, int y, int width, int height, uint flags);
    private static readonly IntPtr HWND_TOP = IntPtr.Zero;
    private const uint SWP_NOACTIVATE = 0x0010;
    private const uint SWP_SHOWWINDOW = 0x0040;

    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    internal NativePdfPreview(IntPtr parent, string path)
    {
        parentHandle = parent;
        documentPath = path;
        pipeName = "FileWorkbenchPdfPreview-" + parent.ToInt64();
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        Controls.Add(browser);
        Shown += async (sender, args) =>
        {
            try
            {
                SetParent(Handle, parentHandle);
                ResizeToParent();
                resizeTimer.Tick += (s, e) => ResizeToParent();
                resizeTimer.Start();
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "FileWorkbench", "WebView2");
                var webViewEnvironment = await CoreWebView2Environment.CreateAsync(null, userDataFolder, null);
                await browser.EnsureCoreWebView2Async(webViewEnvironment);
                browser.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
                browser.CoreWebView2.Settings.AreDevToolsEnabled = false;
                NavigateTo(documentPath);
                StartCommandListener();
            }
            catch (Exception error)
            {
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "native_pdf_preview.log"), error.ToString());
                Close();
            }
        };
        FormClosed += (sender, args) => resizeTimer.Stop();
    }

    private void NavigateTo(string path)
    {
        if (File.Exists(path)) browser.Source = new Uri(path);
    }

    private void StartCommandListener()
    {
        Task.Factory.StartNew(() =>
        {
            while (!IsDisposed)
            {
                try
                {
                    using (var server = new NamedPipeServerStream(pipeName, PipeDirection.In, 1, PipeTransmissionMode.Byte, PipeOptions.None))
                    {
                        server.WaitForConnection();
                        string nextPath;
                        using (var reader = new StreamReader(server)) nextPath = reader.ReadToEnd();
                        if (!String.IsNullOrWhiteSpace(nextPath) && !IsDisposed)
                            BeginInvoke(new Action(() => NavigateTo(nextPath)));
                    }
                }
                catch (ObjectDisposedException) { return; }
                catch (InvalidOperationException) { return; }
            }
        }, TaskCreationOptions.LongRunning);
    }

    private void ResizeToParent()
    {
        RECT rect;
        if (!GetClientRect(parentHandle, out rect)) return;
        SetWindowPos(Handle, HWND_TOP, 0, 0, Math.Max(1, rect.Right - rect.Left), Math.Max(1, rect.Bottom - rect.Top), SWP_NOACTIVATE | SWP_SHOWWINDOW);
    }

    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length != 2) Environment.Exit(2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new NativePdfPreview(new IntPtr(long.Parse(args[0])), args[1]));
    }
}
