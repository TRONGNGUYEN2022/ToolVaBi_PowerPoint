using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ToolVaBiPPT
{
    public class MathEditorForm : Form
    {
        private WebView2? _webView;
        private Panel? _panelTop;
        private Button? _btnInsertToSlide;
        private Button? _btnCompileTikzOnline;
        private Label? _lblStatus;

        public MathEditorForm()
        {
            InitializeComponent();
            this.Load += async (s, e) => await InitializeWebViewAsync();
        }

        private void InitializeComponent()
        {
            this.Text = "ToolVaBi PPT - Bộ Soạn Thảo Toán & Biên Dịch Online";
            this.Size = new Size(1150, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;

            _panelTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = Color.FromArgb(248, 250, 252),
                Padding = new Padding(12)
            };

            _btnInsertToSlide = new Button
            {
                Text = "⚡ Chèn Công Thức Vào Slide",
                Location = new Point(14, 14),
                Width = 220,
                Height = 32,
                BackColor = Color.FromArgb(16, 137, 255),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            _btnInsertToSlide.FlatAppearance.BorderSize = 0;
            _btnInsertToSlide.Click += (s, e) => SendIpcMessage("requestFormulaForSlide");

            _btnCompileTikzOnline = new Button
            {
                Text = "⚡ Biên Dịch TikZ Online Sang Slide",
                Location = new Point(245, 14),
                Width = 250,
                Height = 32,
                BackColor = Color.FromArgb(22, 163, 74),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            _btnCompileTikzOnline.FlatAppearance.BorderSize = 0;
            _btnCompileTikzOnline.Click += (s, e) => SendIpcMessage("requestTikzForSlide");

            _lblStatus = new Label
            {
                Text = "Trạng thái: Đang nạp hệ thống...",
                Location = new Point(510, 22),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
                ForeColor = Color.DarkSlateGray
            };

            _panelTop.Controls.Add(_btnInsertToSlide);
            _panelTop.Controls.Add(_btnCompileTikzOnline);
            _panelTop.Controls.Add(_lblStatus);

            _webView = new WebView2 { Dock = DockStyle.Fill };

            this.Controls.Add(_webView);
            this.Controls.Add(_panelTop);
        }

        private async Task InitializeWebViewAsync()
        {
            try
            {
                string appDir = AppDomain.CurrentDomain.BaseDirectory;
                string userData = Path.Combine(appDir, "WebView2_UserData_PPT");
                string assets = Path.Combine(appDir, "assets");

                var env = await CoreWebView2Environment.CreateAsync(null, userData);
                if (_webView != null)
                {
                    await _webView.EnsureCoreWebView2Async(env);
                    _webView.CoreWebView2.SetVirtualHostNameToFolderMapping("toolvabippt.local", assets, CoreWebView2HostResourceAccessKind.Allow);
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                    _webView.CoreWebView2.Navigate("https://toolvabippt.local/editor.html");

                    if (_lblStatus != null)
                    {
                        _lblStatus.Text = "Hệ thống sẵn sàng!";
                        _lblStatus.ForeColor = Color.ForestGreen;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi WebView2: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SendIpcMessage(string action)
        {
            if (_webView?.CoreWebView2 == null) return;
            string json = JsonSerializer.Serialize(new { action = action });
            _webView.CoreWebView2.PostWebMessageAsJson(json);
        }

        private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                using var doc = JsonDocument.Parse(e.WebMessageAsJson);
                var root = doc.RootElement;
                string action = root.GetProperty("action").GetString() ?? "";

                if (action == "insertFormulaToSlide")
                {
                    string latex = root.GetProperty("latex").GetString() ?? "";
                    if (!string.IsNullOrWhiteSpace(latex))
                    {
                        UpdateStatus("Đang biên dịch LaTeX online (300 DPI)...", Color.DarkBlue);
                        string png = await TikzOnlineCompiler.CompileLatexToPngAsync(latex, 300);
                        bool ok = ThisAddIn.Instance?.InsertImageToActiveSlide(png) ?? false;
                        UpdateStatus(ok ? "Chèn công thức thành công!" : "Lỗi chèn công thức!", ok ? Color.ForestGreen : Color.Red);
                    }
                }
                else if (action == "insertTikzToSlide")
                {
                    string tikz = root.GetProperty("tikz").GetString() ?? "";
                    if (!string.IsNullOrWhiteSpace(tikz))
                    {
                        UpdateStatus("Đang biên dịch hình vẽ TikZ Online...", Color.DarkBlue);
                        string png = await TikzOnlineCompiler.CompileTikzToPngAsync(tikz);
                        bool ok = ThisAddIn.Instance?.InsertImageToActiveSlide(png, 400, 320) ?? false;
                        UpdateStatus(ok ? "Chèn hình vẽ TikZ thành công!" : "Lỗi chèn TikZ!", ok ? Color.ForestGreen : Color.Red);
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Lỗi: {ex.Message}", Color.Red);
            }
        }

        private void UpdateStatus(string text, Color color)
        {
            if (_lblStatus != null)
            {
                _lblStatus.Text = text;
                _lblStatus.ForeColor = color;
            }
        }
    }
}
