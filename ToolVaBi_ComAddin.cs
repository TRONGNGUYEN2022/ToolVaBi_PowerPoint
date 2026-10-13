using System;
using System.IO;
using System.Net;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ToolVaBi
{
    public enum ext_ConnectMode { ext_cm_AfterStartup = 0, ext_cm_Startup = 1, ext_cm_External = 2, ext_cm_CommandLine = 3 }
    public enum ext_DisconnectMode { ext_dm_HostShutdown = 0, ext_dm_UserClosed = 1 }

    [ComImport, Guid("B65CC1D4-DA27-11D2-8F7A-0000F87A77B1"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object Application, ext_ConnectMode ConnectMode, [MarshalAs(UnmanagedType.IDispatch)] object AddInInst, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        void OnDisconnection(ext_DisconnectMode RemoveMode, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        void OnAddInsUpdate([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        void OnStartupComplete([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        void OnBeginShutdown([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComImport, Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    [ComImport, Guid("000C0395-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonControl
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string Id { get; }
        [return: MarshalAs(UnmanagedType.BStr)]
        string Tag { get; }
        object Context { get; }
    }

    [ComVisible(true)]
    [Guid("E83B6745-927A-4D3B-A91B-2B2D6284C8F1")]
    [ProgId("ToolVaBi.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private object _app;

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _app = Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom) { _app = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabToolVaBi"" label=""ToolVaBi PPT"">
        <group id=""grpMath"" label=""Toán Học"">
          <button id=""btnInsertQuick"" label=""Chèn Công Thức Nhanh"" size=""large"" imageMso=""EquationInsert"" onAction=""OnQuickInsert"" />
          <button id=""btnOpenDialog"" label=""Soạn Toán MathLive"" size=""large"" imageMso=""EquationInsertGallery"" onAction=""OnOpenMathEditor"" />
        </group>
        <group id=""grpTikz"" label=""Biên Dịch Online"">
          <button id=""btnCompileTikz"" label=""Vẽ Đồ Thị TikZ"" size=""large"" imageMso=""ChartTypeInsert"" onAction=""OnCompileTikz"" />
        </group>
        <group id=""grpInfo"" label=""Thông Tin"">
          <button id=""btnAbout"" label=""Giới Thiệu"" size=""normal"" imageMso=""Info"" onAction=""OnAbout"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnQuickInsert(IRibbonControl control)
        {
            InsertLatexToActiveSlide(@"\int_{-\infty}^{+\infty} e^{-x^2} dx = \sqrt{\pi}");
        }

        public void OnCompileTikz(IRibbonControl control)
        {
            string tikz = @"\begin{tikzpicture}\draw[thick,->] (-2.5,0) -- (2.5,0) node[right]{$x$};\draw[thick,->] (0,-0.5) -- (0,4) node[above]{$y$};\draw[red,very thick,domain=-1.5:1.5] plot (\x,{\x*\x});\end{tikzpicture}";
            InsertTikzToActiveSlide(tikz);
        }

        public void OnOpenMathEditor(IRibbonControl control)
        {
            ShowEditorForm();
        }

        public void OnAbout(IRibbonControl control)
        {
            MessageBox.Show("ToolVaBi PowerPoint Add-in (Native COM Ribbon Edition)\n\n" +
                            "Tích hợp vĩnh viễn trên thanh Ribbon PowerPoint 64-bit mà không cần ClickOnce.",
                            "ToolVaBi PPT", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void InsertLatexToActiveSlide(string latex)
        {
            try
            {
                string encoded = Uri.EscapeDataString(latex.Trim());
                string url = "https://latex.codecogs.com/png.image?\\dpi{300}\\bg{white}" + encoded;
                string tmpFile = Path.Combine(Path.GetTempPath(), "math_" + Guid.NewGuid().ToString("N") + ".png");

                using (var wc = new WebClient())
                {
                    wc.Headers.Add("User-Agent", "ToolVaBiPPT/1.0");
                    wc.DownloadFile(url, tmpFile);
                }

                AddPictureToSlide(tmpFile, 320, 120);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "ToolVaBi PPT");
            }
        }

        private void InsertTikzToActiveSlide(string tikz)
        {
            try
            {
                string encoded = Uri.EscapeDataString(tikz.Trim());
                string url = "https://quickchart.io/latex?formula=" + encoded + "&density=300";
                string tmpFile = Path.Combine(Path.GetTempPath(), "tikz_" + Guid.NewGuid().ToString("N") + ".png");

                using (var wc = new WebClient())
                {
                    wc.Headers.Add("User-Agent", "ToolVaBiPPT/1.0");
                    wc.DownloadFile(url, tmpFile);
                }

                AddPictureToSlide(tmpFile, 350, 300);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "ToolVaBi PPT");
            }
        }

        private void AddPictureToSlide(string imgPath, float width, float height)
        {
            if (_app == null) return;
            try
            {
                dynamic app = _app;
                dynamic slide = null;

                if (app.ActivePresentation.Slides.Count == 0)
                {
                    slide = app.ActivePresentation.Slides.Add(1, 12);
                }
                else
                {
                    slide = app.ActiveWindow.View.Slide;
                }

                float sw = (float)app.ActivePresentation.PageSetup.SlideWidth;
                float sh = (float)app.ActivePresentation.PageSetup.SlideHeight;
                float left = (sw - width) / 2;
                float top = (sh - height) / 2;

                slide.Shapes.AddPicture(imgPath, 0, 1, left, top, width, height);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi chèn ảnh vào Slide: " + ex.Message, "ToolVaBi PPT");
            }
        }

        private void ShowEditorForm()
        {
            Form form = new Form();
            form.Text = "Trình Soạn Toán Nhanh - ToolVaBi PPT";
            form.Size = new Size(800, 220);
            form.StartPosition = FormStartPosition.CenterScreen;
            form.TopMost = true;

            Label lbl = new Label() { Text = "Mã LaTeX:", Location = new Point(14, 20), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            TextBox txt = new TextBox() { Text = @"\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}", Location = new Point(95, 17), Width = 550, Font = new Font("Consolas", 10.5f) };
            Button btn = new Button() { Text = "Chèn Vào Slide", Location = new Point(660, 15), Width = 110, Height = 30, BackColor = Color.FromArgb(16, 137, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

            btn.Click += (s, e) =>
            {
                InsertLatexToActiveSlide(txt.Text);
                form.Close();
            };

            form.Controls.Add(lbl);
            form.Controls.Add(txt);
            form.Controls.Add(btn);
            form.ShowDialog();
        }
    }
}
