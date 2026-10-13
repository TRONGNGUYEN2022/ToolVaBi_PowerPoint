using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace ToolVaBiPPT
{
    public class RibbonPPT : RibbonBase
    {
        private RibbonTab? _tabToolVaBi;
        private RibbonGroup? _groupMath;
        private RibbonGroup? _groupTikz;
        private RibbonGroup? _groupAbout;
        private RibbonButton? _btnOpenEditor;
        private RibbonButton? _btnQuickInsertMath;
        private RibbonButton? _btnCompileTikz;
        private RibbonButton? _btnAbout;

        public RibbonPPT() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeRibbonComponents();
        }

        private void InitializeRibbonComponents()
        {
            this._tabToolVaBi = this.Factory.CreateRibbonTab();
            this._groupMath = this.Factory.CreateRibbonGroup();
            this._groupTikz = this.Factory.CreateRibbonGroup();
            this._groupAbout = this.Factory.CreateRibbonGroup();
            this._btnOpenEditor = this.Factory.CreateRibbonButton();
            this._btnQuickInsertMath = this.Factory.CreateRibbonButton();
            this._btnCompileTikz = this.Factory.CreateRibbonButton();
            this._btnAbout = this.Factory.CreateRibbonButton();

            // Tab
            this._tabToolVaBi.ControlId.ControlIdType = RibbonControlIdType.Office;
            this._tabToolVaBi.Label = "ToolVaBi PPT";
            this._tabToolVaBi.Name = "tabToolVaBi";

            // Nhóm Toán Học
            this._groupMath.Label = "Soạn Thảo Toán Học";
            this._btnOpenEditor.Label = "Trình Soạn Toán (MathLive)";
            this._btnOpenEditor.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this._btnOpenEditor.Click += (s, e) => ThisAddIn.Instance?.OpenMathEditor();

            this._btnQuickInsertMath.Label = "Chèn Công Thức Nhanh";
            this._btnQuickInsertMath.Click += BtnQuickInsertMath_Click;

            this._groupMath.Items.Add(this._btnOpenEditor);
            this._groupMath.Items.Add(this._btnQuickInsertMath);

            // Nhóm TikZ Online
            this._groupTikz.Label = "Biên Dịch TikZ & LaTeX";
            this._btnCompileTikz.Label = "Biên Dịch TikZ Online";
            this._btnCompileTikz.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this._btnCompileTikz.Click += BtnCompileTikz_Click;
            this._groupTikz.Items.Add(this._btnCompileTikz);

            // Nhóm Thông Tin
            this._groupAbout.Label = "Thông Tin";
            this._btnAbout.Label = "Tác Giả & Hướng Dẫn";
            this._btnAbout.Click += (s, e) =>
            {
                MessageBox.Show("ToolVaBi PowerPoint Add-in (Modern WebView2 + MathLive + TikZ Online)\n" +
                                "Phiên bản dành riêng cho thiết kế Slide bài giảng toán học và khoa học.\n\n" +
                                "Hỗ trợ biên dịch LaTeX/TikZ trực tuyến thành ảnh chất lượng cao 300 DPI.",
                                "Giới Thiệu ToolVaBi PPT", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            this._groupAbout.Items.Add(this._btnAbout);

            this._tabToolVaBi.Groups.Add(this._groupMath);
            this._tabToolVaBi.Groups.Add(this._groupTikz);
            this._tabToolVaBi.Groups.Add(this._groupAbout);
            this.Tabs.Add(this._tabToolVaBi);
        }

        private async void BtnQuickInsertMath_Click(object sender, RibbonControlEventArgs e)
        {
            string defaultFormula = "\\int_{a}^{b} f(x) dx = F(b) - F(a)";
            try
            {
                string pngPath = await TikzOnlineCompiler.CompileLatexToPngAsync(defaultFormula, 300);
                ThisAddIn.Instance?.InsertImageToActiveSlide(pngPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi biên dịch: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void BtnCompileTikz_Click(object sender, RibbonControlEventArgs e)
        {
            // Mẫu hình vẽ đồ thị parabol bằng TikZ
            string tikzSample = @"\begin{tikzpicture}[scale=0.8]
  \draw[thick,->] (-2.5,0) -- (2.5,0) node[right] {$x$};
  \draw[thick,->] (0,-0.5) -- (0,4.5) node[above] {$y$};
  \draw[scale=1,domain=-2:2,smooth,variable=\x,blue,thick] plot ({\x},{\x*\x});
  \node[below right] at (0,0) {$O$};
  \node[above right,blue] at (1.5,2.25) {$y=x^2$};
\end{tikzpicture}";

            try
            {
                string pngPath = await TikzOnlineCompiler.CompileTikzToPngAsync(tikzSample);
                ThisAddIn.Instance?.InsertImageToActiveSlide(pngPath, 350, 300);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi biên dịch TikZ Online: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
