using System;
using System.IO;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ToolVaBiPPT
{
    public partial class ThisAddIn
    {
        public static ThisAddIn? Instance { get; private set; }
        private MathEditorForm? _mathEditorForm;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Instance = this;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (_mathEditorForm != null && !_mathEditorForm.IsDisposed)
            {
                _mathEditorForm.Close();
            }
        }

        public void OpenMathEditor()
        {
            if (_mathEditorForm == null || _mathEditorForm.IsDisposed)
            {
                _mathEditorForm = new MathEditorForm();
            }
            _mathEditorForm.Show();
            _mathEditorForm.BringToFront();
        }

        public bool InsertImageToActiveSlide(string imagePath, float width = -1, float height = -1)
        {
            try
            {
                if (!File.Exists(imagePath)) return false;

                PowerPoint.Slide? slide = null;
                try
                {
                    slide = Application.ActiveWindow.View.Slide as PowerPoint.Slide;
                }
                catch
                {
                    if (Application.ActivePresentation != null && Application.ActivePresentation.Slides.Count > 0)
                    {
                        slide = Application.ActivePresentation.Slides[1];
                    }
                }

                if (slide == null)
                {
                    MessageBox.Show("Vui lòng chọn một Slide trong PowerPoint trước khi chèn!", 
                        "ToolVaBi PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // Kích thước chuẩn slide PowerPoint (16:9 thường là 960x540 pt hoặc 720x540 pt)
                float slideWidth = Application.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = Application.ActivePresentation.PageSetup.SlideHeight;

                // Thêm hình ảnh vào giữa slide
                float left = (slideWidth - (width > 0 ? width : 300)) / 2;
                float top = (slideHeight - (height > 0 ? height : 200)) / 2;

                var shape = slide.Shapes.AddPicture(
                    imagePath,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    left,
                    top,
                    width,
                    height
                );

                shape.Select();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi chèn ảnh vào Slide: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
