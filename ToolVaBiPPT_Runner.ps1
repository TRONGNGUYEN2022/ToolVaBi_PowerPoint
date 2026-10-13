Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$pptApp = $null
try {
    $pptApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
} catch {
    Write-Host ">>> Đang kết nối PowerPoint..." -ForegroundColor Yellow
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "ToolVaBi PowerPoint - Soạn Toán & TikZ Online"
$form.Size = New-Object System.Drawing.Size(1000, 680)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true

$btnPanel = New-Object System.Windows.Forms.Panel
$btnPanel.Dock = "Top"
$btnPanel.Height = 55
$btnPanel.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)

$btnInsert = New-Object System.Windows.Forms.Button
$btnInsert.Text = "⚡ Chèn Công Thức Vào Slide Đang Mở"
$btnInsert.Location = New-Object System.Drawing.Point(14, 12)
$btnInsert.Size = New-Object System.Drawing.Size(260, 32)
$btnInsert.BackColor = [System.Drawing.Color]::FromArgb(16, 137, 255)
$btnInsert.ForeColor = [System.Drawing.Color]::White
$btnInsert.FlatStyle = "Flat"
$btnInsert.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

$btnInsert.Add_Click({
    try {
        if ($pptApp.ActivePresentation.Slides.Count -eq 0) {
            $slide = $pptApp.ActivePresentation.Slides.Add(1, 12)
        } else {
            $slide = $pptApp.ActiveWindow.View.Slide
        }

        $latex = "\int_{a}^{b} f(x) dx = F(b) - F(a)"
        $encoded = [System.Uri]::EscapeDataString($latex)
        $url = "https://latex.codecogs.com/png.image?\dpi{300}\bg{white}$encoded"
        $tmpFile = Join-Path $env:TEMP "latex_slide.png"

        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($url, $tmpFile)

        $slideW = $pptApp.ActivePresentation.PageSetup.SlideWidth
        $slideH = $pptApp.ActivePresentation.PageSetup.SlideHeight
        $left = ($slideW - 300) / 2
        $top = ($slideH - 100) / 2

        $shape = $slide.Shapes.AddPicture($tmpFile, 0, 1, $left, $top, 300, 100)
        [System.Windows.Forms.MessageBox]::Show("Đã chèn công thức vào Slide thành công!", "ToolVaBi PPT")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Lỗi: " + $_.Exception.Message, "Lỗi")
    }
})

$btnPanel.Controls.Add($btnInsert)
$form.Controls.Add($btnPanel)

$wb = New-Object System.Windows.Forms.WebBrowser
$wb.Dock = "Fill"
$wb.Navigate("https://cdn.jsdelivr.net/npm/mathlive/dist/index.html")
$form.Controls.Add($wb)

$form.ShowDialog()
