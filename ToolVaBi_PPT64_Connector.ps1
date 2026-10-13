Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$pptApp = $null
$clsid64 = [System.Guid]"{91493440-5A91-11CF-8700-00AA0060263B}"

try {
    $pptApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application.16")
} catch {
    try {
        $pptApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    } catch {
        try {
            $pptType = [System.Type]::GetTypeFromCLSID($clsid64)
            $pptApp = [System.Activator]::CreateInstance($pptType)
            $pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
        } catch { }
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "ToolVaBi PowerPoint (Office 64-bit Edition)"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true

$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = "Top"
$panel.Height = 65
$panel.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "LaTeX / TikZ:"
$lblInput.Location = New-Object System.Drawing.Point(12, 22)
$lblInput.AutoSize = $true
$lblInput.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

$txtCode = New-Object System.Windows.Forms.TextBox
$txtCode.Text = "\int_{a}^{b} f(x) dx = F(b) - F(a)"
$txtCode.Location = New-Object System.Drawing.Point(110, 19)
$txtCode.Size = New-Object System.Drawing.Size(460, 26)
$txtCode.Font = New-Object System.Drawing.Font("Consolas", 10)

$btnInsert = New-Object System.Windows.Forms.Button
$btnInsert.Text = "⚡ Chèn Vào Slide 64-bit"
$btnInsert.Location = New-Object System.Drawing.Point(585, 16)
$btnInsert.Size = New-Object System.Drawing.Size(200, 32)
$btnInsert.BackColor = [System.Drawing.Color]::FromArgb(16, 137, 255)
$btnInsert.ForeColor = [System.Drawing.Color]::White
$btnInsert.FlatStyle = "Flat"
$btnInsert.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Đang kết nối PowerPoint 64-bit..."
$lblStatus.Location = New-Object System.Drawing.Point(800, 23)
$lblStatus.AutoSize = $true
$lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen

$panel.Controls.Add($lblInput)
$panel.Controls.Add($txtCode)
$panel.Controls.Add($btnInsert)
$panel.Controls.Add($lblStatus)
$form.Controls.Add($panel)

$wb = New-Object System.Windows.Forms.WebBrowser
$wb.Dock = "Fill"
$wb.ScriptErrorsSuppressed = $true
$wb.Navigate("https://latex.codecogs.com/eqneditor/editor.php")
$form.Controls.Add($wb)

$btnInsert.Add_Click({
    try {
        $lblStatus.Text = "Đang biên dịch Online 300 DPI..."
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
        $form.Refresh()

        $latex = $txtCode.Text.Trim()
        $encoded = [System.Uri]::EscapeDataString($latex)
        $url = "https://latex.codecogs.com/png.image?\dpi{300}\bg{white}$encoded"
        $tmp = Join-Path $env:TEMP "ppt64_eq.png"

        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("User-Agent", "ToolVaBiPPT64/1.0")
        $wc.DownloadFile($url, $tmp)

        if ($null -ne $pptApp) {
            $slide = $null
            if ($pptApp.ActivePresentation.Slides.Count -eq 0) {
                $slide = $pptApp.ActivePresentation.Slides.Add(1, 12)
            } else {
                $slide = $pptApp.ActiveWindow.View.Slide
            }

            $slideW = $pptApp.ActivePresentation.PageSetup.SlideWidth
            $slideH = $pptApp.ActivePresentation.PageSetup.SlideHeight
            $left = ($slideW - 320) / 2
            $top = ($slideH - 120) / 2

            $slide.Shapes.AddPicture($tmp, 0, 1, $left, $top, 320, 120) | Out-Null
            $lblStatus.Text = "Đã chèn vào Slide 64-bit thành công!"
            $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
        } else {
            $img = [System.Drawing.Image]::FromFile($tmp)
            [System.Windows.Forms.Clipboard]::SetImage($img)
            $img.Dispose()

            $wsh = New-Object -ComObject WScript.Shell
            $wsh.AppActivate("PowerPoint")
            Start-Sleep -Milliseconds 250
            $wsh.SendKeys("^v")

            $lblStatus.Text = "Đã dán trực tiếp vào Slide PowerPoint!"
            $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
        }
    } catch {
        $lblStatus.Text = "Lỗi: " + $_.Exception.Message
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

$form.ShowDialog()
