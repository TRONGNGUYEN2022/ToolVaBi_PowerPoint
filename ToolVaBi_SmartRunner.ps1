Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$tempDir = "E:\VSCode_Automation\ToolVaBi_PowerPoint\output_images"
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null

$form = New-Object System.Windows.Forms.Form
$form.Text = "ToolVaBi PPT - Soạn Toán & TikZ (Smart Injector)"
$form.Size = New-Object System.Drawing.Size(950, 680)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true

$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = "Top"
$panel.Height = 60
$panel.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)

$lblLatex = New-Object System.Windows.Forms.Label
$lblLatex.Text = "Mã LaTeX / TikZ:"
$lblLatex.Location = New-Object System.Drawing.Point(12, 20)
$lblLatex.AutoSize = $true
$lblLatex.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

$txtLatex = New-Object System.Windows.Forms.TextBox
$txtLatex.Text = "\int_{a}^{b} f(x) dx = F(b) - F(a)"
$txtLatex.Location = New-Object System.Drawing.Point(135, 17)
$txtLatex.Size = New-Object System.Drawing.Size(420, 26)

$btnInsert = New-Object System.Windows.Forms.Button
$btnInsert.Text = "⚡ Chèn Ngay Vào Slide"
$btnInsert.Location = New-Object System.Drawing.Point(570, 14)
$btnInsert.Size = New-Object System.Drawing.Size(180, 32)
$btnInsert.BackColor = [System.Drawing.Color]::FromArgb(16, 137, 255)
$btnInsert.ForeColor = [System.Drawing.Color]::White
$btnInsert.FlatStyle = "Flat"
$btnInsert.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Sẵn sàng chèn vào PowerPoint!"
$lblStatus.Location = New-Object System.Drawing.Point(760, 21)
$lblStatus.AutoSize = $true
$lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen

$panel.Controls.Add($lblLatex)
$panel.Controls.Add($txtLatex)
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
        $lblStatus.Text = "Đang biên dịch 300 DPI..."
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
        $form.Refresh()

        $latexCode = $txtLatex.Text.Trim()
        $encoded = [System.Uri]::EscapeDataString($latexCode)
        $url = "https://latex.codecogs.com/png.image?\dpi{300}\bg{white}$encoded"
        $imgFile = Join-Path $tempDir "slide_eq_$([System.Guid]::NewGuid().ToString('N')).png"

        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("User-Agent", "ToolVaBiPPT/1.0")
        $wc.DownloadFile($url, $imgFile)

        $img = [System.Drawing.Image]::FromFile($imgFile)
        [System.Windows.Forms.Clipboard]::SetImage($img)
        $img.Dispose()

        $wsh = New-Object -ComObject WScript.Shell
        $activated = $wsh.AppActivate("PowerPoint")

        if ($activated) {
            Start-Sleep -Milliseconds 250
            $wsh.SendKeys("^v")
            $lblStatus.Text = "Đã chèn vào Slide thành công!"
            $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
        } else {
            $lblStatus.Text = "Đã copy ảnh! Nhấp vào Slide và ấn Ctrl+V."
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
        }
    } catch {
        $lblStatus.Text = "Lỗi: " + $_.Exception.Message
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

$form.ShowDialog()
