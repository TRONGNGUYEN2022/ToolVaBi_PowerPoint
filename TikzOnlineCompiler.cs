using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ToolVaBiPPT
{
    /// <summary>
    /// Module biên dịch LaTeX & TikZ trực tuyến (Online Compiler)
    /// Hỗ trợ cả công thức toán LaTeX và khối hình học TikZ sang ảnh PNG 300 DPI
    /// </summary>
    public static class TikzOnlineCompiler
    {
        /// <summary>
        /// Biên dịch công thức LaTeX sang file PNG chất lượng cao
        /// </summary>
        public static async Task<string> CompileLatexToPngAsync(string latexCode, int dpi = 300)
        {
            if (string.IsNullOrWhiteSpace(latexCode))
                throw new ArgumentException("Mã LaTeX không được để trống!");

            string cleanLatex = latexCode.Trim();
            // Nếu người dùng bọc $ hoặc $$, loại bỏ để API xử lý chuẩn
            if (cleanLatex.StartsWith("$$") && cleanLatex.EndsWith("$$") && cleanLatex.Length > 4)
                cleanLatex = cleanLatex.Substring(2, cleanLatex.Length - 4).Trim();
            else if (cleanLatex.StartsWith("$") && cleanLatex.EndsWith("$") && cleanLatex.Length > 2)
                cleanLatex = cleanLatex.Substring(1, cleanLatex.Length - 2).Trim();

            string tempFile = Path.Combine(Path.GetTempPath(), $"latex_{Guid.NewGuid():N}.png");

            // URL Endpoint QuickChart / CodeCogs với độ phân giải cao
            string encoded = Uri.EscapeDataString(cleanLatex);
            string requestUrl = $"https://latex.codecogs.com/png.image?\dpi{{{dpi}}}\bg{{white}}{encoded}";

            using (var client = new WebClient())
            {
                client.Headers.Add("User-Agent", "ToolVaBiPPT/1.0 (Windows NT 10.0; Win64; x64)");
                await client.DownloadFileTaskAsync(new Uri(requestUrl), tempFile);
            }

            return tempFile;
        }

        /// <summary>
        /// Biên dịch mã TikZ (hình vẽ hình học, đồ thị hàm số) sang PNG qua REST API online
        /// </summary>
        public static async Task<string> CompileTikzToPngAsync(string tikzCode)
        {
            if (string.IsNullOrWhiteSpace(tikzCode))
                throw new ArgumentException("Mã TikZ không được để trống!");

            string fullDocument = BuildFullTikzDocument(tikzCode);
            string tempPng = Path.Combine(Path.GetTempPath(), $"tikz_{Guid.NewGuid():N}.png");

            // Phương thức 1: Sử dụng RTeX Compiler API (dịch vụ biên dịch LaTeX/TikZ online tiêu chuẩn)
            try
            {
                string rtexUrl = "https://rtex.probablyprogramming.com/api/v2";
                string jsonPayload = Newtonsoft.Json.JsonConvert.SerializeObject(new
                {
                    code = fullDocument,
                    format = "png"
                });

                var request = (HttpWebRequest)WebRequest.Create(rtexUrl);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.Timeout = 25000;

                byte[] bytes = Encoding.UTF8.GetBytes(jsonPayload);
                request.ContentLength = bytes.Length;
                using (var stream = await request.GetRequestStreamAsync())
                {
                    await stream.WriteAsync(bytes, 0, bytes.Length);
                }

                using (var response = (HttpWebResponse)await request.GetResponseAsync())
                using (var reader = new StreamReader(response.GetResponseStream()))
                {
                    string resultJson = await reader.ReadToEndAsync();
                    dynamic res = Newtonsoft.Json.JsonConvert.DeserializeObject(resultJson);
                    if (res != null && res.status == "success")
                    {
                        string downloadUrl = "https://rtex.probablyprogramming.com/api/v2/" + (string)res.filename;
                        using (var wc = new WebClient())
                        {
                            await wc.DownloadFileTaskAsync(new Uri(downloadUrl), tempPng);
                        }
                        return tempPng;
                    }
                }
            }
            catch
            {
                // Fallback: Nếu RTeX bận, biên dịch qua QuickChart / LaTeX Standalone Engine
            }

            // Fallback Engine
            string fallbackUrl = $"https://quickchart.io/latex?formula={Uri.EscapeDataString(tikzCode)}&density=300";
            using (var client = new WebClient())
            {
                await client.DownloadFileTaskAsync(new Uri(fallbackUrl), tempPng);
            }

            return tempPng;
        }

        private static string BuildFullTikzDocument(string body)
        {
            if (body.Contains("\documentclass")) return body;

            var sb = new StringBuilder();
            sb.AppendLine("\documentclass[tikz,border=4pt]{standalone}");
            sb.AppendLine("\usepackage{tikz}");
            sb.AppendLine("\usepackage{amsmath,amssymb}");
            sb.AppendLine("\usepackage{pgfplots}");
            sb.AppendLine("\pgfplotsset{compat=1.18}");
            sb.AppendLine("\begin{document}");
            sb.AppendLine(body);
            sb.AppendLine("\end{document}");
            return sb.ToString();
        }
    }
}
