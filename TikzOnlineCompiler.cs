using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ToolVaBiPPT
{
    public static class TikzOnlineCompiler
    {
        public static async Task<string> CompileLatexToPngAsync(string latexCode, int dpi = 300)
        {
            if (string.IsNullOrWhiteSpace(latexCode))
                throw new ArgumentException("Mã LaTeX không được rỗng!");

            string clean = latexCode.Trim();
            if (clean.StartsWith("$$") && clean.EndsWith("$$") && clean.Length > 4)
                clean = clean.Substring(2, clean.Length - 4).Trim();
            else if (clean.StartsWith("$") && clean.EndsWith("$") && clean.Length > 2)
                clean = clean.Substring(1, clean.Length - 2).Trim();

            string tempFile = Path.Combine(Path.GetTempPath(), "latex_" + Guid.NewGuid().ToString("N") + ".png");
            string encoded = Uri.EscapeDataString(clean);
            string requestUrl = "https://latex.codecogs.com/png.image?\\dpi{" + dpi + "}\\bg{white}" + encoded;

            using (var client = new WebClient())
            {
                client.Headers.Add("User-Agent", "ToolVaBiPPT/1.0");
                await client.DownloadFileTaskAsync(new Uri(requestUrl), tempFile);
            }

            return tempFile;
        }

        public static async Task<string> CompileTikzToPngAsync(string tikzCode)
        {
            if (string.IsNullOrWhiteSpace(tikzCode))
                throw new ArgumentException("Mã TikZ không được rỗng!");

            string tempPng = Path.Combine(Path.GetTempPath(), "tikz_" + Guid.NewGuid().ToString("N") + ".png");
            string encoded = Uri.EscapeDataString(tikzCode);
            string url = "https://quickchart.io/latex?formula=" + encoded + "&density=300";

            using (var client = new WebClient())
            {
                client.Headers.Add("User-Agent", "ToolVaBiPPT/1.0");
                await client.DownloadFileTaskAsync(new Uri(url), tempPng);
            }

            return tempPng;
        }

        private static string BuildFullTikzDocument(string body)
        {
            if (body.Contains(@"\documentclass")) return body;

            var sb = new StringBuilder();
            sb.AppendLine(@"\documentclass[tikz,border=4pt]{standalone}");
            sb.AppendLine(@"\usepackage{tikz}");
            sb.AppendLine(@"\usepackage{amsmath,amssymb}");
            sb.AppendLine(@"\usepackage{pgfplots}");
            sb.AppendLine(@"\pgfplotsset{compat=1.18}");
            sb.AppendLine(@"\begin{document}");
            sb.AppendLine(body);
            sb.AppendLine(@"\end{document}");
            return sb.ToString();
        }
    }
}
