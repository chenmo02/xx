using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Printing;

namespace WpfApp1.Services
{
    // ═══════════════════════════════════════
    // 发票文件项模型
    // ═══════════════════════════════════════
    public class InvoiceFileItem
    {
        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "";
        public string FileType { get; set; } = ""; // PDF / OFD / IMG
        public long FileSize { get; set; }
        public BitmapSource? PreviewImage { get; set; }
        public int PageCount { get; set; } = 1;
        public int SelectedPage { get; set; } = 0;
        public double RotationAngle { get; set; } = 0;
        public Rect CropRect { get; set; } = Rect.Empty; // 裁剪区域
        public bool IsPrinted { get; set; } = false;

        public string DisplayName => $"{FileName} ({FileType})";
        public string FileSizeText => FileSize < 1024 * 1024
            ? $"{FileSize / 1024.0:F1} KB"
            : $"{FileSize / (1024.0 * 1024.0):F1} MB";
    }

    // ═══════════════════════════════════════
    // 打印模板配置
    // ═══════════════════════════════════════
    public class PrintTemplate
    {
        public string Name { get; set; } = "默认模板";
        public string PaperMode { get; set; } = "A4"; // A4 / Invoice (发票专用纸)
        public int LayoutCount { get; set; } = 1; // 每页发票数: 1, 2, 4
        public double MarginTop { get; set; } = 10;
        public double MarginBottom { get; set; } = 10;
        public double MarginLeft { get; set; } = 10;
        public double MarginRight { get; set; } = 10;
        public double OffsetX { get; set; } = 0; // 套打偏移X (mm)
        public double OffsetY { get; set; } = 0; // 套打偏移Y (mm)
        public string PrintQuality { get; set; } = "标准"; // 草稿/标准/高画质
    }

    // ═══════════════════════════════════════
    // 发票打印核心服务
    // ═══════════════════════════════════════
    public class InvoicePrintService
    {
        private readonly string _templatePath;
        private readonly string _historyPath;

        public InvoicePrintService()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            _templatePath = Path.Combine(baseDir, "invoice_templates.json");
            _historyPath = Path.Combine(baseDir, "invoice_print_history.json");
        }

        // ── 文件导入 ──

        public static readonly string[] SupportedExtensions = { ".pdf", ".ofd", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff" };

        public List<InvoiceFileItem> ImportFiles(string[] filePaths)
        {
            var items = new List<InvoiceFileItem>();
            foreach (var path in filePaths)
            {
                if (!File.Exists(path)) continue;
                var ext = Path.GetExtension(path).ToLowerInvariant();
                if (!SupportedExtensions.Contains(ext)) continue;

                var fi = new FileInfo(path);
                var item = new InvoiceFileItem
                {
                    FilePath = path,
                    FileName = fi.Name,
                    FileType = GetFileType(ext),
                    FileSize = fi.Length,
                };

                // 尝试加载预览图
                item.PreviewImage = LoadPreviewImage(path, ext);
                items.Add(item);
            }
            return items;
        }

        public List<InvoiceFileItem> ImportFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath)) return new List<InvoiceFileItem>();
            var files = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)
                .Where(f => SupportedExtensions.Contains(Path.GetExtension(f).ToLowerInvariant()))
                .ToArray();
            return ImportFiles(files);
        }

        private static string GetFileType(string ext)
        {
            return ext switch
            {
                ".pdf" => "PDF",
                ".ofd" => "OFD",
                _ => "IMG"
            };
        }

        // ── PDF/图片 预览渲染 ──

        public static BitmapSource? LoadPreviewImage(string filePath, string ext)
        {
            try
            {
                if (ext == ".pdf")
                {
                    return RenderPdfPage(filePath, 0, 200); // 200 DPI 预览
                }
                else if (ext == ".ofd")
                {
                    // OFD 暂不支持直接渲染，返回占位图
                    return null;
                }
                else
                {
                    // 图片文件直接加载
                    var bi = new BitmapImage();
                    bi.BeginInit();
                    bi.UriSource = new Uri(filePath, UriKind.Absolute);
                    bi.CacheOption = BitmapCacheOption.OnLoad;
                    bi.DecodePixelWidth = 800; // 限制预览尺寸
                    bi.EndInit();
                    bi.Freeze();
                    return bi;
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 使用 Windows.Data.Pdf (WinRT) 渲染 PDF 页面为 BitmapSource
        /// </summary>
        public static BitmapSource? RenderPdfPage(string pdfPath, int pageIndex, int dpi)
        {
            try
            {
                // 使用 WinRT API: Windows.Data.Pdf
                var file = Windows.Storage.StorageFile.GetFileFromPathAsync(pdfPath).AsTask().Result;
                var pdfDoc = Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(file).AsTask().Result;

                if (pageIndex >= (int)pdfDoc.PageCount) return null;

                using var page = pdfDoc.GetPage((uint)pageIndex);
                var options = new Windows.Data.Pdf.PdfPageRenderOptions();
                // 计算渲染尺寸 (基于 DPI)
                double scale = dpi / 96.0;
                options.DestinationWidth = (uint)(page.Size.Width * scale);
                options.DestinationHeight = (uint)(page.Size.Height * scale);

                using var stream = new Windows.Storage.Streams.InMemoryRandomAccessStream();
                page.RenderToStreamAsync(stream, options).AsTask().Wait();

                // 转换为 WPF BitmapSource
                stream.Seek(0);
                var decoder = BitmapDecoder.Create(
                    stream.AsStreamForRead(),
                    BitmapCreateOptions.PreservePixelFormat,
                    BitmapCacheOption.OnLoad);

                var frame = decoder.Frames[0];
                frame.Freeze();
                return frame;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 获取 PDF 总页数
        /// </summary>
        public static int GetPdfPageCount(string pdfPath)
        {
            try
            {
                var file = Windows.Storage.StorageFile.GetFileFromPathAsync(pdfPath).AsTask().Result;
                var pdfDoc = Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(file).AsTask().Result;
                return (int)pdfDoc.PageCount;
            }
            catch { return 0; }
        }

        // ── 图像变换 ──

        public static BitmapSource RotateImage(BitmapSource source, double angle)
        {
            var tb = new TransformedBitmap(source, new RotateTransform(angle));
            tb.Freeze();
            return tb;
        }

        public static BitmapSource CropImage(BitmapSource source, Int32Rect cropRect)
        {
            var cb = new CroppedBitmap(source, cropRect);
            cb.Freeze();
            return cb;
        }

        public static BitmapSource ScaleImage(BitmapSource source, double scale)
        {
            var tb = new TransformedBitmap(source, new ScaleTransform(scale, scale));
            tb.Freeze();
            return tb;
        }

        // ── 高质量打印渲染 ──

        public static BitmapSource? GetPrintImage(InvoiceFileItem item, int dpi = 300)
        {
            BitmapSource? img = null;
            var ext = Path.GetExtension(item.FilePath).ToLowerInvariant();

            if (ext == ".pdf")
            {
                img = RenderPdfPage(item.FilePath, item.SelectedPage, dpi);
            }
            else
            {
                // 高分辨率加载原图
                var bi = new BitmapImage();
                bi.BeginInit();
                bi.UriSource = new Uri(item.FilePath, UriKind.Absolute);
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.EndInit();
                bi.Freeze();
                img = bi;
            }

            if (img == null) return null;

            // 应用旋转
            if (item.RotationAngle != 0)
                img = RotateImage(img, item.RotationAngle);

            // 应用裁剪
            if (item.CropRect != Rect.Empty)
            {
                var cr = new Int32Rect(
                    (int)item.CropRect.X, (int)item.CropRect.Y,
                    (int)item.CropRect.Width, (int)item.CropRect.Height);
                if (cr.Width > 0 && cr.Height > 0)
                    img = CropImage(img, cr);
            }

            return img;
        }

        // ── 排版引擎：生成打印页面 ──

        public static List<DrawingVisual> BuildPrintPages(
            List<InvoiceFileItem> items, PrintTemplate template, Size pageSize)
        {
            var pages = new List<DrawingVisual>();
            int perPage = template.LayoutCount;

            // mm → WPF 单位 (1 inch = 96 WPF units, 1 inch = 25.4 mm)
            double mmToWpf = 96.0 / 25.4;
            double ml = template.MarginLeft * mmToWpf;
            double mr = template.MarginRight * mmToWpf;
            double mt = template.MarginTop * mmToWpf;
            double mb = template.MarginBottom * mmToWpf;
            double ox = template.OffsetX * mmToWpf;
            double oy = template.OffsetY * mmToWpf;

            double contentW = pageSize.Width - ml - mr;
            double contentH = pageSize.Height - mt - mb;

            // 计算每个发票块的尺寸
            int cols, rows;
            switch (perPage)
            {
                case 2: cols = 1; rows = 2; break;
                case 4: cols = 2; rows = 2; break;
                default: cols = 1; rows = 1; break;
            }

            double cellW = contentW / cols;
            double cellH = contentH / rows;
            double gap = 4 * mmToWpf; // 4mm 间距

            // 分页处理
            for (int i = 0; i < items.Count; i += perPage)
            {
                var dv = new DrawingVisual();
                using (var dc = dv.RenderOpen())
                {
                    // 白色背景
                    dc.DrawRectangle(Brushes.White, null, new Rect(0, 0, pageSize.Width, pageSize.Height));

                    for (int j = 0; j < perPage && (i + j) < items.Count; j++)
                    {
                        var item = items[i + j];
                        var img = GetPrintImage(item, GetDpiFromQuality(template.PrintQuality));
                        if (img == null) continue;

                        int col = j % cols;
                        int row = j / cols;

                        double x = ml + col * cellW + ox;
                        double y = mt + row * cellH + oy;
                        double w = cellW - (cols > 1 ? gap : 0);
                        double h = cellH - (rows > 1 ? gap : 0);

                        // 保持宽高比缩放
                        double scaleX = w / img.PixelWidth;
                        double scaleY = h / img.PixelHeight;
                        double scale = Math.Min(scaleX, scaleY);

                        double drawW = img.PixelWidth * scale;
                        double drawH = img.PixelHeight * scale;

                        // 居中
                        double drawX = x + (w - drawW) / 2;
                        double drawY = y + (h - drawH) / 2;

                        dc.DrawImage(img, new Rect(drawX, drawY, drawW, drawH));
                    }
                }
                pages.Add(dv);
            }

            return pages;
        }

        private static int GetDpiFromQuality(string quality)
        {
            return quality switch
            {
                "草稿" => 150,
                "高画质" => 600,
                _ => 300
            };
        }

        // ── 打印执行 ──

        public static bool PrintPages(List<DrawingVisual> pages, int copies, PrintDialog? dialog = null)
        {
            if (pages.Count == 0) return false;

            if (dialog == null)
            {
                dialog = new PrintDialog();
                if (dialog.ShowDialog() != true) return false;
            }

            for (int c = 0; c < copies; c++)
            {
                if (pages.Count == 1)
                {
                    dialog.PrintVisual(pages[0], "CC发票打印");
                }
                else
                {
                    var doc = new FixedDocument();
                    var pageSize = new Size(dialog.PrintableAreaWidth, dialog.PrintableAreaHeight);

                    foreach (var visual in pages)
                    {
                        var fp = new FixedPage { Width = pageSize.Width, Height = pageSize.Height };

                        // 将 DrawingVisual 转为 Image
                        int w = (int)pageSize.Width;
                        int h = (int)pageSize.Height;
                        var rtb = new RenderTargetBitmap(w * 2, h * 2, 192, 192, PixelFormats.Pbgra32);
                        rtb.Render(visual);
                        rtb.Freeze();

                        var img = new System.Windows.Controls.Image
                        {
                            Source = rtb,
                            Width = pageSize.Width,
                            Height = pageSize.Height
                        };

                        fp.Children.Add(img);
                        var pc = new PageContent();
                        ((System.Windows.Markup.IAddChild)pc).AddChild(fp);
                        doc.Pages.Add(pc);
                    }

                    var writer = System.Printing.PrintQueue.CreateXpsDocumentWriter(dialog.PrintQueue);
                    writer.Write(doc);
                }
            }

            return true;
        }

        // ── 模板管理 ──

        public List<PrintTemplate> LoadTemplates()
        {
            try
            {
                if (!File.Exists(_templatePath)) return GetDefaultTemplates();
                var json = File.ReadAllText(_templatePath);
                return JsonSerializer.Deserialize<List<PrintTemplate>>(json) ?? GetDefaultTemplates();
            }
            catch { return GetDefaultTemplates(); }
        }

        public void SaveTemplates(List<PrintTemplate> templates)
        {
            var json = JsonSerializer.Serialize(templates, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_templatePath, json);
        }

        public static List<PrintTemplate> GetDefaultTemplates()
        {
            return new List<PrintTemplate>
            {
                new() { Name = "A4 - 1张/页", PaperMode = "A4", LayoutCount = 1 },
                new() { Name = "A4 - 2张/页", PaperMode = "A4", LayoutCount = 2 },
                new() { Name = "A4 - 4张/页", PaperMode = "A4", LayoutCount = 4 },
                new() { Name = "发票专用纸", PaperMode = "Invoice", LayoutCount = 1,
                         MarginTop = 5, MarginBottom = 5, MarginLeft = 5, MarginRight = 5 }
            };
        }

        // ── 打印历史 ──

        public void RecordPrintHistory(List<InvoiceFileItem> items)
        {
            try
            {
                var history = new List<Dictionary<string, string>>();
                if (File.Exists(_historyPath))
                {
                    var existing = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(
                        File.ReadAllText(_historyPath));
                    if (existing != null) history = existing;
                }

                foreach (var item in items)
                {
                    history.Add(new Dictionary<string, string>
                    {
                        ["FileName"] = item.FileName,
                        ["FilePath"] = item.FilePath,
                        ["PrintTime"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        ["FileType"] = item.FileType
                    });
                }

                // 只保留最近 500 条
                if (history.Count > 500) history = history.TakeLast(500).ToList();

                File.WriteAllText(_historyPath,
                    JsonSerializer.Serialize(history, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { }
        }
    }
}
