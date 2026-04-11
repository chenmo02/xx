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
    // 鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺?
    // 鍙戠エ鏂囦欢椤规ā鍨?
    // 鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺?
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
        public Rect CropRect { get; set; } = Rect.Empty; // 瑁佸壀鍖哄煙
        public bool IsPrinted { get; set; } = false;

        public string DisplayName => $"{FileName} ({FileType})";
        public string FileSizeText => FileSize < 1024 * 1024
            ? $"{FileSize / 1024.0:F1} KB"
            : $"{FileSize / (1024.0 * 1024.0):F1} MB";
    }

    // 鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺?
    // 鎵撳嵃妯℃澘閰嶇疆
    // 鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺?
    public class PrintTemplate
    {
        public string Name { get; set; } = "榛樿妯℃澘";
        public string PaperMode { get; set; } = "A4"; // A4 / Invoice (鍙戠エ涓撶敤绾?
        public int LayoutCount { get; set; } = 1; // 姣忛〉鍙戠エ鏁? 1, 2, 4
        public double MarginTop { get; set; } = 0;
        public double MarginBottom { get; set; } = 0;
        public double MarginLeft { get; set; } = 0;
        public double MarginRight { get; set; } = 0;
        public double OffsetX { get; set; } = 0; // 濂楁墦鍋忕ЩX (mm)
        public double OffsetY { get; set; } = 0; // 濂楁墦鍋忕ЩY (mm)
        public string PrintQuality { get; set; } = "鏍囧噯"; // 鑽夌/鏍囧噯/楂樼敾璐?
    }

    // 鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺?
    // 鍙戠エ鎵撳嵃鏍稿績鏈嶅姟
    // 鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺愨晲鈺?
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

        public class PrintLayoutContext
        {
            public PrintQueue? PrintQueue { get; set; }
            public PrintTicket PrintTicket { get; set; } = new();
            public Size MediaSize { get; set; }
            public Point ContentOrigin { get; set; }
            public Size ContentSize { get; set; }
        }

        // 鈹€鈹€ 鏂囦欢瀵煎叆 鈹€鈹€

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

                // 灏濊瘯鍔犺浇棰勮鍥?
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

        // 鈹€鈹€ PDF/鍥剧墖 棰勮娓叉煋 鈹€鈹€

        public static BitmapSource? LoadPreviewImage(string filePath, string ext)
        {
            try
            {
                if (ext == ".pdf")
                {
                    return RenderPdfPage(filePath, 0, 200); // 200 DPI 棰勮
                }
                else if (ext == ".ofd")
                {
                    // OFD 鏆備笉鏀寔鐩存帴娓叉煋锛岃繑鍥炲崰浣嶅浘
                    return null;
                }
                else
                {
                    // 鍥剧墖鏂囦欢鐩存帴鍔犺浇
                    var bi = new BitmapImage();
                    bi.BeginInit();
                    bi.UriSource = new Uri(filePath, UriKind.Absolute);
                    bi.CacheOption = BitmapCacheOption.OnLoad;
                    bi.DecodePixelWidth = 800; // 闄愬埗棰勮灏哄
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
        /// 浣跨敤 Windows.Data.Pdf (WinRT) 娓叉煋 PDF 椤甸潰涓?BitmapSource
        /// </summary>
        public static BitmapSource? RenderPdfPage(string pdfPath, int pageIndex, int dpi)
        {
            try
            {
                // 浣跨敤 WinRT API: Windows.Data.Pdf
                var file = Windows.Storage.StorageFile.GetFileFromPathAsync(pdfPath).AsTask().Result;
                var pdfDoc = Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(file).AsTask().Result;

                if (pageIndex >= (int)pdfDoc.PageCount) return null;

                using var page = pdfDoc.GetPage((uint)pageIndex);
                var options = new Windows.Data.Pdf.PdfPageRenderOptions();
                // 璁＄畻娓叉煋灏哄 (鍩轰簬 DPI)
                double scale = dpi / 96.0;
                options.DestinationWidth = (uint)(page.Size.Width * scale);
                options.DestinationHeight = (uint)(page.Size.Height * scale);

                using var stream = new Windows.Storage.Streams.InMemoryRandomAccessStream();
                page.RenderToStreamAsync(stream, options).AsTask().Wait();

                // 杞崲涓?WPF BitmapSource
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
        /// 鑾峰彇 PDF 鎬婚〉鏁?
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

        // 鈹€鈹€ 鍥惧儚鍙樻崲 鈹€鈹€

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

        // 鈹€鈹€ 楂樿川閲忔墦鍗版覆鏌?鈹€鈹€

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
                // 楂樺垎杈ㄧ巼鍔犺浇鍘熷浘
                var bi = new BitmapImage();
                bi.BeginInit();
                bi.UriSource = new Uri(item.FilePath, UriKind.Absolute);
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.EndInit();
                bi.Freeze();
                img = bi;
            }

            if (img == null) return null;

            // 搴旂敤鏃嬭浆
            if (item.RotationAngle != 0)
                img = RotateImage(img, item.RotationAngle);

            // 搴旂敤瑁佸壀
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

        // 鈹€鈹€ 鎺掔増寮曟搸锛氱敓鎴愭墦鍗伴〉闈?鈹€鈹€

        public static List<DrawingVisual> BuildPrintPages(
            List<InvoiceFileItem> items, PrintTemplate template, Size pageSize)
        {
            var pages = new List<DrawingVisual>();
            int perPage = template.LayoutCount;

            // mm 鈫?WPF 鍗曚綅 (1 inch = 96 WPF units, 1 inch = 25.4 mm)
            double mmToWpf = 96.0 / 25.4;
            double ml = template.MarginLeft * mmToWpf;
            double mr = template.MarginRight * mmToWpf;
            double mt = template.MarginTop * mmToWpf;
            double mb = template.MarginBottom * mmToWpf;
            double ox = template.OffsetX * mmToWpf;
            double oy = template.OffsetY * mmToWpf;

            double contentW = pageSize.Width - ml - mr;
            double contentH = pageSize.Height - mt - mb;

            // 璁＄畻姣忎釜鍙戠エ鍧楃殑灏哄
            int cols, rows;
            switch (perPage)
            {
                case 2: cols = 1; rows = 2; break;
                case 4: cols = 2; rows = 2; break;
                default: cols = 1; rows = 1; break;
            }

            double gap = 4 * mmToWpf; // 4mm 闂磋窛
            double totalGapW = gap * Math.Max(0, cols - 1);
            double totalGapH = gap * Math.Max(0, rows - 1);
            double cellW = Math.Max(0, (contentW - totalGapW) / cols);
            double cellH = Math.Max(0, (contentH - totalGapH) / rows);

            // 鍒嗛〉澶勭悊
            for (int i = 0; i < items.Count; i += perPage)
            {
                var dv = new DrawingVisual();
                using (var dc = dv.RenderOpen())
                {
                    // 鐧借壊鑳屾櫙
                    dc.DrawRectangle(Brushes.White, null, new Rect(0, 0, pageSize.Width, pageSize.Height));

                    for (int j = 0; j < perPage && (i + j) < items.Count; j++)
                    {
                        var item = items[i + j];
                        var img = GetPrintImage(item, GetDpiFromQuality(template.PrintQuality));
                        if (img == null) continue;

                        int col = j % cols;
                        int row = j / cols;

                        double x = ml + col * (cellW + gap) + ox;
                        double y = mt + row * (cellH + gap) + oy;
                        double w = cellW;
                        double h = cellH;

                        // 淇濇寔瀹介珮姣旂缉鏀?
                        double scaleX = w / img.PixelWidth;
                        double scaleY = h / img.PixelHeight;
                        double scale = Math.Min(scaleX, scaleY);

                        double drawW = img.PixelWidth * scale;
                        double drawH = img.PixelHeight * scale;

                        // 灞呬腑
                        double drawX = x + (w - drawW) / 2;
                        double drawY = y + (h - drawH) / 2;

                        dc.DrawImage(img, new Rect(drawX, drawY, drawW, drawH));
                    }
                }
                pages.Add(dv);
            }

            return pages;
        }

        public static PrintLayoutContext CreatePrintLayoutContext(PrintDialog dialog, PrintTemplate template, bool isLandscape)
        {
            if (dialog.PrintQueue == null)
                throw new InvalidOperationException("未找到可用的打印机。");

            var requestedMediaSize = GetRequestedPaperSize(template, isLandscape);
            var ticket = dialog.PrintTicket ?? new PrintTicket();
            ticket.PageOrientation = isLandscape ? PageOrientation.Landscape : PageOrientation.Portrait;
            ticket.PageMediaSize = template.PaperMode == "Invoice"
                ? new PageMediaSize(requestedMediaSize.Width, requestedMediaSize.Height)
                : new PageMediaSize(PageMediaSizeName.ISOA4);

            PrintTicket effectiveTicket;
            try
            {
                effectiveTicket = dialog.PrintQueue
                    .MergeAndValidatePrintTicket(dialog.PrintQueue.DefaultPrintTicket, ticket)
                    .ValidatedPrintTicket;
            }
            catch
            {
                effectiveTicket = ticket;
            }

            effectiveTicket.PageOrientation = isLandscape ? PageOrientation.Landscape : PageOrientation.Portrait;
            dialog.PrintTicket = effectiveTicket;

            var mediaSize = NormalizeOrientation(GetMediaSizeFromTicket(effectiveTicket, requestedMediaSize), isLandscape);
            var capabilities = dialog.PrintQueue.GetPrintCapabilities(effectiveTicket);
            var imageableArea = capabilities.PageImageableArea;

            var contentOrigin = imageableArea == null
                ? new Point(0, 0)
                : new Point(Math.Max(0, imageableArea.OriginWidth), Math.Max(0, imageableArea.OriginHeight));

            var contentSize = imageableArea == null
                ? new Size(dialog.PrintableAreaWidth, dialog.PrintableAreaHeight)
                : new Size(imageableArea.ExtentWidth, imageableArea.ExtentHeight);

            contentSize = NormalizeOrientation(contentSize, isLandscape);
            if (contentSize.Width <= 0 || contentSize.Height <= 0)
            {
                contentSize = NormalizeOrientation(
                    new Size(Math.Max(1, dialog.PrintableAreaWidth), Math.Max(1, dialog.PrintableAreaHeight)),
                    isLandscape);
            }

            if (contentSize.Width <= 0 || contentSize.Height <= 0)
                contentSize = mediaSize;

            return new PrintLayoutContext
            {
                PrintQueue = dialog.PrintQueue,
                PrintTicket = effectiveTicket,
                MediaSize = mediaSize,
                ContentOrigin = contentOrigin,
                ContentSize = contentSize
            };
        }

        public static FixedDocument BuildFixedDocument(List<DrawingVisual> pages, PrintLayoutContext context)
        {
            var doc = new FixedDocument();
            doc.DocumentPaginator.PageSize = context.MediaSize;

            foreach (var visual in pages)
            {
                var fp = new FixedPage
                {
                    Width = context.MediaSize.Width,
                    Height = context.MediaSize.Height
                };

                int pixelWidth = Math.Max(1, (int)Math.Ceiling(context.ContentSize.Width * 2));
                int pixelHeight = Math.Max(1, (int)Math.Ceiling(context.ContentSize.Height * 2));

                var rtb = new RenderTargetBitmap(pixelWidth, pixelHeight, 192, 192, PixelFormats.Pbgra32);
                rtb.Render(visual);
                rtb.Freeze();

                var img = new System.Windows.Controls.Image
                {
                    Source = rtb,
                    Width = context.ContentSize.Width,
                    Height = context.ContentSize.Height,
                    Stretch = Stretch.Fill
                };

                FixedPage.SetLeft(img, context.ContentOrigin.X);
                FixedPage.SetTop(img, context.ContentOrigin.Y);
                fp.Children.Add(img);

                var pc = new PageContent();
                ((System.Windows.Markup.IAddChild)pc).AddChild(fp);
                doc.Pages.Add(pc);
            }

            return doc;
        }

        private static int GetDpiFromQuality(string quality)
        {
            return quality switch
            {
                "草稿" => 150,
                "高质量" => 600,
                _ => 300
            };
        }

        // 鈹€鈹€ 鎵撳嵃鎵ц 鈹€鈹€

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
                    dialog.PrintVisual(pages[0], "CC鍙戠エ鎵撳嵃");
                }
                else
                {
                    var doc = new FixedDocument();
                    var pageSize = new Size(dialog.PrintableAreaWidth, dialog.PrintableAreaHeight);

                    foreach (var visual in pages)
                    {
                        var fp = new FixedPage { Width = pageSize.Width, Height = pageSize.Height };

                        // 灏?DrawingVisual 杞负 Image
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

        public static bool PrintPages(List<DrawingVisual> pages, int copies, PrintLayoutContext context)
        {
            if (pages.Count == 0) return false;
            if (context.PrintQueue == null) return false;

            var doc = BuildFixedDocument(pages, context);
            var writer = PrintQueue.CreateXpsDocumentWriter(context.PrintQueue);

            for (int c = 0; c < copies; c++)
            {
                writer.Write(doc.DocumentPaginator, context.PrintTicket);
            }

            return true;
        }

        private static Size GetRequestedPaperSize(PrintTemplate template, bool isLandscape)
        {
            double mmToWpf = 96.0 / 25.4;
            double width = template.PaperMode == "Invoice" ? 241 * mmToWpf : 210 * mmToWpf;
            double height = template.PaperMode == "Invoice" ? 140 * mmToWpf : 297 * mmToWpf;
            return NormalizeOrientation(new Size(width, height), isLandscape);
        }

        private static Size GetMediaSizeFromTicket(PrintTicket ticket, Size fallback)
        {
            var media = ticket.PageMediaSize;
            if (media?.Width is double width && width > 0 &&
                media.Height is double height && height > 0)
            {
                return new Size(width, height);
            }

            return fallback;
        }

        private static Size NormalizeOrientation(Size size, bool isLandscape)
        {
            if (size.Width <= 0 || size.Height <= 0)
                return size;

            if (isLandscape && size.Width < size.Height)
                return new Size(size.Height, size.Width);

            if (!isLandscape && size.Width > size.Height)
                return new Size(size.Height, size.Width);

            return size;
        }

        // 鈹€鈹€ 妯℃澘绠＄悊 鈹€鈹€

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
                new() { Name = "A4 - 1页", PaperMode = "A4", LayoutCount = 1 },
                new() { Name = "A4 - 2页", PaperMode = "A4", LayoutCount = 2 },
                new() { Name = "A4 - 4页", PaperMode = "A4", LayoutCount = 4 },
                new() { Name = "发票专用纸", PaperMode = "Invoice", LayoutCount = 1,
                         MarginTop = 0, MarginBottom = 0, MarginLeft = 0, MarginRight = 0 }
            };
        }

        // 鈹€鈹€ 鎵撳嵃鍘嗗彶 鈹€鈹€

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

                // 鍙繚鐣欐渶杩?500 鏉?
                if (history.Count > 500) history = history.TakeLast(500).ToList();

                File.WriteAllText(_historyPath,
                    JsonSerializer.Serialize(history, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { }
        }
    }
}

