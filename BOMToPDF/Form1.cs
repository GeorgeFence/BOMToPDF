using ClosedXML.Excel;
using HtmlAgilityPack;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using Microsoft.Playwright;
using Microsoft.Web.WebView2.Core;
using SkiaSharp;
using Svg.Skia;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using HorizontalAlignment = iText.Layout.Properties.HorizontalAlignment;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using Image = iText.Layout.Element.Image;
using Path = System.IO.Path;
using Rectangle = System.Drawing.Rectangle;
using Size = System.Drawing.Size;
using SKSvg = Svg.Skia.SKSvg;

namespace BOMToPDF
{
    public partial class Form1 : Form
    {
        private string ExcelPath;
        private readonly HttpClient httpClient = new HttpClient();

        public Form1()
        {
            InitializeComponent();
            httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void LoadBomAndFillCombos()
        {
            using OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv";

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            string path = ofd.FileName;
            ExcelPath = path;
            var headers = new List<string>();
            var table = new DataTable();

            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheet(1);

                foreach (var cell in ws.Row(1).CellsUsed())
                {
                    string header = cell.GetString();
                    headers.Add(header);
                    table.Columns.Add(header);
                }

                foreach (var row in ws.RowsUsed().Skip(1).Take(20))
                {
                    table.Rows.Add(
                        headers.Select((h, i) => row.Cell(i + 1).GetString()).ToArray()
                    );
                }
            }

            dataGridView1.DataSource = table;

            comboID.DataSource = headers.ToList();

            var headers2 = headers.ToList();
            headers2.Insert(0, "Name from LSCS");
            comboName.DataSource = headers2;
        }

        private void btnSelectBOM_Click(object sender, EventArgs e)
        {
            LoadBomAndFillCombos();
        }

        private async Task<string?> GetLcscDescriptionAsync(string url)
        {
            try
            {
                string html = await httpClient.GetStringAsync(url);
                var doc = new HtmlDocument();
                doc.LoadHtml(html);

                var trNode = doc.DocumentNode.SelectSingleNode("//tr[td[contains(normalize-space(.), 'Description')]]");
                if (trNode == null) return null;

                var spanNode = trNode.SelectSingleNode(".//span[contains(@class,'major2--text')]");
                if (spanNode == null)
                {
                    var td = trNode.SelectSingleNode("./td[2]");
                    return td?.InnerText.Trim();
                }

                return spanNode.InnerText.Trim();
            }
            catch (HttpRequestException)
            {
                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private async void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(ExcelPath))
            {
                MessageBox.Show("Load BOM first!");
                return;
            }

            bool useScraper = (comboName.SelectedItem?.ToString() ?? "") == "Name from LSCS";
            string colLcsc = comboID.SelectedItem?.ToString() ?? "";
            string colName = comboName.SelectedItem?.ToString() ?? "";

            try
            {
                var items = await ParseBomItemsAsync(ExcelPath, colLcsc, colName, useScraper,true);
                CreateLabelsPdf(items);
                MessageBox.Show("PDF generated.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private async Task<List<BomItem>> ParseBomItemsAsync(
    string excelPath,
    string colLcsc,
    string colName,
    bool useLcscScraper,
    bool useJlcFootprint)
        {
            List<BomItem> items = new();

            using var wb = new XLWorkbook(excelPath);
            var ws = wb.Worksheet(1);

            var headers = ws.Row(1)
                .CellsUsed()
                .Select((c, i) => new { Name = c.GetString(), Index = i + 1 })
                .ToList();

            int idxOrThrow(string name)
            {
                var found = headers.FirstOrDefault(h => h.Name == name);
                if (found == null) throw new Exception($"Header '{name}' not found in Excel.");
                return found.Index;
            }

            int iLcsc = idxOrThrow(colLcsc);
            int iName = -1;
            if (!useLcscScraper)
                iName = idxOrThrow(colName);

            foreach (var row in ws.RowsUsed().Skip(1))
            {
                string lcsc = row.Cell(iLcsc).GetString().Trim();
                if (string.IsNullOrWhiteSpace(lcsc))
                    continue;

                string name = "";
                byte[]? footprint = null;

                if (!useLcscScraper)
                    name = row.Cell(iName).GetString().Trim();
                else
                {
                    try
                    {
                        var scraped = await GetLcscDescriptionAsync($"https://www.lcsc.com/product-detail/{lcsc}.html");
                        name = scraped ?? "HTTP ERROR";
                    }
                    catch { name = "HTTP ERROR"; }
                }

                if (useJlcFootprint)
                {
                    try
                    {
                        using Bitmap bmp = await GetJlcPcbFootprintAsync(lcsc,Int32.Parse(numTimeout.Value.ToString()) * 1000);
                        using var ms = new MemoryStream();
                        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        footprint = ms.ToArray();
                    }
                    catch
                    {
                        footprint = GenerateGenericFootprint();
                    }
                }
                else
                {
                    footprint = GenerateGenericFootprint();
                }

                items.Add(new BomItem
                {
                    LcscId = lcsc,
                    Name = name,
                    FootprintBitmap = footprint
                });
            }

            return items;
        }

        public async Task<Bitmap> GetJlcPcbFootprintAsync(string jlcCode,int timeout = 2000)
        {
            string url =
                $"https://jlcpcb.com/user-center/lcsvg/svg.html?code={jlcCode}&detail=1";

            var tcsLoaded = new TaskCompletionSource<bool>();

            void Handler(object? s, CoreWebView2NavigationCompletedEventArgs e)
            {
                webView21.NavigationCompleted -= Handler;
                tcsLoaded.SetResult(true);
            }

            webView21.NavigationCompleted += Handler;
            webView21.Source = new Uri(url);

            await tcsLoaded.Task;
            await Task.Delay(timeout);

            string js = @"
(() => {
    const g = document.querySelector('g[c_origin]');
    if (!g) return null;

    const rect = g.getBoundingClientRect();

    return {
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height,
        dpr: window.devicePixelRatio
    };
})();
";

            string result = await webView21.ExecuteScriptAsync(js);

            if (result == "null")
                throw new Exception("SVG group <g c_origin> not found");

            var box = JsonSerializer.Deserialize<SvgBox>(result)
                      ?? throw new Exception("Invalid JS result");

            using var ms = new MemoryStream();
            await webView21.CoreWebView2.CapturePreviewAsync(
                CoreWebView2CapturePreviewImageFormat.Png, ms);

            using var fullBmp = new Bitmap(ms);

            float scale = box.dpr;

            Rectangle crop = new Rectangle(
                (int)(box.x * scale),
                (int)(box.y * scale),
                (int)(box.width * scale),
                (int)(box.height * scale)
            );

            Bitmap cropped = fullBmp.Clone(crop, fullBmp.PixelFormat);
            return cropped;
        }

        private void CreateLabelsPdf(List<BomItem> items)
        {
            using SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "PDF file (*.pdf)|*.pdf";
            sfd.DefaultExt = "pdf";
            sfd.FileName = "LCSC_Labels.pdf";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            string outputPath = sfd.FileName;

            float labelWidth = MmToPt(34f);
            float labelHeight = MmToPt(12.5f);

            using (var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                PdfWriter writer = new PdfWriter(fs);
                PdfDocument pdf = new PdfDocument(writer);
                Document doc = new Document(pdf, PageSize.A4);
                doc.SetMargins(20, 20, 20, 20);

                float usableWidth = PageSize.A4.GetWidth() - doc.GetLeftMargin() - doc.GetRightMargin();
                float usableHeight = PageSize.A4.GetHeight() - doc.GetTopMargin() - doc.GetBottomMargin();

                int cols = Math.Max(1, (int)(usableWidth / labelWidth));
                int rows = Math.Max(1, (int)(usableHeight / labelHeight));

                Table table = new Table(cols);
                table.SetWidth(UnitValue.CreatePointValue(cols * labelWidth));

                int labelIndex = 0;

                foreach (var item in items)
                {
                    Cell cell = new Cell()
                        .SetWidth(labelWidth)
                        .SetHeight(labelHeight)
                        .SetPadding(2)
                        .SetBorder(new SolidBorder(0.5f));

                    float imgColWidth = labelWidth * 0.45f;
                    float textColWidth = labelWidth - imgColWidth - 2;

                    Table innerTable = new Table(new float[] { imgColWidth, textColWidth })
                        .SetWidth(UnitValue.CreatePointValue(labelWidth - 2))
                        .SetBorder(Border.NO_BORDER);

                    Cell imgCell = new Cell().SetBorder(Border.NO_BORDER).SetVerticalAlignment(VerticalAlignment.MIDDLE);
                    if (item.FootprintBitmap != null && item.FootprintBitmap.Length > 0)
                    {
                        var imgData = iText.IO.Image.ImageDataFactory.Create(item.FootprintBitmap);
                        Image img = new Image(imgData)
                            .ScaleToFit(imgColWidth, labelHeight - 4)
                            .SetHorizontalAlignment(HorizontalAlignment.RIGHT);
                        imgCell.Add(img);
                    }
                    innerTable.AddCell(imgCell);

                    Cell textCell = new Cell().SetBorder(Border.NO_BORDER).SetVerticalAlignment(VerticalAlignment.MIDDLE);
                    textCell.Add(new Paragraph(item.LcscId).SetFontSize(6).SetMargin(0));
                    textCell.Add(new Paragraph(item.Name).SetFontSize(5).SetMargin(0));
                    innerTable.AddCell(textCell);

                    cell.Add(innerTable);
                    table.AddCell(cell);
                    labelIndex++;

                    if (labelIndex % (cols * rows) == 0)
                    {
                        doc.Add(table);
                        doc.Add(new AreaBreak());
                        table = new Table(cols);
                        table.SetWidth(UnitValue.CreatePointValue(cols * labelWidth));
                    }
                }

                if (table.GetNumberOfRows() > 0)
                    doc.Add(table);

                doc.Close();
            }
        }


        private byte[] GenerateGenericFootprint()
        {
            using Bitmap bmp = new Bitmap(200, 80);
            using Graphics g = Graphics.FromImage(bmp);
            g.Clear(Color.White);
            g.DrawRectangle(Pens.Black, 5, 5, 190, 70);
            g.DrawString("ERROR", new Font("Arial", 28, FontStyle.Bold), Brushes.Black, new PointF(60, 15));
            using MemoryStream ms = new MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            return ms.ToArray();
        }

        private float MmToPt(float mm)
        {
            return mm * 72f / 25.4f;
        } 
    }

    class BomItem
    {
        public string LcscId { get; set; } = "";
        public string Name { get; set; } = "";

        public byte[] FootprintBitmap { get; set; } = Array.Empty<byte>();
    }
    public class SvgBox
    {
        public float x { get; set; }
        public float y { get; set; }
        public float width { get; set; }
        public float height { get; set; }
        public float dpr { get; set; }
    }
}
