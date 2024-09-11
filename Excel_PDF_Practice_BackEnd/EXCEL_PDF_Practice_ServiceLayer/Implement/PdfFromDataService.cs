using AutoMapper;
using EXCEL_PDF_Practice_ServiceLayer.Interface;
using Microsoft.Extensions.Logging;
using PdfSharpCore.Drawing;
using PdfSharpCore.Fonts;
using PdfSharpCore.Pdf;
using MigraDocCore.Rendering;
using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.Tables;
using EXCEL_PDF_Practice_DataBaseLayer;
using EXCEL_PDF_Practice_DataBaseLayer.Interface;
using EXCEL_PDF_Practice_DataBaseLayer.Implement;
using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.DataModel;
using EXCEL_PDF_Practice_ParameterLayer.DataBaseModel.ResultDto;

namespace EXCEL_PDF_Practice_ServiceLayer.Implement
{
    public class PdfFromDataService : IPdfFromDataService
    {
        private readonly ILogger<PdfFromDataService> _logger;
        private readonly IMapper _mapper;
        private readonly IStoreOrderProvider _storeOrderProvider;

        public PdfFromDataService(
            ILogger<PdfFromDataService> logger,
            IMapper mapper,
            IStoreOrderProvider storeOrderProvider)
        {
            _logger = logger;
            _mapper = mapper;
            _storeOrderProvider = storeOrderProvider;
        }

        public PdfDocument BuildPdfFromDataBaseTest()
        {
            try
            {
                // 建立 PDF 文檔
                var cover = new PdfDocument();
                cover.Info.Title = "Created with PDFsharp";
                cover.Info.Subject = "Just a simple Hello-World program.";

                // 新增一頁
                var page = cover.AddPage();

                // 使用 XGraphics 繪製圖像
                var gfx = XGraphics.FromPdfPage(page);
                var width = page.Width;
                var height = page.Height;

                // 背景填色
                gfx.DrawRectangle(XBrushes.SteelBlue, 0, 0, width, height);

                // 繪製灰框白底矩形
                double margin = 80;
                var grayPen = new XPen(XColors.Gray, 4);
                gfx.DrawRectangle(grayPen, XBrushes.White, margin, margin, width - 2 * margin, 150);

                // 載入中文字體
                var fontResolver = new FontResolver();
                fontResolver.GetFont("NotoSerifTC-Light.ttf");
                GlobalFontSettings.FontResolver = fontResolver;

                // 使用自定義字體繪製文字
                var font = new XFont("NotoSerifTC-Light.ttf", 32, XFontStyle.Bold); // 指定中文字體
                gfx.DrawString("這是一本書。", font, XBrushes.DarkGray,
                    new XRect(margin, margin, width - 2 * margin, 150), XStringFormats.Center);

                return cover; // 返回 PDF 文件
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error generating PDF: {ex.ToString()}");
                throw;
            }
        }

        public PdfDocument BuildPdfFromDataBase(string num)
        {
            // 建立文檔
            var document = new Document();
            document.Info.Title = "Sample PDF created with MigraDocCore";
            document.Info.Author = "Your Name";

            // 添加一個章節
            var section = document.AddSection();

            // 調整頁面設置以達到置中效果
            section.PageSetup.LeftMargin = "2cm";  // 左邊距 2cm
            section.PageSetup.RightMargin = "2cm"; // 右邊距 2cm
            section.PageSetup.TopMargin = "2cm";   // 頁頂邊距
            section.PageSetup.BottomMargin = "2cm"; // 頁底邊距

            // 添加標題
            var title = section.AddParagraph("Order Information");
            title.Format.Font.Size = 18;
            title.Format.Font.Bold = true;
            title.Format.SpaceAfter = 12;
            title.Format.Alignment = ParagraphAlignment.Center; // 標題居中

            // 添加表格，Field 和 Value 分別在第一行和第二行
            var table = section.AddTable();
            table.Borders.Width = 0.75;

            // 將表格整體設置為居中
            table.Format.Alignment = ParagraphAlignment.Center; // 表格居中

            // 設定三列的欄寬
            var column1 = table.AddColumn("3cm"); // 第一列
            var column2 = table.AddColumn("3cm"); // 第二列
            var column3 = table.AddColumn("3cm"); // 第三列
            var column4 = table.AddColumn("3cm"); // 第三列
            var column5 = table.AddColumn("3cm"); // 第三列
            var column6 = table.AddColumn("3cm"); // 第三列

            // 添加欄位名稱（Field），顯示在第一行
            AddFieldRowToTable(table, "Order Number", "Store", "Product Name", "Price", "Order Date", "Quantity");

            // 添加欄位對應的值（Value），顯示在第二行
            var queries = _storeOrderProvider.GetStoreOrders(num);

            for (int i = 0; i < queries.Count(); i++)
            {
                StoreOrderResultDto query = queries.ElementAt(i);
                var tranQuery = _mapper.Map<PdfFromDataBaseModel>(query);

                AddValueRowToTable(
                    table,
                    tranQuery.OrderNumber,
                    tranQuery.Store,
                    tranQuery.ProductName,
                    tranQuery.Price,
                    tranQuery.OrderDate,
                    tranQuery.Quantity
                    );
            }

            // 生成 PDF 文件
            var pdfDocument = RenderPdf(document);
            return pdfDocument;
        }

        // 添加欄位名稱的行（第一行）
        private void AddFieldRowToTable(Table table, string field1, string field2, string field3, string field4, string field5, string field6)
        {

            var row = table.AddRow();

            // 設置欄位名稱字體
            foreach (Cell cell in row.Cells)
            {
                var paragraph = cell.AddParagraph();
                paragraph.Format.Font.Name = "NotoSerifTC-Light";
                paragraph.Format.Font.Size = 12;
                paragraph.Format.Font.Bold = true;
                paragraph.Format.Alignment = ParagraphAlignment.Center;
            }

            row.Cells[0].AddParagraph(field1).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[1].AddParagraph(field2).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].AddParagraph(field3).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[3].AddParagraph(field4).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[4].AddParagraph(field5).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[5].AddParagraph(field6).Format.Alignment = ParagraphAlignment.Center;

            // 加粗欄位名稱
            foreach (Cell cell in row.Cells)
            {
                cell.Format.Font.Bold = true;
            }
        }

        // 添加欄位值的行（第二行）
        private void AddValueRowToTable(Table table, string value1, string value2, string value3, string value4, string value5, string value6)
        {
            var row = table.AddRow();

            // 設置欄位值字體
            foreach (Cell cell in row.Cells)
            {
                var paragraph = cell.AddParagraph();
                paragraph.Format.Font.Name = "NotoSerifTC-Light";
                paragraph.Format.Font.Size = 10;
                paragraph.Format.Alignment = ParagraphAlignment.Center;
            }

            row.Cells[0].AddParagraph(value1).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[1].AddParagraph(value2).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].AddParagraph(value3).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[3].AddParagraph(value4).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[4].AddParagraph(value5).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[5].AddParagraph(value6).Format.Alignment = ParagraphAlignment.Center;
        }

        private PdfDocument RenderPdf(Document document)
        {
            // 載入中文字體
            var fontResolver = new FontResolver();
            fontResolver.GetFont("NotoSerifTC-Light.ttf");
            GlobalFontSettings.FontResolver = fontResolver;

            // 使用 PdfDocumentRenderer 將 MigraDoc 文檔渲染為 PdfSharpCore 的 PdfDocument
            var pdfRenderer = new PdfDocumentRenderer(true);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();

            // 返回生成的 PdfDocument
            return pdfRenderer.PdfDocument;
        }
    }

    // 自定義字體解析器
    internal class FontResolver : IFontResolver
    {
        private readonly string BaseFontPath = @"C:\Arvin\GitHub\Excel_PDF_Practice\Excel_PDF_Practice_BackEnd\EXCEL_PDF_Practice_sln\Source\Font\";
        private string cusFontName = string.Empty;

        public string DefaultFontName => "TW-Kai-98_1.ttf";

        public byte[] GetFont(string faceName)
        {
            cusFontName = faceName;

            return File.ReadAllBytes($"{BaseFontPath}{faceName}");
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            return new FontResolverInfo(cusFontName, isBold, isItalic);
        }
    }
}
