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

        /// <summary>
        /// Generate PDF file
        /// </summary>
        /// <returns>PDF document</returns>
        /// <info>Author: Arvin; Date: 2024/09/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/12  1.00    Arvin       Create Generate PDF file
        /// </history>
        public PdfDocument BuildPdfFromDataBaseTest()
        {
            try
            {
                // Create PDF document
                var cover = new PdfDocument();
                cover.Info.Title = "Created with PDFsharp";
                cover.Info.Subject = "Just a simple Hello-World program.";

                // Add a page
                var page = cover.AddPage();

                // Use XGraphics to draw image
                var gfx = XGraphics.FromPdfPage(page);
                var width = page.Width;
                var height = page.Height;

                // Fill background color
                gfx.DrawRectangle(XBrushes.SteelBlue, 0, 0, width, height);

                // Draw gray frame and white background rectangle
                double margin = 80;
                var grayPen = new XPen(XColors.Gray, 4);
                gfx.DrawRectangle(grayPen, XBrushes.White, margin, margin, width - 2 * margin, 150);

                // Load Chinese font
                var fontResolver = new FontResolver();
                fontResolver.GetFont("NotoSerifTC-Light.ttf");
                GlobalFontSettings.FontResolver = fontResolver;

                // Use custom font to draw text
                // Specify Chinese font
                var font = new XFont("NotoSerifTC-Light.ttf", 32, XFontStyle.Bold);
                gfx.DrawString("這是一本書。", font, XBrushes.DarkGray,
                    new XRect(margin, margin, width - 2 * margin, 150), XStringFormats.Center);

                return cover;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error generating PDF: {ex.ToString()}");
                throw;
            }
        }

        /// <summary>
        /// Generate PDF file from database
        /// </summary>
        /// <param name="num">Order number</param>
        /// <returns>PDF document</returns>
        /// <info>Author: Arvin; Date: 2024/09/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        public PdfDocument BuildPdfFromDataBase(string num)
        {
            // Create document
            var document = new Document();
            document.Info.Title = "Sample PDF created with MigraDocCore";
            document.Info.Author = "Your Name";

            // Add a section
            var section = document.AddSection();

            // Adjust page setup for centering effect
            section.PageSetup.LeftMargin = "2cm";
            section.PageSetup.RightMargin = "2cm";
            section.PageSetup.TopMargin = "2cm";
            section.PageSetup.BottomMargin = "2cm";

            // Add title
            var title = section.AddParagraph("Order Information");
            title.Format.Font.Size = 18;
            title.Format.Font.Bold = true;
            title.Format.SpaceAfter = 12;
            title.Format.Alignment = ParagraphAlignment.Center;

            // Add table, Field and Value in first and second row respectively
            var table = section.AddTable();
            table.Borders.Width = 0.75;

            // Set the table as a whole to center
            table.Format.Alignment = ParagraphAlignment.Center;

            // Set the column width of three columns
            var column1 = table.AddColumn("3cm");
            var column2 = table.AddColumn("3cm");
            var column3 = table.AddColumn("3cm");
            var column4 = table.AddColumn("3cm");
            var column5 = table.AddColumn("3cm");
            var column6 = table.AddColumn("3cm");

            // Add field names (Field), displayed in the first row
            AddFieldRowToTable(table, "Order Number", "Store", "Product Name", "Price", "Order Date", "Quantity");

            // Add field values (Value) corresponding to the field, displayed in the second row
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

            // Generate PDF file
            var pdfDocument = RenderPdf(document);
            return pdfDocument;
        }

        /// <summary>
        /// Add field row to table
        /// </summary>
        /// <param name="table">Table to add field row to</param>
        /// <param name="field1">Field 1</param>
        /// <param name="field2">Field 2</param>
        /// <param name="field3">Field 3</param>
        /// <param name="field4">Field 4</param>
        /// <param name="field5">Field 5</param>
        /// <param name="field6">Field 6</param>
        /// <returns>Table with field row added</returns>
        /// <info>Author: Arvin; Date: 2024/09/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  =========================================
        /// 01.  2024/09/12  1.00    Arvin       Create Add field row to table
        /// </history>
        private void AddFieldRowToTable(Table table, string field1, string field2, string field3, string field4, string field5, string field6)
        {

            var row = table.AddRow();

            // Set field name font
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

            // Bold field names
            foreach (Cell cell in row.Cells)
            {
                cell.Format.Font.Bold = true;
            }
        }

        /// <summary>
        /// Add value row to table
        /// </summary>
        /// <param name="table">Table to add value row to</param>
        /// <param name="value1">Value 1</param>
        /// <param name="value2">Value 2</param>
        /// <param name="value3">Value 3</param>
        /// <param name="value4">Value 4</param>
        /// <param name="value5">Value 5</param>
        /// <param name="value6">Value 6</param>
        /// <returns>Table with value row added</returns>
        /// <info>Author: Arvin; Date: 2024/09/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  =========================================
        /// 01.  2024/09/12  1.00    Arvin       Create Add value row to table
        /// </history> 
        private void AddValueRowToTable(Table table, string value1, string value2, string value3, string value4, string value5, string value6)
        {
            var row = table.AddRow();

            // Set value font
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

        /// <summary>
        /// Render PDF document
        /// </summary>
        /// <param name="document">Document to render</param>
        /// <returns>Rendered PDF document</returns>
        /// <info>Author: Arvin; Date: 2024/09/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  =========================================
        /// 01.  2024/09/12  1.00    Arvin       Create Render PDF document
        /// </history>
        private PdfDocument RenderPdf(Document document)
        {
            // Load Chinese font
            var fontResolver = new FontResolver();
            fontResolver.GetFont("NotoSerifTC-Light.ttf");
            GlobalFontSettings.FontResolver = fontResolver;

            // Use PdfDocumentRenderer to render MigraDoc document as PdfSharpCore's PdfDocument
            var pdfRenderer = new PdfDocumentRenderer(true);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();

            // Return the generated PdfDocument
            return pdfRenderer.PdfDocument;
        }
    }

    /// <summary>
    /// Font resolver
    /// </summary>
    /// <info>Author: Arvin; Date: 2024/09/12  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  =========================================
    /// 01.  2024/09/12  1.00    Arvin       Create Font resolver
    /// </history>
    internal class FontResolver : IFontResolver
    {
        /// <summary>Base font path</summary>
        private readonly string BaseFontPath = @"C:\Arvin\GitHub\Excel_PDF_Practice\Excel_PDF_Practice_BackEnd\EXCEL_PDF_Practice_sln\Source\Font\";

        /// <summary>Custom font name</summary>
        private string cusFontName = string.Empty;

        /// <summary>Default font name</summary>
        public string DefaultFontName => "TW-Kai-98_1.ttf";

        /// <summary>
        /// Get font
        /// </summary>
        /// <param name="faceName">Font name</param>
        /// <returns>Font</returns>
        /// <info>Author: Arvin; Date: 2024/09/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  =========================================
        /// 01.  2024/09/12  1.00    Arvin       Create Get font
        /// </history>
        public byte[] GetFont(string faceName)
        {
            cusFontName = faceName;

            return File.ReadAllBytes($"{BaseFontPath}{faceName}");
        }

        /// <summary>
        /// Resolve typeface
        /// </summary>
        /// <param name="familyName">Family name</param>
        /// <param name="isBold">Is bold</param>
        /// <param name="isItalic">Is italic</param>
        /// <returns>Font resolver info</returns>
        /// <info>Author: Arvin; Date: 2024/09/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  =========================================
        /// 01.  2024/09/12  1.00    Arvin       Create Resolve typeface
        /// </history>
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            return new FontResolverInfo(cusFontName, isBold, isItalic);
        }
    }
}
