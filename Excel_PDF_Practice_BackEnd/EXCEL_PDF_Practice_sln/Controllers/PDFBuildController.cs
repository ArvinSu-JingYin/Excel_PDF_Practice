using Microsoft.AspNetCore.Mvc;
using EXCEL_PDF_Practice_ServiceLayer.Interface;
using PdfSharpCore.Pdf;

namespace EXCEL_PDF_Practice_sln.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class PDFBuildController : ControllerBase
    {
        private readonly IPdfFromDataService _pdfFromDataService;

        public PDFBuildController(IPdfFromDataService pdfFromDataService)
        {
            _pdfFromDataService = pdfFromDataService;
        }

        [HttpPost("/GetPdfFromDataBase", Name= "GetPdfFromDataBase")]
        public IActionResult GetPdfFromDataBase(string num = null)
        {
            PdfDocument document = new PdfDocument();

            document = _pdfFromDataService.BuildPdfFromDataBase(num);

            // 將 PDF 文件存儲在內存流中
            using (MemoryStream stream = new MemoryStream())
            {
                document.Save(stream, false);
                stream.Position = 0;

                // 以 application/pdf 的格式返回文件供下載
                return File(stream.ToArray(), "application/pdf", "測試文件.pdf");
            }
        }
    }
}
