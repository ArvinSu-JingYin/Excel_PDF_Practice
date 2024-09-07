using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EXCEL_PDF_Practice_ServiceLayer.Interface;
using System.Collections.Generic;

namespace EXCEL_PDF_Practice_sln.Controllers
{
    public class ExcelReaderAPIController : ControllerBase
    {
        private readonly IWebHostEnvironment _iwebHostEnvironment;
        private readonly IExcelPdfPracticeService _excelPdfPracticeService;

        public ExcelReaderAPIController(IWebHostEnvironment iwebHostEnvironment, IExcelPdfPracticeService excelPdfPracticeService)
        {
            _iwebHostEnvironment = iwebHostEnvironment;
            _excelPdfPracticeService = excelPdfPracticeService;
        }

        public string Index()
        {
            return "Y";
        }

        public IActionResult GetExcelFromTemplateXlsx()
        {
            // 構建文件的相對路徑
            var filePath = Path.Combine(_iwebHostEnvironment.ContentRootPath, "template", "FackData20240907.xlsx");

            if (!System.IO.File.Exists(filePath))
                return NotFound("模板檔案不存在");

            using (var spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                var sheets = workbookPart.Workbook.Sheets;

                if (sheets != null)
                {
                    // 假設只讀取第一個工作表
                    var firstSheet = (Sheet)sheets.First();
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id);

                    // 獲取 Worksheet 並讀取數據
                    var rows = worksheetPart.Worksheet.Descendants<Row>();

                    // 將 Excel 數據傳遞到服務層進行處理
                    var result = _excelPdfPracticeService.GetExcelFromTemplateXlsxContext(rows);

                    if (result)
                        return Ok("模板檔案讀取並處理成功");
                    else
                        return StatusCode(500, "處理模板檔案時發生錯誤");

                }
            }

            return NotFound("讀取失敗");
        }

        
    }
}
