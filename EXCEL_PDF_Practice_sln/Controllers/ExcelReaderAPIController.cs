using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EXCEL_PDF_Practice_ServiceLayer.Interface;
using System.Collections.Generic;
using AutoMapper;
using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.ResultModel;
using EXCEL_PDF_Practice_ParameterLayer.SlnModel.SearchModel;

namespace EXCEL_PDF_Practice_sln.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelReaderAPIController : ControllerBase
    {
        private readonly IWebHostEnvironment _iwebHostEnvironment;
        private readonly IExcelPdfPracticeService _excelPdfPracticeService;
        private readonly IMapper _mapper;

        public ExcelReaderAPIController(IWebHostEnvironment iwebHostEnvironment
            , IExcelPdfPracticeService excelPdfPracticeService
            , IMapper mapper)
        {
            _iwebHostEnvironment = iwebHostEnvironment;
            _excelPdfPracticeService = excelPdfPracticeService;
            _mapper = mapper;

        }

        [HttpGet("/GetExcelFromTemplateXlsx", Name = "GetExcelFromTemplateXlsx")]
        public IActionResult GetExcelFromTemplateXlsx()
        {
            // 構建文件的相對路徑
            var filePath = Path.Combine(_iwebHostEnvironment.ContentRootPath, "template", "FackData20240907.xlsx");

            if (!System.IO.File.Exists(filePath))
                return NotFound("模板檔案不存在");

            // 使用 MemoryStream 將檔案讀取到記憶體中
            byte[] fileBytes;
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var memoryStream = new MemoryStream())
                {
                    fileStream.CopyTo(memoryStream);
                    fileBytes = memoryStream.ToArray();
                }
            }

            using (var memoryStream = new MemoryStream(fileBytes))
            {
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

                        var modelList = new List<GetExcelFromTemplateXlsxContextSearchModel>();

                        // 假設第一行是標題行，跳過它
                        foreach (var row in rows.Skip(1)) 
                        {
                            var cells = row.Elements<Cell>().ToArray();

                            var model = new GetExcelFromTemplateXlsxContextSearchModel
                            {
                                OrderNumber = GetCellValue(cells[0], workbookPart),
                                Store = GetCellValue(cells[1], workbookPart),
                                ProductName = GetCellValue(cells[2], workbookPart),
                                Price = GetCellValue(cells[3], workbookPart),
                                OrderDate = GetCellValue(cells[4], workbookPart),
                                Quantity = GetCellValue(cells[5], workbookPart)
                            };

                            modelList.Add(model);
                        }

                        // 將 Excel 數據傳遞到服務層進行處理
                        var result = _excelPdfPracticeService.GetExcelFromTemplateXlsxContext(modelList);

                        if (result)
                            return Ok("模板檔案讀取並處理成功");
                        else
                            return StatusCode(500, "處理模板檔案時發生錯誤");
                    }
                }
            }

            return NotFound("讀取失敗");
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            var value = cell?.CellValue?.InnerText;

            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>()
                    .ElementAt(int.Parse(value))
                    .InnerText;
            }

            return value;
        }
    }
}
