using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EXCEL_PDF_Practice_ServiceLayer.Interface;
using AutoMapper;
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

        /// <summary>
        /// Get Excel from template Xlsx
        /// </summary>
        /// <returns>Excel file content</returns>
        /// <info>Author: Arvin; Date: 2024/09/09  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create Get Excel from template Xlsx
        /// </history>
        [HttpGet("/GetExcelFromTemplateXlsx", Name = "GetExcelFromTemplateXlsx")]
        public IActionResult GetExcelFromTemplateXlsx()
        {
            // Construct the relative path of the file
            var filePath = Path.Combine(_iwebHostEnvironment.ContentRootPath, "template", "FackData20240907.xlsx");

            if (!System.IO.File.Exists(filePath))
                return NotFound("Template file does not exist");

            // Use MemoryStream to read the file into memory
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
                        // Assuming only the first worksheet is read
                        var firstSheet = (Sheet)sheets.First();
                        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id);

                        // Get Worksheet and read data
                        var rows = worksheetPart.Worksheet.Descendants<Row>();

                        var modelList = new List<GetExcelFromTemplateXlsxContextSearchModel>();

                        // Assuming the first row is the header row, skip it
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

                        // Pass Excel data to the service layer for processing
                        var result = _excelPdfPracticeService.GetExcelFromTemplateXlsxContext(modelList);

                        return Ok(result);
                    }
                }
            }

            return NotFound("Failed to read");
        }

        /// <summary>
        /// Get the value of a cell
        /// </summary>
        /// <param name="cell">Cell</param>
        /// <param name="workbookPart">Workbook part</param>
        /// <returns>Cell value</returns>
        /// <info>Author: Arvin; Date: 2024/09/09  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// 01.  2024/09/09  1.00    Arvin       Create Get the value of a cell
        /// </history>
        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            // Define variable value and initialize it with the cell value
            var value = cell?.CellValue?.InnerText;

            // If the cell type is SharedString, get the corresponding value from the shared string table
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                // Get the shared string table from the workbook part's shared string table part
                var sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;

                // Get the corresponding shared string item from the shared string table
                var sharedStringItem = sharedStringTable
                    .Elements<SharedStringItem>()
                    .ElementAt(int.Parse(value));

                // Get the content from the shared string item and return
                return sharedStringItem.InnerText;
            }

            // If the cell type is not SharedString, directly return the cell value
            return value;
        }
    }
}
