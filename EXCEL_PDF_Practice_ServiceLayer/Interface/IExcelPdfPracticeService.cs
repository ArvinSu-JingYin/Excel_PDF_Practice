using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.ResultModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EXCEL_PDF_Practice_ServiceLayer.Interface
{
    public interface IExcelPdfPracticeService
    {
        public bool GetExcelFromTemplateXlsxContext(IEnumerable<Row> row);
    }
}
