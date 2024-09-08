using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.ResultModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using EXCEL_PDF_Practice_ParameterLayer.SlnModel.SearchModel;

namespace EXCEL_PDF_Practice_ServiceLayer.Interface
{
    public interface IExcelPdfPracticeService
    {
        
        public bool GetExcelFromTemplateXlsxContext(List<GetExcelFromTemplateXlsxContextSearchModel> TranTemplateRows);
    }
}
