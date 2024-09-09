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
        /// <summary>
        /// Get Excel from template Xlsx
        /// </summary>
        /// <param name="TranTemplateRows">The list of models to write to the Excel file</param>
        /// <returns>A string indicating the result of the operation</returns>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create GetExcelFromTemplateXlsxContext
        /// </history>
        /// <info>Author: Arvin; Date: 2024/09/09  </info>
        public string GetExcelFromTemplateXlsxContext(List<GetExcelFromTemplateXlsxContextSearchModel> TranTemplateRows);
    }
}
