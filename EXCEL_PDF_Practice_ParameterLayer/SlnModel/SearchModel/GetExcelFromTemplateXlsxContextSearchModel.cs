using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EXCEL_PDF_Practice_ParameterLayer.SlnModel.SearchModel
{
    public class GetExcelFromTemplateXlsxContextSearchModel
    {
        public string? OrderNumber { get; set; }
        public string? Store { get; set; }
        public string? ProductName { get; set; }
        public string? Price { get; set; }
        public string? OrderDate { get; set; }
        public string? Quantity { get; set; }

    }
}
