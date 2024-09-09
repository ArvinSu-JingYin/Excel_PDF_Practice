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
        /// <summary>Order Number</summary>
        public string? OrderNumber { get; set; }

        /// <summary>Store</summary>
        public string? Store { get; set; }

        /// <summary>Product Name</summary>
        public string? ProductName { get; set; }

        /// <summary>Price</summary>
        public string? Price { get; set; }

        /// <summary>Order Date</summary>
        public string? OrderDate { get; set; }

        /// <summary>Quantity</summary>
        public string? Quantity { get; set; }

    }
}
