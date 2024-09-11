using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXCEL_PDF_Practice_ParameterLayer.ServiceModel.DataModel
{
    public class PdfFromDataBaseModel
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
