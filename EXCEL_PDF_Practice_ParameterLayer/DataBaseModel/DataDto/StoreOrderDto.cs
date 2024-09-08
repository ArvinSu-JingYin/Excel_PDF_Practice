using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXCEL_PDF_Practice_ParameterLayer.DataBaseModel.DataDto
{
    public class StoreOrderDto
    {
        public string? OrderNumber { get; set; }
        public string? Store { get; set; }
        public string? ProductName { get; set; }
        public decimal Price { get; set; }
        public DateTime OrderDate { get; set; }
        public int? Quantity { get; set; }
    }
}
