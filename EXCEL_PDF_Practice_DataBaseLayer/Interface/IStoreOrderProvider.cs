using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.DataModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXCEL_PDF_Practice_DataBaseLayer.Interface
{
    public interface IStoreOrderProvider
    {
        public bool InsertStoreOrder(List<GetExcelFromTemplateXlsxContextDataModel> dataModel);

        public IEnumerable<dynamic> GetNonExistentStoreOrders(List<string> orderList);
    }
}
