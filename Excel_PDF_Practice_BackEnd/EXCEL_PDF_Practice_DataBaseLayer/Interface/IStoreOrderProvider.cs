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
        /// <summary>
        /// Retrieves a list of store orders based on the provided order number.
        /// </summary>
        /// <param name="num">The order number to filter by</param>
        /// <returns>An IEnumerable of dynamic objects representing the store orders</returns>
        /// <info>Author: Arvin; Date: 2024/09/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/12  1.00    Arvin       Create GetStoreOrders
        /// </history>
        public IEnumerable<dynamic> GetStoreOrders(string num);

        /// <summary>
        /// Inserts a list of store orders into the database.
        /// </summary>
        /// <param name="dataModel">The list of store orders to insert</param>
        /// <returns>True if the insertion was successful, false otherwise</returns>
        /// <info>Author: Arvin; Date: 2024/09/09  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create InsertStoreOrder
        /// </history>
        public bool InsertStoreOrder(List<GetExcelFromTemplateXlsxContextDataModel> dataModel);

        /// <summary>
        /// Retrieves a list of store orders that do not exist in the database.
        /// </summary>
        /// <param name="orderList">The list of order numbers to check</param>
        /// <returns>A list of order numbers that do not exist in the database</returns>
        /// <info>Author: Arvin; Date: 2024/09/09  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create GetNonExistentStoreOrders
        public IEnumerable<dynamic> GetNonExistentStoreOrders(List<string> orderList);
    }
}
