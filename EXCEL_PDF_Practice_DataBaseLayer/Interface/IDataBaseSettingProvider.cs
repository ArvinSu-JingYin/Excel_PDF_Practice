using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXCEL_PDF_Practice_DataBaseLayer.Interface
{
    public interface IDataBaseSettingProvider
    {
        /// <summary>
        /// Gets the connection string for the database.
        /// </summary>
        /// <value>The connection string for the database.</value>
        /// <info>Author: Arvin; Date: 2024/09/09  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create ConnectionString
        /// </history>
        public string ConnectionString { get;}
    }
}
