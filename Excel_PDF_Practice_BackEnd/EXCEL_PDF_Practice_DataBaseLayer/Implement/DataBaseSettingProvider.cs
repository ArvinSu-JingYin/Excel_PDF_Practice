using EXCEL_PDF_Practice_DataBaseLayer.Interface;

namespace EXCEL_PDF_Practice_DataBaseLayer.Implement
{
    public class DataBaseSettingProvider : IDataBaseSettingProvider
    {
        public string ConnectionString { get; private set; }

        /// <summary>
        /// Constructor for DataBaseSettingProvider
        /// </summary>
        /// <param name="connectionString">The connection string to be used for database operations</param>
        /// <info>Author: Arvin; Date: 2024/09/09  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create DataBaseSettingProvider
        /// </history>
        public DataBaseSettingProvider(string connectionString) {

            ConnectionString = connectionString;
        }
    }
}
