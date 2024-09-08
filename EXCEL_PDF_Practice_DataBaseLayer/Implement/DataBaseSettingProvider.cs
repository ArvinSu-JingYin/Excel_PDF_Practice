using EXCEL_PDF_Practice_DataBaseLayer.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXCEL_PDF_Practice_DataBaseLayer.Implement
{
    public class DataBaseSettingProvider : IDataBaseSettingProvider
    {
        public string ConnectionString { get; private set; }

        public DataBaseSettingProvider(string connectionString) {

            ConnectionString = connectionString;
        }
    }
}
