using CommonHelper.Interface;
using EXCEL_PDF_Practice_ParameterLayer;
using EXCEL_PDF_Practice_ParameterLayer.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonHelper.Implement
{
    public class CusErroMessageHelper : ICusErroMessageHelper
    {
        /// <summary>
        /// Get Custom Error Message
        /// </summary>
        /// <param name="cusErroCode">The custom error code</param>
        /// <returns>The custom error message</returns>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create CusErroCodeHelper
        /// </history>
        public string CusErroCodeHelper(string cusErroCode)
        {
            return CusErroMessageResource.ResourceManager.GetString(cusErroCode) ?? "未知錯誤";
        }

    }
}
