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

        public string CusErroCodeHelper(string cusErroCode)
        {
            return CusErroMessageResource.ResourceManager.GetString(cusErroCode) ?? "未知錯誤";
        }

    }
}
