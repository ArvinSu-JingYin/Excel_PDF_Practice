using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.ResultModel;
using EXCEL_PDF_Practice_ServiceLayer.Interface;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using AutoMapper;
using EXCEL_PDF_Practice_ParameterLayer;
using EXCEL_PDF_Practice_ParameterLayer.SlnModel.SearchModel;
using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.DataModel;
using EXCEL_PDF_Practice_DataBaseLayer.Interface;
using Microsoft.IdentityModel.Tokens;
using CommonHelper.Interface;
using Microsoft.Extensions.Logging;

namespace EXCEL_PDF_Practice_ServiceLayer.Implement
{
    public class ExcelPdfPracticeService : IExcelPdfPracticeService
    {
        private readonly IStoreOrderProvider _storeOrderProvider;
        private readonly ICusErroMessageHelper _cusErroMessageHelper;
        private readonly IMapper _mapper;
        private readonly ILogger<ExcelPdfPracticeService> _logger;

        public ExcelPdfPracticeService(IMapper mapper
            , IStoreOrderProvider storeOrderProvider
            , ICusErroMessageHelper cusErroMessageHelper
            , ILogger<ExcelPdfPracticeService> logger)
        {
            _mapper = mapper;
            _storeOrderProvider = storeOrderProvider;
            _cusErroMessageHelper = cusErroMessageHelper;
            _logger = logger;
        }

        public string GetExcelFromTemplateXlsxContext(List<GetExcelFromTemplateXlsxContextSearchModel> TranTemplateRowList)
        {
            var queryList = new List<string>();
            string result = string.Empty;

            try
            {
                #region 檢查是否有空欄位

                bool hasEmptyFields = TranTemplateRowList.Any(x => HasEmptyFields(x));

                if (hasEmptyFields)
                {
                    return _cusErroMessageHelper.CusErroCodeHelper("ColumsEmpty");
                }

                #endregion

                var mappings = _mapper.Map<List<GetExcelFromTemplateXlsxContextDataModel>>(TranTemplateRowList);

                //查資料庫內是否有相同訂單的資料
                foreach (var mapping in mappings.Where(x => !string.IsNullOrEmpty(x.OrderNumber)))
                {
                    queryList.Add(mapping.OrderNumber);
                }

                var query = _storeOrderProvider.GetNonExistentStoreOrders(queryList);

                //將資料寫入資料庫
                if (query.Count() == 0)
                {
                    result = _storeOrderProvider.InsertStoreOrder(mappings) ? "寫入成功"
                        : _cusErroMessageHelper.CusErroCodeHelper("InsertDataFail");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error occurred in [ GetExcelFromTemplateXlsxContext ] : {ex}");
                return _cusErroMessageHelper.CusErroCodeHelper("SystemErro");
            }

            return result;
        }

        private bool HasEmptyFields(GetExcelFromTemplateXlsxContextSearchModel row)
        {
            var propertiesToCheck = row.GetType().GetProperties();

            foreach (var prop in propertiesToCheck)
            {
                var propertyValue = prop.GetValue(row);

                if (propertyValue == null)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
