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

        /// <summary>
        /// Get Excel from template Xlsx
        /// </summary>
        /// <returns>Excel file content</returns>
        /// <info>Author: Arvin; Date: 2024/09/09  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create Get Excel from template Xlsx
        /// </history>
        public string GetExcelFromTemplateXlsxContext(List<GetExcelFromTemplateXlsxContextSearchModel> TranTemplateRowList)
        {
            var queryList = new List<string>();
            string result = string.Empty;

            try
            {
                #region Check for empty fields

                bool hasEmptyFields = TranTemplateRowList.Any(x => HasEmptyFields(x));

                if (hasEmptyFields)
                {
                    return _cusErroMessageHelper.CusErroCodeHelper("ColumsEmpty");
                }

                #endregion

                var mappings = _mapper.Map<List<GetExcelFromTemplateXlsxContextDataModel>>(TranTemplateRowList);

                // Check if there is data for the same order in the database
                foreach (var mapping in mappings.Where(x => !string.IsNullOrEmpty(x.OrderNumber)))
                {
                    queryList.Add(mapping.OrderNumber);
                }

                var query = _storeOrderProvider.GetNonExistentStoreOrders(queryList);

                // Write data to the database
                if (query.Count() == 0)
                {
                    result = _storeOrderProvider.InsertStoreOrder(mappings) ? "Write successfully"
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


        /// <summary>
        /// Checks if the model has any empty fields
        /// </summary>
        /// <param name="row">The model to check</param>
        /// <returns>True if the model has any empty fields, false otherwise</returns>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create HasEmptyFields
        /// </history>
        /// <info>Author: Arvin; Date: 2024/09/09  </info>
        private bool HasEmptyFields(GetExcelFromTemplateXlsxContextSearchModel row)
        {
            // Get all properties of the model
            var propertiesToCheck = row.GetType().GetProperties();

            // Iterate over each property and check its value for null
            foreach (var prop in propertiesToCheck)
            {
                var propertyValue = prop.GetValue(row);

                // If any empty field is found, return true immediately
                if (propertyValue == null)
                {
                    return true;
                }
            }

            // If all properties are non-empty, return false
            return false;
        }
    }
}
