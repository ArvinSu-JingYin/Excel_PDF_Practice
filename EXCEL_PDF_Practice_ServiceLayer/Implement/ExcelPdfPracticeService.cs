using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.ResultModel;
using EXCEL_PDF_Practice_ServiceLayer.Interface;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using AutoMapper;
using EXCEL_PDF_Practice_ParameterLayer.SlnModel.SearchModel;
using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.DataModel;
using EXCEL_PDF_Practice_DataBaseLayer.Interface;
using EXCEL_PDF_Practice_DataBaseLayer.Implement;


namespace EXCEL_PDF_Practice_ServiceLayer.Implement
{
    public class ExcelPdfPracticeService : IExcelPdfPracticeService
    {
        private readonly IStoreOrderProvider _storeOrderProvider;
        private readonly IMapper _mapper;

        public ExcelPdfPracticeService(IMapper mapper, IStoreOrderProvider storeOrderProvider) 
        {
            _mapper = mapper;
            _storeOrderProvider = storeOrderProvider;
        }

        public bool GetExcelFromTemplateXlsxContext(List<GetExcelFromTemplateXlsxContextSearchModel> TranTemplateRowList)
        {
            
            var mappings = _mapper.Map<List<GetExcelFromTemplateXlsxContextDataModel>>(TranTemplateRowList);

            string result = _storeOrderProvider.InsertStoreOrder(mappings);

            return true;
        }
    }
}
