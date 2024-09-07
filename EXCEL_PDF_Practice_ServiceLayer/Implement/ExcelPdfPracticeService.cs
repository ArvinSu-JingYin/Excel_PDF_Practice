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


namespace EXCEL_PDF_Practice_ServiceLayer.Implement
{
    public class ExcelPdfPracticeService : IExcelPdfPracticeService
    {
        private readonly IMapper _mapper;

        public ExcelPdfPracticeService(IMapper mapper) 
        {
            _mapper = mapper;  
        }

        public bool GetExcelFromTemplateXlsxContext(IEnumerable<Row> row)
        {
            // 將 DTO 映射為實體
            var userEntity = _mapper.Map<GetExcelFromTemplateXlsxContextResultModel>(row);

            return true;
        }
    }
}
