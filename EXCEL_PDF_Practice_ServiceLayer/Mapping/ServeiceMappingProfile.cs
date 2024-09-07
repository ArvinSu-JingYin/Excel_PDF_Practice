using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using EXCEL_PDF_Practice_ParameterLayer;
using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.DataModel;
using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.ResultModel;
using EXCEL_PDF_Practice_ParameterLayer.SlnModel.SearchModel;

namespace EXCEL_PDF_Practice_ServiceLayer.Mapping
{
    public class ServeiceMappingProfile : Profile
    {
        public ServeiceMappingProfile()
        {
            //Data into ServiceLayer
            CreateMap<GetExcelFromTemplateXlsxContextSearchModel, GetExcelFromTemplateXlsxContextDataModel>();

            //Data out to ServiceLayer
            CreateMap<GetExcelFromTemplateXlsxContextDataModel, GetExcelFromTemplateXlsxContextResultModel>();

        }
    }
}
