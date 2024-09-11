using AutoMapper;
using EXCEL_PDF_Practice_ParameterLayer.DataBaseModel.DataDto;
using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.DataModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace EXCEL_PDF_Practice_DataBaseLayer.Mapping
{
    public class DataBaseMappingProfile : Profile
    {
        public DataBaseMappingProfile() {

            //in
            CreateMap<GetExcelFromTemplateXlsxContextDataModel, StoreOrderDto>();


            //out
            CreateMap< dynamic, StoreOrderDto>(); 
        }
    }
}
