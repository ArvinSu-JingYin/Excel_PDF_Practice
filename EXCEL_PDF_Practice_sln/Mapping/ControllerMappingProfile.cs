using AutoMapper;
using EXCEL_PDF_Practice_ParameterLayer.SlnModel.SearchModel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EXCEL_PDF_Practice_sln.Mapping
{
    public class ControllerMappingProfile : Profile
    {
        public ControllerMappingProfile() {
            //Controller to Service
            CreateMap<IEnumerable<Row>, GetExcelFromTemplateXlsxContextSearchModel > ();
        }
    }
}
