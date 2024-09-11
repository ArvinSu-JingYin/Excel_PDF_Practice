using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using EXCEL_PDF_Practice_ParameterLayer;
using EXCEL_PDF_Practice_ParameterLayer.DataBaseModel.ResultDto;
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
            CreateMap<GetExcelFromTemplateXlsxContextSearchModel, GetExcelFromTemplateXlsxContextDataModel>()
                .ForMember(x => x.Price,
                           y => y.MapFrom(z => ParsePrice(z.Price?? "0")))
                .ForMember(x => x.Quantity,
                           y => y.MapFrom(z => ParseQuantity(z.Quantity ?? "0")))
                .ForMember(x => x.OrderDate,
                           y => y.MapFrom( z => ParseOrderDate(z.OrderDate ?? "0")));

            CreateMap< StoreOrderResultDto, PdfFromDataBaseModel>()
                .ForMember(x => x.Price,
                           y => y.MapFrom(z => z.Price.ToString()))
                .ForMember(x => x.Quantity,
                           y => y.MapFrom(z => z.Quantity.ToString()))
                .ForMember(x => x.OrderDate,
                           y => y.MapFrom(z => z.OrderDate.ToString())); ;

            //Data out to ServiceLayer
            CreateMap<GetExcelFromTemplateXlsxContextDataModel, GetExcelFromTemplateXlsxContextResultModel>();

        }

        private decimal ParsePrice(string price)
        {
            return decimal.TryParse(price, out var result) ? result : 0;
        }

        private int ParseQuantity(string quantity)
        {
            return int.TryParse(quantity, out var result) ? result : 0;
        }

        private DateTime ParseOrderDate(string orderDate) 
        {
            return DateTime.FromOADate(
                   double.TryParse(orderDate, out var result)? result : 0
                );
        }
    }
}
