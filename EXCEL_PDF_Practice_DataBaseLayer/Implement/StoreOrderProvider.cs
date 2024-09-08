using DocumentFormat.OpenXml.Wordprocessing;
using EXCEL_PDF_Practice_DataBaseLayer.Interface;
using EXCEL_PDF_Practice_ParameterLayer.ServiceModel.DataModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Logging;
using System.Data;
using AutoMapper;
using DocumentFormat.OpenXml.InkML;
using EXCEL_PDF_Practice_ParameterLayer.DataBaseModel.DataDto;


namespace EXCEL_PDF_Practice_DataBaseLayer.Implement
{
    public class StoreOrderProvider : IStoreOrderProvider
    {
        private readonly string _dataBaseSettingProvider;
        private readonly ILogger<StoreOrderProvider> _logger;
        private readonly IMapper _mapper;

        public StoreOrderProvider(
            IDataBaseSettingProvider dataBaseSettingProvider,
            ILogger<StoreOrderProvider> logger,
            IMapper mapper)
        {
            _dataBaseSettingProvider = dataBaseSettingProvider.ConnectionString;
            _logger = logger;
            _mapper = mapper;
        }

        public string InsertStoreOrder(List<GetExcelFromTemplateXlsxContextDataModel> dataModel)
        {
            var dtoLists = _mapper.Map<List<StoreOrderDto>>(dataModel);
            StoreOrderDto dto = new StoreOrderDto();
            var valuesList = new List<string>();

            StringBuilder sql = new StringBuilder();
            DynamicParameters obj = new DynamicParameters();

            sql.Append("INSERT INTO [dbo].[PracticeProject_StoreOrder]" + Environment.NewLine);
            sql.Append("(" + Environment.NewLine);
            sql.Append("    OrderNumber," + Environment.NewLine);
            sql.Append("    Store," + Environment.NewLine);
            sql.Append("    ProductName," + Environment.NewLine);
            sql.Append("    Price," + Environment.NewLine);
            sql.Append("    Quantity," + Environment.NewLine);
            sql.Append("    OrderDate" + Environment.NewLine);
            sql.Append(")" + Environment.NewLine);
            sql.Append("VALUES" + Environment.NewLine);

            if (dtoLists != null && dtoLists.Count == 1)
            {
                sql.Append("(" + Environment.NewLine);
                sql.Append("    @OrderNumber," + Environment.NewLine);
                sql.Append("    @Store," + Environment.NewLine);
                sql.Append("    @ProductName," + Environment.NewLine);
                sql.Append("    @Price," + Environment.NewLine);
                sql.Append("    @Quantity," + Environment.NewLine);
                sql.Append("    @OrderDate," + Environment.NewLine);
                sql.Append(")" + Environment.NewLine);
            }
            else if (dtoLists != null && dtoLists.Count > 1)
            {
                for (int i = 0; i < dtoLists.Count; i++)
                {
                    valuesList.Add($"(@OrderNumber{i}, @Store{i}, @ProductName{i}, @Price{i}, @Quantity{i}, @OrderDate{i})");
                    obj.Add($"@OrderNumber{i}", dtoLists[i].OrderNumber, DbType.String);
                    obj.Add($"@Store{i}", dtoLists[i].Store, DbType.String);
                    obj.Add($"@ProductName{i}", dtoLists[i].ProductName, DbType.String);
                    obj.Add($"@Price{i}", dtoLists[i].Price, DbType.String);
                    obj.Add($"@Quantity{i}", dtoLists[i].Quantity, DbType.String);
                    obj.Add($"@OrderDate{i}", dtoLists[i].OrderDate, DbType.String);
                }

                sql.Append(string.Join(", ", valuesList));
            }

            using (var conn = new SqlConnection(_dataBaseSettingProvider))
            {
                conn.Open();

                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        var affectedRows = conn.Execute(sql.ToString(), obj, transaction: transaction);
                        transaction.Commit();
                        _logger.LogInformation($"[ InsertStoreOrder ] succeeded. Affected rows: {affectedRows}.");
                        return "OK";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        _logger.LogError($"Error occurred in [ InsertStoreOrder ] : {ex}");
                        return "Fail";
                    }

                }
            }
        }
    }
}
