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
using EXCEL_PDF_Practice_ParameterLayer.DataBaseModel.ResultDto;
using CommonHelper.Interface;


namespace EXCEL_PDF_Practice_DataBaseLayer.Implement
{
    public class StoreOrderProvider : IStoreOrderProvider
    {
        private readonly string _dataBaseSettingProvider;
        private readonly ILogger<StoreOrderProvider> _logger;
        private readonly IMapper _mapper;
        private readonly ICusErroMessageHelper _messageHelper;

        public StoreOrderProvider(
            IDataBaseSettingProvider dataBaseSettingProvider,
            ILogger<StoreOrderProvider> logger,
            IMapper mapper,
            ICusErroMessageHelper cusErroMessageHelper)
        {
            _dataBaseSettingProvider = dataBaseSettingProvider.ConnectionString;
            _logger = logger;
            _mapper = mapper;
            _messageHelper = cusErroMessageHelper;
        }

        public IEnumerable<dynamic> GetStoreOrders(string num) 
        { 
            StringBuilder sql = new StringBuilder();
            DynamicParameters obj = new DynamicParameters();

            sql.Append("SELECT" + Environment.NewLine); 
            sql.Append("    OrderNumber," + Environment.NewLine);
            sql.Append("    Store," + Environment.NewLine);
            sql.Append("    ProductName," + Environment.NewLine);
            sql.Append("    Price," + Environment.NewLine);
            sql.Append("    OrderDate," + Environment.NewLine);
            sql.Append("    Quantity" + Environment.NewLine);
            sql.Append("FROM [PracticeProject_StoreOrder]");
            sql.Append("WHERE 1 = 1");
            sql.Append("    AND OrderNumber = @OrderNumber");

            obj.Add("@OrderNumber", num, DbType.String);

            using (var conn = new SqlConnection(_dataBaseSettingProvider))
            {
                conn.Open();
                try
                {
                    return conn.Query(sql.ToString(), obj);
                }
                catch (Exception ex)
                {
                    var result = new List<dynamic>();

                    _logger.LogError($"Error occurred in [ GetStoreOrders ] : {ex}");

                    result.Add(_messageHelper.CusErroCodeHelper("DataBaseErro"));

                    return result;
                }
            }
        }

        /// <summary>
        /// Get Non Existent Store Orders
        /// </summary>
        /// <param name="orderList">The list of order numbers to check</param>
        /// <returns>A list of order numbers that do not exist in the database</returns>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create GetNonExistentStoreOrders
        /// </history>
        /// <info>Author: Arvin; Date: 2024/09/09  </info>
        public IEnumerable<dynamic> GetNonExistentStoreOrders(List<string> orderList)
        {
            var sqlList = new List<string>();

            StringBuilder sql = new StringBuilder();
            DynamicParameters obj = new DynamicParameters();

            sql.Append("WITH TempData AS" + Environment.NewLine);
            sql.Append("(" + Environment.NewLine);
            sql.Append("    SELECT OrderNumber" + Environment.NewLine);
            sql.Append("    FROM (  VALUES" + Environment.NewLine);

            if (orderList != null && orderList.Count == 1)
            {
                sql.Append("(@orderList)" + Environment.NewLine);
                obj.Add($"@orderList", orderList, DbType.String);
            }
            else
            {
                for (int i = 0; i < orderList.Count; i++)
                {
                    sqlList.Add($"(@orderList{i})");
                    obj.Add($"@orderList{i}", orderList[i]?? "0", DbType.String);
                }

                sql.Append(string.Join(", ", sqlList));
            }

            sql.Append("         ) AS TempData(OrderNumber)" + Environment.NewLine);
            sql.Append(")" + Environment.NewLine);
            sql.Append("SELECT temp.OrderNumber" + Environment.NewLine);
            sql.Append("FROM TempData temp" + Environment.NewLine);
            sql.Append("WHERE NOT EXISTS" + Environment.NewLine);
            sql.Append("(" + Environment.NewLine);
            sql.Append("    SELECT 1" + Environment.NewLine);
            sql.Append("    FROM [PracticeProject_StoreOrder] p" + Environment.NewLine);
            sql.Append("    WHERE p.OrderNumber = temp.OrderNumber" + Environment.NewLine);
            sql.Append(")" + Environment.NewLine);

            using (var conn = new SqlConnection(_dataBaseSettingProvider))
            {
                conn.Open();
                try
                {
                    return conn.Query(sql.ToString(), obj);
                }
                catch (Exception ex)
                {
                    var result = new List<dynamic>();

                    _logger.LogError($"Error occurred in [ GetNonExistentStoreOrders ] : {ex}");

                    result.Add(_messageHelper.CusErroCodeHelper("DataBaseErro"));

                    return result;
                }
            }
        }

        /// <summary>
        /// Insert Store Order
        /// </summary>
        /// <param name="dataModel">The list of data models to insert</param>
        /// <returns>True if the insertion was successful, false otherwise</returns>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create InsertStoreOrder
        /// </history>
        public bool InsertStoreOrder(List<GetExcelFromTemplateXlsxContextDataModel> dataModel)
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
                        return true;
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        _logger.LogError($"Error occurred in [ InsertStoreOrder ] : {ex}");
                        return false;
                    }

                }
            }
        }
    }
}
