
using EXCEL_PDF_Practice_ServiceLayer.Implement;
using EXCEL_PDF_Practice_ServiceLayer.Interface;
using NLog;
using NLog.Web;

namespace EXCEL_PDF_Practice_sln
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var logger = NLog.LogManager.Setup().LoadConfigurationFromAppSettings().GetCurrentClassLogger();
            logger.Debug("init main");

            try
            {
                var builder = WebApplication.CreateBuilder(args);

                //NLog: Setup NLog for Dependency injection
                builder.Logging.ClearProviders();
                builder.Host.UseNLog();

                // Add services to the container.
                builder.Services.AddControllers();
                builder.Services.AddScoped<IExcelPdfPracticeService, ExcelPdfPracticeService>();

                // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
                builder.Services.AddEndpointsApiExplorer();
                builder.Services.AddSwaggerGen();

                var app = builder.Build();

                // Configure the HTTP request pipeline.
                if (app.Environment.IsDevelopment())
                {
                    app.UseSwagger();
                    app.UseSwaggerUI();
                }

                app.UseHttpsRedirection();

                app.UseAuthorization();


                app.MapControllers();

                app.Run();
            }
            catch (Exception ex)
            {
                // NLog: catch setup errors
                logger.Error(ex, "Stopp program because of exception");
                throw;
            }
            finally
            {
                // Ensure to flush and stop internal timers/thresds before application-exit (Avoid segmentation fault on Linux )
                NLog.LogManager.Shutdown();
            }
        }
    }
}
