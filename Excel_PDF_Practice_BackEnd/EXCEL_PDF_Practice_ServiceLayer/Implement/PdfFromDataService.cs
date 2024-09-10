using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using DocumentFormat.OpenXml.Spreadsheet;
using EXCEL_PDF_Practice_ServiceLayer.Interface;
using Microsoft.Extensions.Logging;
using PdfSharpCore.Drawing;
using PdfSharpCore.Fonts;
using PdfSharpCore.Pdf;
using PdfSharpCore.Utils;
using SixLabors.Fonts;

namespace EXCEL_PDF_Practice_ServiceLayer.Implement
{
    public class PdfFromDataService : IPdfFromDataService
    {
        private readonly ILogger<PdfFromDataService> _logger;
        private readonly IMapper _mapper;
        private readonly IServiceProvider _serviceProvider;

        public PdfFromDataService(
            ILogger<PdfFromDataService> logger,
            IMapper mapper,
            IServiceProvider serviceProvider)
        {
            _logger = logger;
            _mapper = mapper;
            _serviceProvider = serviceProvider;
        }

        public PdfDocument BuildPdfFromDataBase()
        {
            PdfDocument document = new PdfDocument();

            return document;
        }
    }
}
