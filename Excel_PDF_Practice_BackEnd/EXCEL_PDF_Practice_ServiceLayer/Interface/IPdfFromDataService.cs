using PdfSharpCore.Pdf;

namespace EXCEL_PDF_Practice_ServiceLayer.Interface
{
    public interface IPdfFromDataService
    {
        public PdfDocument BuildPdfFromDataBase();
    }
}