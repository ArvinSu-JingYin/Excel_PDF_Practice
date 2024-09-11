using PdfSharpCore.Pdf;

namespace EXCEL_PDF_Practice_ServiceLayer.Interface
{
    public interface IPdfFromDataService
    {
        /// <summary>
        /// Build PDF from data base test
        /// </summary>
        /// <returns>PDF document</returns>
        /// <info>Author: Arvin; Date: 2024/09/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/12  1.00    Arvin       Create BuildPdfFromDataBaseTest
        /// </history>
        public PdfDocument BuildPdfFromDataBaseTest();

        /// <summary>
        /// Build PDF from data base
        /// </summary>
        /// <param name="num">The order number to filter by</param>
        /// <returns>PDF document</returns>
        /// <info>Author: Arvin; Date: 2024/09/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/12  1.00    Arvin       Create BuildPdfFromDataBase
        /// </history>
        public PdfDocument BuildPdfFromDataBase(string num);
    }
}