using Microsoft.AspNetCore.Mvc;

namespace EXCEL_PDF_Practice_sln.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class PDFBuildController : ControllerBase
    {

        public IActionResult Index()
        {
            return Ok();
        }

        [HttpPost]
        public IActionResult GetPdfFromDataBase()
        {


            return Ok();
        }
    }
}
