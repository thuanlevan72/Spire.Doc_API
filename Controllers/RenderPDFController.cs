using Microsoft.AspNetCore.Mvc;
using Spire.Doc;
using WebApplication7.Helpers;
using WebApplication7.ViewModel;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace WebApplication7.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RenderPDFController : ControllerBase
    {
        private readonly ConvertWordToPDF _convertWordToPDF;

        public RenderPDFController()
        {
            _convertWordToPDF = new ConvertWordToPDF();
        }
        // GET: api/<RenderPDFController>
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/<RenderPDFController>/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }


        // POST api/<RenderPDFController>
        [HttpPost("GetPhieuKhamBenh")]
        public IActionResult GetPhieuKhamBenh([FromBody] PhieuKhamBenhViewModel value)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "DocxTemplates", "40. Phiếu khám bệnh.docx");
            byte[] fileContent = _convertWordToPDF.ReplaceTextInWord(value, path);


            var (fileData, fileName) = _convertWordToPDF.ConvertDocumentToPdf(fileContent);

            // Set up the response
            Response.Headers["Content-Disposition"] = $"inline; filename={fileName}";

            // Return the PDF file to the client.
            return new FileContentResult(fileData, "application/pdf");
        }

        // PUT api/<RenderPDFController>/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/<RenderPDFController>/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
