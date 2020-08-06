using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using BlazorServer_Spreadsheet_MailMerge.Code;

namespace BlazorServer_Spreadsheet_MailMerge.Api {
    [Route("api/[controller]")]
    [ApiController]
    public class ExportController : ControllerBase {
        readonly DocumentService documentService;

        public ExportController(DocumentService documentService) {
            this.documentService = documentService;
        }

        [HttpGet]
        [Route("[action]")]
        public async Task<IActionResult> Xlsx([FromQuery] double loanAmount, [FromQuery] int periodInYears, [FromQuery] double interestRate, [FromQuery] DateTime startDateOfLoan) {
            var document = await documentService.GetXlsxDocumentAsync(loanAmount, periodInYears, interestRate, startDateOfLoan);
            return File(document, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "output.xlsx");
        }

        [HttpGet]
        [Route("[action]")]
        public async Task<IActionResult> Pdf([FromQuery] double loanAmount, [FromQuery] int periodInYears, [FromQuery] double interestRate, [FromQuery] DateTime startDateOfLoan) {
            var document = await documentService.GetPdfDocumentAsync(loanAmount, periodInYears, interestRate, startDateOfLoan);
            return File(document, "application/pdf", "output.pdf");
        }
    }
}
