using System;
using System.IO;
using System.Threading.Tasks;
using DevExpress.Spreadsheet;

namespace BlazorServer_Spreadsheet_MailMerge.Code {
    public class DocumentService {
        public async Task<byte[]> GetXlsxDocumentAsync(double loanAmount, int periodInYears, double interestRate, DateTime startDateOfLoan) {
            using var workbook = await GenerateDocumentAsync(loanAmount, periodInYears, interestRate, startDateOfLoan);
            return await workbook.SaveDocumentAsync(DocumentFormat.Xlsx);
        }

        public async Task<byte[]> GetPdfDocumentAsync(double loanAmount, int periodInYears, double interestRate, DateTime startDateOfLoan) {
            using var workbook = await GenerateDocumentAsync(loanAmount, periodInYears, interestRate, startDateOfLoan);
            var ms = new MemoryStream();
            await workbook.ExportToPdfAsync(ms);
            return ms.ToArray();
        }

        public async Task<byte[]> GetHtmlDocumentAsync(double loanAmount, int periodInYears, double interestRate, DateTime startDateOfLoan) {
            using var workbook = await GenerateDocumentAsync(loanAmount, periodInYears, interestRate, startDateOfLoan);
            var ms = new MemoryStream();
            await workbook.ExportToHtmlAsync(ms, workbook.Worksheets[0]);
            return ms.ToArray();
        }

        async Task<Workbook> GenerateDocumentAsync(double loanAmount, int periodInYears, double interestRate, DateTime startDateOfLoan) {
            var workbook = new Workbook();
            await workbook.LoadDocumentAsync("Data/LoanAmortizationSchedule_template.xltx");
            new LoanAmortizationScheduleDocumentGenerator(workbook)
                .GenerateDocument(loanAmount, periodInYears, interestRate, startDateOfLoan);
            return workbook;
        }
    }
}
