using System;
using System.Drawing;
using DevExpress.Spreadsheet;

namespace BlazorServer_Spreadsheet_MailMerge.Code {
    public class LoanAmortizationScheduleDocumentGenerator {
        IWorkbook workbook;

        public LoanAmortizationScheduleDocumentGenerator(IWorkbook workbook) {
            this.workbook = workbook;
        }

        #region Properties
        Worksheet Sheet { get { return workbook.Worksheets[0]; } }
        DateTime StartDateOfLoan { get { return Sheet["E8"].Value.DateTimeValue; } set { Sheet["E8"].Value = value; } }
        double LoanAmount { get { return Sheet["E4"].Value.NumericValue; } set { Sheet["E4"].Value = value; } }
        double InterestRate { get { return Sheet["E5"].Value.NumericValue; } set { Sheet["E5"].Value = value; } }
        int PeriodInYears { get { return (int)Sheet["E6"].Value.NumericValue; } set { Sheet["E6"].Value = value; } }
        int ActualNumberOfPayments { get { return (int)Math.Round(Sheet["I6"].Value.NumericValue); } }
        int ScheduledNumberOfPayments { get { return (int)Math.Round(Sheet["I5"].Value.NumericValue); } }
        string ActualLastRow { get { return (11 + ActualNumberOfPayments).ToString(); } }
        #endregion

        public void GenerateDocument(double loanAmount, int periodInYears, double interestRate, DateTime startDateOfLoan) {
            workbook.BeginUpdate();
            try {
                Cleanup();
                LoanAmount = loanAmount;
                InterestRate = interestRate;
                PeriodInYears = periodInYears;
                StartDateOfLoan = startDateOfLoan;
                GenerateAnnuityPaymentsContent();
                ApplyFormatting();
                AdjustPrintOptions();
            }
            finally {
                workbook.EndUpdate();
            }
        }

        void Cleanup() {
            var range = Sheet.GetDataRange().Exclude(Sheet["1:11"]);
            if (range != null)
                range.Clear();
            Sheet["I4"].ClearContents();
            Sheet["I6:I8"].ClearContents();
            workbook.DefinedNames.Clear();
        }

        void GenerateAnnuityPaymentsContent() {
            AddDefinedNamesForAnnuityPayments();

            Sheet["I4"].FormulaInvariant = "=PMT(Interest_Rate_Per_Month,Scheduled_Number_Payments,-Loan_Amount)";
            Sheet["I5"].FormulaInvariant = "=Loan_Years*Number_of_Payments_Per_Year";
            Sheet["I6"].FormulaInvariant = "=ROUNDUP(Actual_Number_Payments,0)";
            workbook.Calculate();
            Sheet["I7"].FormulaInvariant = "=SUM(F12:F" + ActualLastRow + ")";
            Sheet["I8"].FormulaInvariant = "=SUM($I$12:$I$" + ActualLastRow + ")";

            if (ScheduledNumberOfPayments == 0)
                return;

            for (int i = 0; i < ActualNumberOfPayments; i++)
                Sheet["B" + (i + 12).ToString()].Value = i + 1;

            Sheet["C12:C" + ActualLastRow].FormulaInvariant = "=DATE(YEAR(Loan_Start),MONTH(Loan_Start)+(B12)*12/Number_of_Payments_Per_Year,DAY(Loan_Start))";
            Sheet["D12"].Formula = "=Loan_Amount";

            if (ScheduledNumberOfPayments > 1)
                Sheet["D13:D" + ActualLastRow].Formula = "=J12";

            Sheet["E12:E" + ActualLastRow].FormulaInvariant = "=IF(D12>0,IF(Scheduled_payment<D12, Scheduled_payment, D12),0)";
            Sheet["F12:F" + ActualLastRow].FormulaInvariant = "=IF(Extra_Payments<>0, IF(Scheduled_payment<D12, G12-E12, 0), 0)";
            Sheet["G12:G" + ActualLastRow].FormulaInvariant = "=H12+I12";
            Sheet["H12:H" + ActualLastRow].FormulaInvariant = "=IF(J12>0,PPMT(Interest_Rate_Per_Month,B12,Actual_Number_Payments,-Loan_Amount),D12)";
            Sheet["I12:I" + ActualLastRow].FormulaInvariant = "=IF(D12>0,IPMT(Interest_Rate_Per_Month,B12,Actual_Number_Payments,-Loan_Amount),0)";
            Sheet["J12:J" + ActualLastRow].FormulaInvariant = "=IF(D12-PPMT(Interest_Rate_Per_Month,B12,Actual_Number_Payments,-Loan_Amount)>0,D12-PPMT(Interest_Rate_Per_Month,B12,Actual_Number_Payments,-Loan_Amount),0)";
            Sheet["K12:K" + ActualLastRow].FormulaInvariant = "=SUM($I$12:$I12)";

            workbook.Calculate();
        }

        void AddDefinedNamesForAnnuityPayments() {
            string sheetName = "'" + Sheet.Name + "'";
            char separator = workbook.Options.Culture.TextInfo.ListSeparator[0];

            DefinedNameCollection definedNames = workbook.DefinedNames;
            definedNames.Add("Loan_Amount", sheetName + "!$E$4");
            definedNames.Add("Interest_Rate", sheetName + "!$E$5");
            definedNames.Add("Loan_Years", sheetName + "!$E$6");
            definedNames.Add("Number_of_Payments_Per_Year", sheetName + "!$E$7");
            definedNames.Add("Loan_Start", sheetName + "!$E$8");
            definedNames.Add("Extra_Payments", sheetName + "!$E$9");
            definedNames.Add("Scheduled_payment", sheetName + "!$I$4");
            definedNames.Add("Scheduled_Number_Payments", sheetName + "!$I$5");
            definedNames.Add("Interest_Rate_Per_Month", "=Interest_Rate/Number_of_Payments_Per_Year");
            definedNames.Add("Actual_Number_Payments", "=NPER(Interest_Rate_Per_Month" + separator + " " + sheetName + "!$I$4+Extra_Payments" + separator + " -Loan_Amount)");
        }

        void ApplyFormatting() {
            CellRange range;
            for (int i = 1; i < ActualNumberOfPayments; i += 2) {
                range = Sheet.Range.FromLTRB(1, 11 + i, 10, 11 + i);
                range.Fill.BackgroundColor = Color.FromArgb(217, 217, 217);
            }

            range = Sheet["B11:K" + ActualLastRow];
            Formatting formatting = range.BeginUpdateFormatting();
            try {
                formatting.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin;
                formatting.Borders.InsideVerticalBorders.Color = Color.White;
                formatting.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            }
            finally {
                range.EndUpdateFormatting(formatting);
            }

            Sheet["B12:C" + ActualLastRow].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right;
            Sheet["C11:C" + ActualLastRow].NumberFormat = "m/d/yyyy";
            Sheet["D11:K" + ActualLastRow].NumberFormat = "_(\\$* #,##0.00_);_(\\$ (#,##0.00);_(\\$* \" - \"??_);_(@_)";
        }

        void AdjustPrintOptions() {
            Sheet.SetPrintRange(Sheet.GetDataRange());
            Sheet.PrintOptions.FitToPage = true;
            Sheet.PrintOptions.FitToWidth = 1;
            Sheet.PrintOptions.FitToHeight = 0; // automatic
        }

    }
}
