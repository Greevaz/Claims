using System;
using System.Collections.Generic;
using System.Text;

namespace ClaimsCalculatorApp
{
    public class CoverLetterInput
    {
        public string ClientName { get; set; } = "";
        public string AccountRefNo { get; set; } = "";
        public string PolicyNo { get; set; } = "";
        public DateTime LoanAgreementDate { get; set; }
        public DateTime DateOfDeath { get; set; }
        public decimal ClaimAmountInclInterest { get; set; }
        public decimal InterestRatePercent { get; set; }
        public List<string> PresentDocuments { get; set; } = new();
        public List<string> OutstandingDocuments { get; set; } = new();
        public string ClaimsOfficerName { get; set; } = "Anthony Greeves";
        public string SupervisorName { get; set; } = "Wesley Ramothibe";
    }

    public static class CoverLetterGenerator
    {
        public static string Generate(CoverLetterInput input)
        {
            var sb = new StringBuilder();

            sb.AppendLine("         Gustav Volgts Centre,");
            sb.AppendLine("         Erf 5518, 131 Independence Avenue, Windhoek");
            sb.AppendLine("         Tel: 061 308724");
            sb.AppendLine();
            sb.AppendLine(DateTime.Today.ToString("dd MMMM yyyy"));
            sb.AppendLine();
            sb.AppendLine("The Manager Claims");
            sb.AppendLine("Hollard Insurance Namibia");
            sb.AppendLine("PO Box 5077");
            sb.AppendLine("Ausspannplatz");
            sb.AppendLine("Windhoek");
            sb.AppendLine("Namibia");
            sb.AppendLine();
            sb.AppendLine("Dear Sir / Madam");
            sb.AppendLine();
            sb.AppendLine($"Life Assured: {input.ClientName}");
            sb.AppendLine($"Policy No: {input.PolicyNo}");
            sb.AppendLine($"Loan Account Number: AZ {input.AccountRefNo}");
            sb.AppendLine();
            sb.AppendLine(
                $"The client entered into a loan agreement with us on " +
                $"{input.LoanAgreementDate:dd MMMM yyyy} and passed away on " +
                $"{input.DateOfDeath:dd MMMM yyyy}.");
            sb.AppendLine();
            sb.AppendLine(
                $"We herewith submit a death claim against the credit life policy for N$ " +
                $"{input.ClaimAmountInclInterest:#,##0.00} with interest at a rate of " +
                $"{input.InterestRatePercent} percent.");
            sb.AppendLine();
            sb.AppendLine("Attached herewith please find the following documents:");
            sb.AppendLine();
            foreach (var doc in input.PresentDocuments)
                sb.AppendLine($"    - {doc}");

            if (input.OutstandingDocuments.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Outstanding Documents:");
                foreach (var doc in input.OutstandingDocuments)
                    sb.AppendLine($"    - {doc}");
            }

            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine($"{input.ClaimsOfficerName,-40}{input.SupervisorName}");
            sb.AppendLine($"{"Claims Officer",-40}Supervisor: Back Office");
            sb.AppendLine();
            sb.AppendLine("Company Registration Number: 2007 / 0642");
            sb.AppendLine(
                "CEO: Sam Ikela | Directors: James Cumming (Chairman); Willem Mouton; " +
                "Sonia Bergh; Eldorette Harmse; Batsirai Pfigirai");

            return sb.ToString();
        }
    }
}
