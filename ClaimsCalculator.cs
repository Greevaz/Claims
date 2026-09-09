using System;
using System.Collections.Generic;
using System.Linq;

namespace ClaimsCalculatorApp
{
    public class ScheduleResult
    {
        public List<DailyRow> Rows { get; set; } = new();
        public decimal TotalInterest { get; set; }
        public decimal ClaimAmountInclInterest { get; set; }
    }

    /// <summary>
    /// Reproduces the "Template" sheet's daily interest logic:
    ///   - Simple daily interest = OpeningBalance * Rate% / 100 / 365
    ///   - Interest is capitalised (added to the balance) only at each
    ///     month-end and on the date of death, matching the original
    ///     EOMONTH/date-of-death compounding rule.
    ///   - When "Multiple Interest Rates" is Yes, each day's rate is the
    ///     most recent Rate Change whose Effective Date is on or before
    ///     that day; otherwise it falls back to the single base rate.
    ///   - Payments received are deducted from the balance on their date.
    ///
    /// Note: the original spreadsheet's row grid is hard-capped at 365
    /// days after the Start Date (a limitation of it being a fixed table
    /// of rows). This calculator has no such cap — it runs for the full
    /// Start Date -> Date of Death span, however long that is.
    /// </summary>
    public static class ClaimsCalculatorEngine
    {
        public static ScheduleResult CalculateSchedule(
            decimal openingBalance,
            DateTime startDate,
            DateTime dateOfDeath,
            decimal singleRatePercent,
            bool multipleRates,
            IEnumerable<RateChange> rateChanges,
            IEnumerable<PaymentEntry> payments)
        {
            var result = new ScheduleResult();

            if (dateOfDeath.Date < startDate.Date)
                return result; // nothing to calculate — invalid range

            var sortedRates = rateChanges
                .OrderBy(r => r.EffectiveDate.Date)
                .ToList();

            decimal GetRateForDate(DateTime date)
            {
                if (!multipleRates)
                    return singleRatePercent;

                RateChange? applicable = null;
                foreach (var rc in sortedRates)
                {
                    if (rc.EffectiveDate.Date <= date.Date)
                        applicable = rc; // keep the latest one that still qualifies
                    else
                        break;
                }
                return applicable?.NewRatePercent ?? singleRatePercent;
            }

            var paymentsByDate = payments
                .GroupBy(p => p.Date.Date)
                .ToDictionary(g => g.Key, g => g.Sum(p => p.Amount));

            decimal openingBal = openingBalance;
            decimal accruedSinceLastCompound = 0m;
            decimal totalInterest = 0m;

            for (DateTime date = startDate.Date; date <= dateOfDeath.Date; date = date.AddDays(1))
            {
                decimal rate = GetRateForDate(date);
                decimal dailyInterest = openingBal * rate / 100m / 365m;

                accruedSinceLastCompound += dailyInterest;

                bool isMonthEnd = date == new DateTime(date.Year, date.Month,
                    DateTime.DaysInMonth(date.Year, date.Month));
                bool isDeathDate = date == dateOfDeath.Date;
                bool isCompoundingDay = isMonthEnd || isDeathDate;

                decimal paymentsToday = paymentsByDate.TryGetValue(date, out var amt) ? amt : 0m;
                const decimal feesToday = 0m; // the source workbook has no fee entries

                decimal? compoundToday = null;
                decimal closingBal;

                if (isCompoundingDay)
                {
                    compoundToday = accruedSinceLastCompound;
                    closingBal = openingBal - paymentsToday + accruedSinceLastCompound + feesToday;
                    totalInterest += accruedSinceLastCompound;
                    accruedSinceLastCompound = 0m;
                }
                else
                {
                    closingBal = openingBal - paymentsToday + feesToday;
                }

                result.Rows.Add(new DailyRow
                {
                    Date = date,
                    OpeningBalance = openingBal,
                    DailyInterest = dailyInterest,
                    InterestRatePercent = rate,
                    CompoundInterest = compoundToday,
                    Payments = paymentsToday,
                    Fees = feesToday,
                    ClosingBalance = closingBal
                });

                openingBal = closingBal;
            }

            result.TotalInterest = totalInterest;
            result.ClaimAmountInclInterest = result.Rows.Count > 0
                ? result.Rows[^1].ClosingBalance
                : 0m;

            return result;
        }
    }

    /// <summary>
    /// Reproduces the Sheet2.UpdateDocumentLists VBA logic, including the
    /// special rule that suppresses "Burial Removal Order" from the
    /// Outstanding list under specific conditions.
    /// </summary>
    public static class DocumentChecklistEngine
    {
        public static readonly string[] DocNames = new[]
        {
            "PostFin Claim Form",
            "Deceased Identity Document",
            "Death Certificate",
            "Loan Statement",
            "Loan Settlement",
            "Arrear History Statement",
            "Medical Certificate of the Causes of Death",
            "Claimant's Identity Document",
            "Burial Removal Order",
            "Loan Agreement",
            "Health Passport/ Medical Aid Card"
        };

        public static (List<string> Present, List<string> Outstanding) Evaluate(IList<bool> checks)
        {
            if (checks.Count != DocNames.Length)
                throw new ArgumentException($"Expected {DocNames.Length} checklist values.");

            // Are all documents except BRO/LoanAgreement/HealthPassport present?
            bool othersAllPresent = true;
            for (int i = 0; i <= 7; i++)
            {
                if (!checks[i]) { othersAllPresent = false; break; }
            }

            bool broMissing = !checks[8];
            bool loanAgreementMissing = !checks[9];
            bool healthPassportMissing = !checks[10];

            bool suppressBro = false;
            if (othersAllPresent && broMissing)
            {
                if (!loanAgreementMissing && !healthPassportMissing)
                    suppressBro = true; // BRO is the only thing missing
                else if (loanAgreementMissing && healthPassportMissing)
                    suppressBro = true; // BRO missing together with LA and HP only
            }

            var present = new List<string>();
            var outstanding = new List<string>();

            for (int i = 0; i < DocNames.Length; i++)
            {
                if (checks[i])
                {
                    present.Add(DocNames[i]);
                }
                else if (i == 8 && suppressBro)
                {
                    // Burial Removal Order missing but suppressed — skip both lists
                }
                else
                {
                    outstanding.Add(DocNames[i]);
                }
            }

            return (present, outstanding);
        }
    }
}
