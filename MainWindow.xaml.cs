using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using Microsoft.Win32;

namespace ClaimsCalculatorApp
{
    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<RateChange> _rateChanges = new();
        private readonly ObservableCollection<PaymentEntry> _payments = new();
        private readonly ObservableCollection<DocumentCheckItem> _checklist = new();
        private readonly ObservableCollection<DailyRow> _scheduleRows = new();

        private ScheduleResult? _lastResult;

        public MainWindow()
        {
            InitializeComponent();

            // Sensible defaults so the app is immediately usable/testable.
            DpStartDate.SelectedDate = DateTime.Today;
            DpDateOfDeath.SelectedDate = DateTime.Today;
            DpLoanAgreementDate.SelectedDate = DateTime.Today;
            TxtSingleRate.Text = "0";
            TxtOpeningBalance.Text = "0";

            GridRateChanges.ItemsSource = _rateChanges;
            GridPayments.ItemsSource = _payments;
            GridSchedule.ItemsSource = _scheduleRows;

            foreach (var name in DocumentChecklistEngine.DocNames)
                _checklist.Add(new DocumentCheckItem(name, isChecked: false));
            ChecklistItemsControl.ItemsSource = _checklist;
        }

        private void BtnAddRate_Click(object sender, RoutedEventArgs e)
        {
            _rateChanges.Add(new RateChange { EffectiveDate = DateTime.Today, NewRatePercent = 0 });
        }

        private void BtnRemoveRate_Click(object sender, RoutedEventArgs e)
        {
            if (GridRateChanges.SelectedItem is RateChange rc)
                _rateChanges.Remove(rc);
        }

        private void BtnAddPayment_Click(object sender, RoutedEventArgs e)
        {
            _payments.Add(new PaymentEntry { Date = DateTime.Today, Amount = 0 });
        }

        private void BtnRemovePayment_Click(object sender, RoutedEventArgs e)
        {
            if (GridPayments.SelectedItem is PaymentEntry p)
                _payments.Remove(p);
        }

        private void BtnCalculate_Click(object sender, RoutedEventArgs e)
        {
            TxtStatus.Text = "";

            if (!decimal.TryParse(TxtOpeningBalance.Text, NumberStyles.Number, CultureInfo.InvariantCulture, out var openingBalance))
            {
                TxtStatus.Text = "Opening Balance must be a valid number.";
                return;
            }

            if (!decimal.TryParse(TxtSingleRate.Text, NumberStyles.Number, CultureInfo.InvariantCulture, out var singleRate))
            {
                TxtStatus.Text = "Interest Rate must be a valid number.";
                return;
            }

            if (DpStartDate.SelectedDate is null || DpDateOfDeath.SelectedDate is null)
            {
                TxtStatus.Text = "Start Date and Date of Death are required.";
                return;
            }

            var startDate = DpStartDate.SelectedDate.Value;
            var dateOfDeath = DpDateOfDeath.SelectedDate.Value;

            if (dateOfDeath.Date < startDate.Date)
            {
                TxtStatus.Text = "Date of Death cannot be before Start Date.";
                return;
            }

            var result = ClaimsCalculatorEngine.CalculateSchedule(
                openingBalance,
                startDate,
                dateOfDeath,
                singleRate,
                ChkMultipleRates.IsChecked == true,
                _rateChanges,
                _payments);

            _lastResult = result;

            _scheduleRows.Clear();
            foreach (var row in result.Rows)
                _scheduleRows.Add(row);

            TxtTotalInterest.Text = $"N$ {result.TotalInterest:#,##0.00}";
            TxtClaimAmount.Text = $"N$ {result.ClaimAmountInclInterest:#,##0.00}";

            TxtStatus.Text = $"Calculated {result.Rows.Count} day(s).";
        }

        private void BtnGenerateLetter_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult is null)
            {
                TxtStatus.Text = "Run Calculate first, then generate the letter.";
                return;
            }

            if (!decimal.TryParse(TxtSingleRate.Text, NumberStyles.Number, CultureInfo.InvariantCulture, out var singleRate))
                singleRate = 0;

            var (present, outstanding) = DocumentChecklistEngine.Evaluate(
                _checklist.Select(c => c.IsChecked).ToList());

            var input = new CoverLetterInput
            {
                ClientName = TxtClientName.Text,
                AccountRefNo = TxtAccountRefNo.Text,
                PolicyNo = TxtPolicyNo.Text,
                LoanAgreementDate = DpLoanAgreementDate.SelectedDate ?? DateTime.Today,
                DateOfDeath = DpDateOfDeath.SelectedDate ?? DateTime.Today,
                ClaimAmountInclInterest = _lastResult.ClaimAmountInclInterest,
                InterestRatePercent = singleRate,
                PresentDocuments = present,
                OutstandingDocuments = outstanding,
                ClaimsOfficerName = string.IsNullOrWhiteSpace(TxtClaimsOfficer.Text) ? "Claims Officer" : TxtClaimsOfficer.Text,
                SupervisorName = string.IsNullOrWhiteSpace(TxtSupervisor.Text) ? "Supervisor" : TxtSupervisor.Text
            };

            TxtLetterOutput.Text = CoverLetterGenerator.Generate(input);
            TxtStatus.Text = "Cover letter generated.";
        }

        private void BtnSaveLetter_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtLetterOutput.Text))
            {
                TxtStatus.Text = "Nothing to save — generate the letter first.";
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*",
                FileName = $"Claim Cover Letter - {TxtAccountRefNo.Text}.txt"
            };

            if (dialog.ShowDialog() == true)
            {
                System.IO.File.WriteAllText(dialog.FileName, TxtLetterOutput.Text);
                TxtStatus.Text = $"Saved to {dialog.FileName}";
            }
        }

        private void BtnCopyLetter_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(TxtLetterOutput.Text))
            {
                Clipboard.SetText(TxtLetterOutput.Text);
                TxtStatus.Text = "Letter copied to clipboard.";
            }
        }
    }
}
