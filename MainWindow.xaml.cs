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
        private readonly ObservableCollection<ClaimField> _claimFields = new();
        private readonly RateChangesStore _rateChangesStore = new();

        private ScheduleResult? _lastResult;
        private LocalExcelClaimsService? _localExcelClaimsService;
        private ClaimRecord? _loadedClaim;
        private CancellationTokenSource? _claimLookupCancellation;
        private bool _loadingRateChanges;

        private void SetLoading(bool isLoading)
        {
            LoadingOverlay.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
        }

        public MainWindow()
        {
            InitializeComponent();
            Closing += MainWindow_Closing;

            // Keep database-driven dates and opening balance blank until a claim is loaded.
            DpStartDate.SelectedDate = null;
            DpDateOfDeath.SelectedDate = null;
            DpLoanAgreementDate.SelectedDate = null;
            TxtSingleRate.Text = "0.00";
            TxtOpeningBalance.Text = "";
            SetClaimInputsEnabled(false);

            GridRateChanges.ItemsSource = _rateChanges;
            _rateChanges.CollectionChanged += RateChanges_CollectionChanged;
            GridPayments.ItemsSource = _payments;
            GridSchedule.ItemsSource = _scheduleRows;
            GridClaimFields.ItemsSource = _claimFields;

            foreach (var name in DocumentChecklistEngine.DocNames)
            {
                var item = new DocumentCheckItem(name, isChecked: true);
                item.PropertyChanged += Checklist_PropertyChanged;
                _checklist.Add(item);
            }
            ChecklistItemsControl.ItemsSource = _checklist;
            TxtStatus.Text = "";
        }

        private void BtnOpenClaimsCalculator_Click(object sender, RoutedEventArgs e)
        {
            HomeGrid.Visibility = Visibility.Collapsed;
            MainTabs.Visibility = Visibility.Visible;
            DatabaseGrid.Visibility = Visibility.Collapsed;
            BtnClaimsBack.Visibility = Visibility.Visible;
            MainTabs.SelectedItem = ClaimsCalculatorTab;
            TxtStatus.Text = "Claims Calculator and Cover Page opened.";
        }

        private void BtnOpenDatabaseRecords_Click(object sender, RoutedEventArgs e)
        {
            HomeGrid.Visibility = Visibility.Collapsed;
            MainTabs.Visibility = Visibility.Collapsed;
            DatabaseGrid.Visibility = Visibility.Visible;
            TxtDatabaseSearch.Focus();
            TxtStatus.Text = "Enter an AZ to load a database record.";
        }

        private void BtnBackToHome_Click(object sender, RoutedEventArgs e)
        {
            DatabaseGrid.Visibility = Visibility.Collapsed;
            MainTabs.Visibility = Visibility.Collapsed;
            BtnClaimsBack.Visibility = Visibility.Collapsed;
            HomeGrid.Visibility = Visibility.Visible;
            TxtStatus.Text = "";
        }

        private async void BtnSearchDatabase_Click(object sender, RoutedEventArgs e)
        {
            var query = TxtDatabaseSearch.Text.Trim();
            var fieldName = ((System.Windows.Controls.ComboBoxItem)CmbDatabaseSearchField.SelectedItem).Content.ToString()!;
            if (string.IsNullOrWhiteSpace(query))
            {
                TxtStatus.Text = "Enter a search value.";
                return;
            }

            try
            {
                SetLoading(true);
                TxtStatus.Text = "Loading database record...";
                _localExcelClaimsService ??= new LocalExcelClaimsService(GraphSettings.Load());
                var claim = await Task.Run(() => _localExcelClaimsService.Find(fieldName, query));
                if (claim is null)
                {
                    _loadedClaim = null;
                    _claimFields.Clear();
                    TxtStatus.Text = $"No record found for {fieldName} '{query}'.";
                    return;
                }

                _loadedClaim = claim;
                _claimFields.Clear();
                for (var i = 0; i < claim.Headers.Length; i++)
                    _claimFields.Add(new ClaimField(claim.Headers[i],
                        i < claim.Values.Length ? claim.Values[i] : null));
                TxtStatus.Text = $"Loaded database record for {claim.Name}.";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"Excel load failed: {ex.Message}";
            }
            finally
            {
                SetLoading(false);
            }
        }

        private void TxtDatabaseSearch_PreviewTextInput(object sender,
            System.Windows.Input.TextCompositionEventArgs e) =>
            e.Handled = IsDatabaseAzSearch() &&
                        e.Text.Any(character => !char.IsDigit(character));

        private void TxtDatabaseSearch_Pasting(object sender, System.Windows.DataObjectPastingEventArgs e)
        {
            if (!e.DataObject.GetDataPresent(typeof(string)) ||
                e.DataObject.GetData(typeof(string)) is not string pasted ||
                (IsDatabaseAzSearch() && pasted.Any(character => !char.IsDigit(character))))
                e.CancelCommand();
        }

        private bool IsDatabaseAzSearch() =>
            CmbDatabaseSearchField.SelectedIndex == 0;

        private void TxtDatabaseSearch_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                e.Handled = true;
                BtnSearchDatabase_Click(sender, new RoutedEventArgs());
            }
        }

        private void BtnAddRate_Click(object sender, RoutedEventArgs e)
        {
            var change = new RateChange { EffectiveDate = DateTime.Today, NewRatePercent = 0 };
            _rateChanges.Add(change);
            CalculateClaims();
        }

        private void BtnRemoveRate_Click(object sender, RoutedEventArgs e)
        {
            if (GridRateChanges.SelectedItem is RateChange rc)
            {
                _rateChanges.Remove(rc);
                CalculateClaims();
            }
        }

        private void RateChanges_CollectionChanged(object? sender,
            System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems is not null)
                foreach (RateChange change in e.OldItems)
                    change.PropertyChanged -= RateChange_PropertyChanged;
            if (e.NewItems is not null)
                foreach (RateChange change in e.NewItems)
                    change.PropertyChanged += RateChange_PropertyChanged;
            SaveRateChanges();
        }

        private void RateChange_PropertyChanged(object? sender,
            System.ComponentModel.PropertyChangedEventArgs e)
        {
            SaveRateChanges();
            CalculateClaims();
        }

        private void GridRateChanges_CellEditEnding(object sender,
            System.Windows.Controls.DataGridCellEditEndingEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                GridRateChanges.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
                GridRateChanges.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);
                SaveRateChanges();
                CalculateClaims();
            }));
        }

        private void BtnSaveRates_Click(object sender, RoutedEventArgs e)
        {
            GridRateChanges.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
            GridRateChanges.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);
            SaveRateChanges();
            TxtStatus.Text = "Rate changes saved.";
        }

        private void GridRateChanges_PreviewTextInput(object sender,
            System.Windows.Input.TextCompositionEventArgs e)
        {
            if (GridRateChanges.CurrentColumn?.Header?.ToString() == "New Rate (%)")
            {
                var textBox = e.OriginalSource as System.Windows.Controls.TextBox;
                var proposed = textBox is null ? e.Text : textBox.Text.Insert(textBox.SelectionStart, e.Text);
                e.Handled = proposed.Any(character => !char.IsDigit(character) && character != '.');
            }
        }

        private void GridRateChanges_Pasting(object sender,
            System.Windows.DataObjectPastingEventArgs e)
        {
            if (GridRateChanges.CurrentColumn?.Header?.ToString() == "New Rate (%)" &&
                (!e.DataObject.GetDataPresent(typeof(string)) ||
                 e.DataObject.GetData(typeof(string)) is not string pasted ||
                 pasted.Any(character => !char.IsDigit(character) && character != '.')))
                e.CancelCommand();
        }

        private void RateTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.TextBox textBox)
                textBox.SelectAll();
        }

        private void RateTextBox_PreviewTextInput(object sender,
            System.Windows.Input.TextCompositionEventArgs e)
        {
            if (sender is not System.Windows.Controls.TextBox textBox)
                return;
            var proposed = textBox.Text.Remove(textBox.SelectionStart, textBox.SelectionLength)
                .Insert(textBox.SelectionStart, e.Text);
            e.Handled = proposed.Any(character => !char.IsDigit(character) && character != '.');
        }

        private void RateTextBox_Pasting(object sender, System.Windows.DataObjectPastingEventArgs e)
        {
            if (sender is System.Windows.Controls.TextBox &&
                (!e.DataObject.GetDataPresent(typeof(string)) ||
                 e.DataObject.GetData(typeof(string)) is not string pasted ||
                 pasted.Any(character => !char.IsDigit(character) && character != '.')))
                e.CancelCommand();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            GridRateChanges.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
            GridRateChanges.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);
            SaveRateChanges();
        }

        private void SaveRateChanges()
        {
            if (!_loadingRateChanges && !string.IsNullOrWhiteSpace(TxtAccountRefNo.Text))
            {
                try
                {
                    _rateChangesStore.Save(TxtAccountRefNo.Text, _rateChanges);
                }
                catch (Exception ex)
                {
                    TxtStatus.Text = $"Rate changes could not be saved: {ex.Message}";
                }
            }
        }

        private void BtnAddPayment_Click(object sender, RoutedEventArgs e)
        {
            var payment = new PaymentEntry { Date = DateTime.Today, Amount = 0 };
            payment.PropertyChanged += Payment_PropertyChanged;
            _payments.Add(payment);
            CalculateClaims();
        }

        private void BtnRemovePayment_Click(object sender, RoutedEventArgs e)
        {
            if (GridPayments.SelectedItem is PaymentEntry p)
            {
                p.PropertyChanged -= Payment_PropertyChanged;
                _payments.Remove(p);
                CalculateClaims();
            }
        }

        private void Payment_PropertyChanged(object? sender,
            System.ComponentModel.PropertyChangedEventArgs e) => CalculateClaims();

        private void Checklist_PropertyChanged(object? sender,
            System.ComponentModel.PropertyChangedEventArgs e) => CalculateClaims();

        private void ClaimInputChanged(object sender, System.Windows.Controls.TextChangedEventArgs e) =>
            CalculateClaims();

        private void ClaimDateChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) =>
            CalculateClaims();

        private void ClaimOptionChanged(object sender, RoutedEventArgs e) =>
            CalculateClaims();

        private async void BtnSearchClaim_Click(object sender, RoutedEventArgs e)
        {
            var az = TxtAccountRefNo.Text.Trim();
            if (string.IsNullOrWhiteSpace(az))
            {
                _loadedClaim = null;
                SetClaimInputsEnabled(false);
                TxtStatus.Text = "Enter an AZ to search.";
                return;
            }

            try
            {
                SetLoading(true);
                SetClaimInputsEnabled(false);
                TxtStatus.Text = "Loading claim from Excel...";
                var service = _localExcelClaimsService ??= new LocalExcelClaimsService(GraphSettings.Load());
                var claim = await Task.Run(() => service.FindByAz(az));
                ApplyLoadedClaim(claim, az);
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"Excel load failed: {ex.Message}";
            }
            finally
            {
                SetLoading(false);
            }
        }

        private void TxtAccountRefNo_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                e.Handled = true;
                BtnSearchClaim_Click(sender, new RoutedEventArgs());
            }
        }

        private void TxtAccountRefNo_PreviewTextInput(object sender,
            System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = e.Text.Any(character => !char.IsDigit(character));
        }

        private void TxtAccountRefNo_Pasting(object sender, System.Windows.DataObjectPastingEventArgs e)
        {
            if (!e.DataObject.GetDataPresent(typeof(string)) ||
                e.DataObject.GetData(typeof(string)) is not string pasted ||
                pasted.Any(character => !char.IsDigit(character)))
                e.CancelCommand();
        }

        private void SetClaimInputsEnabled(bool isEnabled)
        {
            TxtClientName.IsEnabled = isEnabled;
            TxtPolicyNo.IsEnabled = isEnabled;
            TxtOpeningBalance.IsEnabled = isEnabled;
            DpLoanAgreementDate.IsEnabled = isEnabled;
            DpStartDate.IsEnabled = isEnabled;
            DpDateOfDeath.IsEnabled = isEnabled;
            TxtSingleRate.IsEnabled = isEnabled;
            ChkMultipleRates.IsEnabled = isEnabled;
            TxtClaimsOfficer.IsEnabled = isEnabled;
            TxtSupervisor.IsEnabled = isEnabled;
            RateChangesTab.IsEnabled = isEnabled;
            PaymentsTab.IsEnabled = isEnabled;
            ChecklistTab.IsEnabled = isEnabled;
            ResultsTab.IsEnabled = isEnabled;
            CoverLetterTab.IsEnabled = isEnabled;
        }

        private void ApplyLoadedClaim(ClaimRecord? claim, string az)
        {
            if (claim is null)
            {
                _loadedClaim = null;
                SetClaimInputsEnabled(false);
                TxtStatus.Text = $"No claim found for AZ '{az}'.";
                return;
            }

            _loadedClaim = claim;
            SetClaimInputsEnabled(true);
            _loadingRateChanges = true;
            try
            {
                _rateChanges.Clear();
                foreach (var change in _rateChangesStore.Load(az))
                    _rateChanges.Add(change);
            }
            finally
            {
                _loadingRateChanges = false;
            }
            _claimFields.Clear();
            for (var i = 0; i < claim.Headers.Length; i++)
            {
                var value = i < claim.Values.Length ? claim.Values[i] : null;
                _claimFields.Add(new ClaimField(claim.Headers[i], value));
            }
            TxtClientName.Text = claim.Name;
            TxtPolicyNo.Text = claim.PolicyNumber;
            // Opening Balance is entered separately and is never loaded from Excel.
            TxtOpeningBalance.Text = "";
            // Map each date field to its corresponding Excel database column.
            DpLoanAgreementDate.SelectedDate = claim.LoanStartDate?.Date;
            DpStartDate.SelectedDate = null;
            DpDateOfDeath.SelectedDate = claim.IncidentDate?.Date;

            TxtStatus.Text = $"Loaded claim for {claim.Name}.";
        }

        private async void BtnSaveClaim_Click(object sender, RoutedEventArgs e)
        {
            if (_loadedClaim is null)
            {
                TxtStatus.Text = "Load a claim first, then save it.";
                return;
            }

            try
            {
                SetLoading(true);
                TxtStatus.Text = "Saving claim to Excel...";
                _localExcelClaimsService ??= new LocalExcelClaimsService(GraphSettings.Load());
                var values = _claimFields.Select(field => (object?)field.Value).ToArray();
                var valueFor = (string name) =>
                {
                    var field = _claimFields.FirstOrDefault(item =>
                        string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
                    return field?.Value ?? "";
                };
                var parsedAmount = decimal.TryParse(valueFor("Loan Amount"), NumberStyles.Number,
                    CultureInfo.InvariantCulture, out var amount) ? amount : (decimal?)null;
                var parsedLoanDate = DateTime.TryParse(valueFor("Loan Start Date"),
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var loanDate) ? loanDate : (DateTime?)null;
                var parsedIncidentDate = DateTime.TryParse(valueFor("Incident Date"),
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var incidentDate) ? incidentDate : (DateTime?)null;
                var updated = new ClaimRecord
                {
                    Az = valueFor("AZ"),
                    PolicyNumber = valueFor("Policy Number"),
                    ClientId = valueFor("Client ID"),
                    Name = valueFor("Name"),
                    LoanAmount = parsedAmount,
                    LoanStartDate = parsedLoanDate,
                    IncidentDate = parsedIncidentDate,
                    Headers = _loadedClaim.Headers,
                    Values = values
                };
                if (string.IsNullOrWhiteSpace(updated.Az) ||
                    !string.Equals(updated.Az, _loadedClaim.Az, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The AZ value cannot be changed.");
                await Task.Run(() => _localExcelClaimsService.UpdateByAz(_loadedClaim.Az, updated));
                _loadedClaim = updated;
                TxtStatus.Text = "Claim saved to the Excel database.";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"Excel save failed: {ex.Message}";
            }
            finally
            {
                SetLoading(false);
            }
        }

        private void CalculateClaims()
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
