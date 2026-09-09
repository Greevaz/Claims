using System;
using System.ComponentModel;
using System.Globalization;

namespace ClaimsCalculatorApp
{
    public sealed class ClaimRecord
    {
        public string Az { get; init; } = "";
        public string PolicyNumber { get; init; } = "";
        public string ClientId { get; init; } = "";
        public string Name { get; init; } = "";
        public decimal? LoanAmount { get; init; }
        public DateTime? LoanStartDate { get; init; }
        public DateTime? IncidentDate { get; init; }
        public int RowIndex { get; init; }
        public string[] Headers { get; init; } = Array.Empty<string>();
        public object?[] Values { get; init; } = Array.Empty<object?>();
    }

    public sealed class ClaimField : INotifyPropertyChanged
    {
        private string _value;

        public string Name { get; }
        public string Value
        {
            get => _value;
            set
            {
                _value = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Value)));
            }
        }

        public ClaimField(string name, object? value)
        {
            Name = name;
            _value = FormatValue(name, value);
        }

        private static string FormatValue(string name, object? value)
        {
            if (value is null)
                return "";
            if (name.Contains("Date", StringComparison.OrdinalIgnoreCase))
            {
                if (value is double || value is float || value is decimal ||
                    value is int || value is long ||
                    double.TryParse(value.ToString(), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out _))
                {
                    var serial = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                    if (serial > 0)
                        return DateTime.FromOADate(serial).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                }
                if (value is DateTime date)
                    return date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            }
            if (name.Contains("Amount", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Balance", StringComparison.OrdinalIgnoreCase))
            {
                if (decimal.TryParse(value.ToString(), NumberStyles.Any,
                    CultureInfo.InvariantCulture, out var amount))
                    return amount.ToString("#,##0.00", CultureInfo.InvariantCulture);
            }
            return value.ToString() ?? "";
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    /// <summary>
    /// One entry in the "Rate Changes" table — a new interest rate that takes
    /// effect from EffectiveDate onward, until the next entry's date.
    /// </summary>
    public class RateChange : INotifyPropertyChanged
    {
        private DateTime _effectiveDate = DateTime.Today;
        private decimal _newRatePercent;

        public DateTime EffectiveDate
        {
            get => _effectiveDate;
            set { _effectiveDate = value; OnPropertyChanged(nameof(EffectiveDate)); }
        }

        public decimal NewRatePercent
        {
            get => _newRatePercent;
            set { _newRatePercent = value; OnPropertyChanged(nameof(NewRatePercent)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>
    /// One entry in the "Payments Received" table.
    /// </summary>
    public class PaymentEntry : INotifyPropertyChanged
    {
        private DateTime _date = DateTime.Today;
        private decimal _amount;

        public DateTime Date
        {
            get => _date;
            set { _date = value; OnPropertyChanged(nameof(Date)); }
        }

        public decimal Amount
        {
            get => _amount;
            set { _amount = value; OnPropertyChanged(nameof(Amount)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>
    /// One computed day in the interest schedule — mirrors a single row
    /// (columns A–H) of the "Template" sheet.
    /// </summary>
    public class DailyRow
    {
        public DateTime Date { get; set; }
        public decimal OpeningBalance { get; set; }
        public decimal DailyInterest { get; set; }
        public decimal InterestRatePercent { get; set; }
        public decimal? CompoundInterest { get; set; } // null unless this day compounds
        public decimal Payments { get; set; }
        public decimal Fees { get; set; }
        public decimal ClosingBalance { get; set; }

        public bool IsCompoundingDay => CompoundInterest.HasValue;
    }

    /// <summary>
    /// A single checkable document line for the claim package.
    /// </summary>
    public class DocumentCheckItem : INotifyPropertyChanged
    {
        private bool _isChecked;

        public string Name { get; }

        public bool IsChecked
        {
            get => _isChecked;
            set { _isChecked = value; OnPropertyChanged(nameof(IsChecked)); }
        }

        public DocumentCheckItem(string name, bool isChecked)
        {
            Name = name;
            _isChecked = isChecked;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
