using System;
using System.ComponentModel;

namespace ClaimsCalculatorApp
{
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
