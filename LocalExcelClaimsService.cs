using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace ClaimsCalculatorApp;

public sealed class LocalExcelClaimsService
{
    private readonly string _workbookPath;
    private readonly string _tableName;

    public LocalExcelClaimsService(GraphSettings settings)
    {
        _workbookPath = settings.LocalWorkbookPath;
        _tableName = settings.TableName;
    }

    public ClaimRecord? FindByAz(string az)
        => Find("AZ", az);

    public ClaimRecord? Find(string fieldName, string query)
    {
        if (string.IsNullOrWhiteSpace(query))
            throw new ArgumentException("A search value is required.", nameof(query));

        EnsureWorkbookExists();
        dynamic? excel = null;
        dynamic? workbook = null;
        dynamic? sheet = null;
        dynamic? table = null;
        dynamic? headersRange = null;
        dynamic? dataRange = null;

        try
        {
            excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("Microsoft Excel is not installed."));
            excel.Visible = false;
            excel.DisplayAlerts = false;
            workbook = excel.Workbooks.Open(_workbookPath, ReadOnly: true);
            var tableInfo = FindTable((object)workbook);
            sheet = tableInfo.Sheet;
            table = tableInfo.Table;
            headersRange = table.HeaderRowRange;
            dataRange = table.DataBodyRange;
            return FindRecord(headersRange!.Value2, dataRange?.Value2, fieldName, query);
        }
        finally
        {
            CloseExcel(excel, workbook, sheet, table, headersRange, dataRange);
        }
    }

    public void UpdateByAz(string az, ClaimRecord updatedRecord)
    {
        if (string.IsNullOrWhiteSpace(az))
            throw new ArgumentException("AZ is required.", nameof(az));

        EnsureWorkbookExists();
        dynamic? excel = null;
        dynamic? workbook = null;
        dynamic? sheet = null;
        dynamic? table = null;
        dynamic? headersRange = null;
        dynamic? dataRange = null;

        try
        {
            excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("Microsoft Excel is not installed."));
            excel.Visible = false;
            excel.DisplayAlerts = false;
            workbook = excel.Workbooks.Open(_workbookPath, ReadOnly: false);
            var tableInfo = FindTable((object)workbook);
            sheet = tableInfo.Sheet;
            table = tableInfo.Table;
            headersRange = table.HeaderRowRange;
            dataRange = table.DataBodyRange;

            var headers = ToRow(headersRange!.Value2);
            if (dataRange is null)
                throw new InvalidOperationException($"The {_tableName} table has no data rows.");

            var values = dataRange.Value2;
            var rowIndex = FindRowIndex(headers, values, az);
            if (rowIndex < 0)
                throw new InvalidOperationException($"No claim found for AZ '{az}'.");

            for (var column = 0; column < headers.Length; column++)
            {
                var value = column < updatedRecord.Values.Length ? updatedRecord.Values[column] : null;
                dataRange.Cells[rowIndex + 1, column + 1].Value2 = ToExcelValue(value);
            }
            workbook.Save();
        }
        finally
        {
            CloseExcel(excel, workbook, sheet, table, headersRange, dataRange);
        }
    }

    private (object Sheet, object Table) FindTable(object workbookObject)
    {
        dynamic workbook = workbookObject;
        for (int sheetIndex = 1; sheetIndex <= workbook.Worksheets.Count; sheetIndex++)
        {
            dynamic candidateSheet = workbook.Worksheets[sheetIndex];
            for (int tableIndex = 1; tableIndex <= candidateSheet.ListObjects.Count; tableIndex++)
            {
                dynamic candidateTable = candidateSheet.ListObjects[tableIndex];
                if (string.Equals((string)candidateTable.Name, _tableName, StringComparison.OrdinalIgnoreCase))
                    return (candidateSheet, candidateTable);
                Release(candidateTable);
            }
            Release(candidateSheet);
        }

        throw new InvalidOperationException($"Excel table '{_tableName}' was not found.");
    }

    private static ClaimRecord? FindRecord(object? headerValues, object? dataValues,
        string fieldName, string query)
    {
        var headers = ToRow(headerValues);
        var searchColumn = Array.FindIndex(headers, h =>
            string.Equals(h, fieldName, StringComparison.OrdinalIgnoreCase));
        if (searchColumn < 0)
            throw new InvalidOperationException($"The ClaimsData table must contain a '{fieldName}' column.");

        var rowIndex = FindRowIndex(headers, dataValues, searchColumn, query);
        if (rowIndex < 0)
            return null;

        var row = ToRow(dataValues, rowIndex);
        object? Value(string name) => GetValue(headers, row, name);
        return new ClaimRecord
        {
            Az = Value("AZ")?.ToString() ?? "",
            PolicyNumber = Value("Policy Number")?.ToString() ?? "",
            ClientId = Value("Client ID")?.ToString() ?? "",
            Name = Value("Name")?.ToString() ?? "",
            LoanAmount = ParseDecimal(Value("Loan Amount")),
            LoanStartDate = ParseDate(Value("Loan Start Date")),
            IncidentDate = ParseDate(Value("Incident Date")),
            RowIndex = rowIndex,
            Headers = headers,
            Values = row
        };
    }

    private static int FindRowIndex(string[] headers, object? dataValues, string az)
    {
        var azColumn = Array.FindIndex(headers, h => string.Equals(h, "AZ", StringComparison.OrdinalIgnoreCase));
        if (azColumn < 0)
            throw new InvalidOperationException("The ClaimsData table must contain an AZ column.");

        return FindRowIndex(headers, dataValues, azColumn, az);
    }

    private static int FindRowIndex(string[] headers, object? dataValues, int searchColumn, string query)
    {
        var rowCount = GetDimension(dataValues, 0);
        for (var row = 0; row < rowCount; row++)
        {
            var value = GetArrayValue(dataValues, row, searchColumn)?.ToString()?.Trim() ?? "";
            var matches = searchColumn == Array.FindIndex(headers, h =>
                string.Equals(h, "AZ", StringComparison.OrdinalIgnoreCase))
                ? string.Equals(value, query.Trim(), StringComparison.OrdinalIgnoreCase)
                : value.Contains(query.Trim(), StringComparison.OrdinalIgnoreCase);
            if (matches)
                return row;
        }
        return -1;
    }

    private static void SetCell(dynamic dataRange, int row, string[] headers, string header, object? value)
    {
        var column = Array.FindIndex(headers, h => string.Equals(h, header, StringComparison.OrdinalIgnoreCase));
        if (column >= 0)
            dataRange.Cells[row + 1, column + 1].Value2 = ToExcelValue(value);
    }

    private static object? ToExcelValue(object? value)
    {
        if (value is DateTime date)
            return date.ToOADate();
        if (value is string text &&
            DateTime.TryParseExact(text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var parsedDate))
            return parsedDate.ToOADate();
        return value;
    }

    private static object? GetValue(string[] headers, object?[] row, string header)
    {
        var column = Array.FindIndex(headers, h => string.Equals(h, header, StringComparison.OrdinalIgnoreCase));
        return column >= 0 && column < row.Length ? row[column] : null;
    }

    private static string[] ToRow(object? values, int row = 0) =>
        Enumerable.Range(0, GetDimension(values, 1)).Select(column => GetArrayValue(values, row, column)?.ToString() ?? "").ToArray();

    private static int GetDimension(object? values, int dimension)
    {
        if (values is not Array array)
            return values is null ? 0 : 1;
        return array.GetLength(dimension);
    }

    private static object? GetArrayValue(object? values, int row, int column)
    {
        if (values is object[,] matrix)
            return matrix[row + 1, column + 1];
        return row == 0 && column == 0 ? values : null;
    }

    private static decimal? ParseDecimal(object? value) =>
        decimal.TryParse(value?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)
            ? parsed : null;

    private static DateTime? ParseDate(object? value)
    {
        if (value is null)
            return null;

        // Excel Value2 returns date cells as OLE Automation numbers.
        if (value is double serial || value is float || value is decimal || value is int || value is long)
        {
            var number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (number > 0)
                return DateTime.FromOADate(number).Date;
        }

        var text = value.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(text))
            return null;

        // Table rows are normalized to strings before mapping, so preserve
        // Excel serial dates in that representation as well.
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var textSerial) &&
            textSerial > 0)
            return DateTime.FromOADate(textSerial).Date;

        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed) ||
            DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out parsed))
            return parsed.Date;

        return null;
    }

    private void EnsureWorkbookExists()
    {
        if (!File.Exists(_workbookPath))
            throw new FileNotFoundException("The local OneDrive workbook was not found.", _workbookPath);
    }

    private static void CloseExcel(dynamic? excel, dynamic? workbook, params dynamic?[] objects)
    {
        foreach (var item in objects.Reverse())
            Release(item);
        if (workbook is not null)
        {
            try { workbook.Close(false); } catch (COMException) { }
            Release(workbook);
        }
        if (excel is not null)
        {
            try { excel.Quit(); } catch (COMException) { }
            Release(excel);
        }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
            Marshal.FinalReleaseComObject(value);
    }
}
