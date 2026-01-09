using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Crotating.Models;

namespace Crotating.Services
{
    public class ExcelReader
    {

        private static DateTime ParseExcelDate(object cellValue, int rowIndex, string columnName)
{
    if (cellValue == null)
        throw new InvalidDataException(
            columnName + " is empty at row " + rowIndex);

    // Case 1: Excel numeric date
    if (cellValue is double)
    {
        return DateTime.FromOADate((double)cellValue);
    }

    // Case 2: String date
    var text = cellValue.ToString().Trim();

    DateTime dt;
    if (DateTime.TryParse(text, out dt))
    {
        return dt;
    }

    // Case 3: Exact format fallback
    if (DateTime.TryParseExact(
        text,
        "yyyy-MM-dd HH:mm",
        System.Globalization.CultureInfo.InvariantCulture,
        System.Globalization.DateTimeStyles.None,
        out dt))
    {
        return dt;
    }

    throw new InvalidDataException(
        "Invalid " + columnName + " format at row " + rowIndex + ": '" + text + "'");
}

        public IList<WorkEntry> ReadWorkEntries(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path is empty.");

            var results = new List<WorkEntry>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();

                bool isHeader = true;

                foreach (var row in rows)
                {
                    if (isHeader)
                    {
                        isHeader = false;
                        continue;
                    }

                    try
                    {
                        var entry = ParseRow(row);
                        results.Add(entry);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            "Error parsing Excel row " + row.RowNumber(),
                            ex);
                    }
                }
            }

            return results;
        }

        private WorkEntry ParseRow(IXLRow row)
        {
            var name = row.Cell(1).GetString().Trim();
            var startRaw = row.Cell(3).GetString();
            var endRaw = row.Cell(4).GetString();
            var hoursRaw = row.Cell(6).GetString();

            if (string.IsNullOrEmpty(name))
                throw new InvalidDataException("Name is missing.");

            DateTime startTime;
            DateTime endTime;
            decimal hoursWorked;

            if (!DateTime.TryParseExact(
                    startRaw,
                    "yyyy-MM-dd HH:mm",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out startTime))
                throw new InvalidDataException("Invalid start time format.");

            if (!DateTime.TryParseExact(
                    endRaw,
                    "yyyy-MM-dd HH:mm",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out endTime))
                throw new InvalidDataException("Invalid end time format.");

            if (!decimal.TryParse(
                    hoursRaw,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out hoursWorked))
                throw new InvalidDataException("Invalid hours worked.");

            if (endTime < startTime)
                throw new InvalidDataException("End time is before start time.");

            return new WorkEntry
            {
                Name = name,
                StartTime = startTime,
                EndTime = endTime,
                HoursWorked = hoursWorked
            };
        }
    }
}
