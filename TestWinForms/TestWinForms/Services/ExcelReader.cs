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
    if (cellValue is double d)
    {
        return DateTime.FromOADate(d);
    }

    // Case 2: String date
    var text = cellValue.ToString().Trim();

            if (DateTime.TryParse(text, out DateTime dt))
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

                    var name = row.Cell(1).GetString().Trim();
                    if (string.IsNullOrEmpty(name))
                        continue;

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
    }
}
