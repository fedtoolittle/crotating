using Crotating.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace Crotating.Services
{
    public class CrabalTimecardReader : IWorkEntryReader
    {
        public List<WorkEntry> ReadEntries(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            ExcelPackage.License.SetNonCommercialPersonal("Crotating");

            var results = new List<WorkEntry>();

            using var package = new ExcelPackage(new FileInfo(filePath));

            var ws = package.Workbook.Worksheets.Count > 0
                ? package.Workbook.Worksheets[0]
                : throw new InvalidDataException("Workbook contains no worksheets.");

            int lastRow = ws.Dimension?.End.Row ?? 0;
            string currentName = null;

            // Start at row 2 (skip header)
            for (int row = 2; row <= lastRow; row++)
            {
                var nameCell = ws.Cells[row, 1].Value;
                var dateCell = ws.Cells[row, 2].Value;
                var hoursCell = ws.Cells[row, 4].Value;

                // ---- Carry-forward name ----
                if (!string.IsNullOrWhiteSpace(nameCell?.ToString()))
                {
                    currentName = nameCell.ToString().Trim();
                }

                if (string.IsNullOrWhiteSpace(currentName))
                    continue;

                // ---- Date ----
                if (!TryGetDate(dateCell, out var date))
                    continue;

                // ---- Hours ----
                var hours = CellNormalization.GetDecimalOrZero(hoursCell);

                results.Add(new WorkEntry
                {
                    Name = currentName,
                    Date = date,
                    Hours = hours
                });
            }

            return results;
        }

        private static bool TryGetDate(object value, out DateTime date)
        {
            date = default;

            if (value == null)
                return false;

            if (value is DateTime dt)
            {
                date = dt.Date;
                return true;
            }

            if (value is double oa)
            {
                date = DateTime.FromOADate(oa).Date;
                return true;
            }
            if (value is decimal dec)
            {
                date = DateTime.FromOADate((double)dec).Date;
                return true;
            }

            return DateTime.TryParseExact(
                value.ToString().Trim(),
                "MM/dd/yyyy",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out date);
        }

    }
}
