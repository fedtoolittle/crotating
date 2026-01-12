using Crotating.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;

namespace Crotating.Services
{
    public class CrabalTimesheetReader : IWorkEntryReader
    {
        public List<WorkEntry> ReadEntries(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            var results = new List<WorkEntry>();

            // Set the license using the new EPPlus 8+ API
            OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("Your Name or Organization");

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                if (worksheet == null)
                    throw new InvalidDataException("No worksheet found in Excel file.");

                int lastRow = worksheet.Dimension.End.Row;
                string currentName = null;

                for (int row = 2; row <= lastRow; row++)
                {
                    object nameCell = worksheet.Cells[row, 1].Value;
                    object dateCell = worksheet.Cells[row, 2].Value;
                    object durationCell = worksheet.Cells[row, 3].Value;
                    object hoursCell = worksheet.Cells[row, 4].Value;

                    // ---- Skip TOTAL rows ----
                    if (dateCell != null &&
                        dateCell.ToString().Trim().Equals("Total", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }


                    // ---- Name (carry-forward) ----
                    if (nameCell != null && !string.IsNullOrWhiteSpace(nameCell.ToString()))
                    {
                        currentName = nameCell.ToString().Trim();
                        continue; // summary row → do not create WorkEntry
                    }

                    if (currentName == null)
                    {
                        throw new InvalidDataException(
                            "Name missing before data rows (row " + row + ")");
                    }

                    // ---- Date (detail rows only) ----
                    if (dateCell == null || string.IsNullOrWhiteSpace(dateCell.ToString()))
                    {
                        continue; // blank date → summary or spacer row
                    }

                    DateTime date;

                    // Excel numeric date (most common)
                    if (dateCell is double)
                    {
                        date = DateTime.FromOADate((double)dateCell);
                    }
                    // Already a DateTime
                    else if (dateCell is DateTime)
                    {
                        date = ((DateTime)dateCell);
                    }
                    // String fallback (MM/DD/YYYY)
                    else if (!DateTime.TryParseExact(
                        dateCell.ToString().Trim(),
                        "MM/dd/yyyy",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out date))
                    {
                        throw new InvalidDataException(
                            "Invalid date format at row " + row +
                            ": '" + dateCell + "'");
                    }

                    // ---- Duration (validated but not stored) ----
                    if (durationCell == null)
                        throw new InvalidDataException("Duration is empty at row " + row);

                    //Diagnostic for TimeSpan duration
                    var rawValue = durationCell;
                    var rawType = rawValue == null ? "null" : rawValue.GetType().FullName;

                    System.Diagnostics.Debug.WriteLine(
                        "Row " + row +
                        " | Duration raw value = [" + rawValue + "]" +
                        " | Type = " + rawType);

                    TimeSpan duration;

                    // Case 1: Excel numeric duration (fraction of a day)
                    if (durationCell is double)
                    {
                        duration = TimeSpan.FromDays((double)durationCell);
                    }
                    // Case 2: Excel DateTime (time value)
                    else if (durationCell is DateTime)
                    {
                        duration = ((DateTime)durationCell).TimeOfDay;
                    }
                    // Case 3: Text duration (HH:MM:SS, can exceed 24)
                    else
                    {
                        var text = durationCell.ToString().Trim();
                        var parts = text.Split(':');

                        if (parts.Length != 3)
                            throw new InvalidDataException(
                                "Invalid duration format at row " + row + ": '" + text + "'");

                        int durHours, durMinutes, durSeconds;

                        if (!int.TryParse(parts[0], out durHours) ||
                            !int.TryParse(parts[1], out durMinutes) ||
                            !int.TryParse(parts[2], out durSeconds))
                        {
                            throw new InvalidDataException(
                                "Invalid duration format at row " + row + ": '" + text + "'");
                        }

                        duration = new TimeSpan(durHours, durMinutes, durSeconds);

                    }

                    // ---- Hours (decimal) ----
                    if (hoursCell == null)
                        throw new InvalidDataException("Hours is empty at row " + row);

                    double hours;
                    if (!double.TryParse(
                        hoursCell.ToString().Trim(),
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out hours))
                    {
                        throw new InvalidDataException(
                            "Invalid hours value at row " + row +
                            ": '" + hoursCell + "'");
                    }

                    results.Add(new WorkEntry
                    {
                        Name = currentName,
                        Date = date.Date,
                        Hours = hours
                    });
                }
            }

            return results;
        }
    }
}
