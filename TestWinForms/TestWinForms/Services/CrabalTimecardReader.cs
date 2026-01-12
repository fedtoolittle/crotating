using ClosedXML.Excel;
using Crotating.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Crotating.Services
{
    public class CrabalTimecardReader : IWorkEntryReader
    {
        public List<WorkEntry> ReadEntries(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            var results = new List<WorkEntry>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);

                // Skip header row
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    try
                    {
                        if (row.Cells().Any(c =>
                         c.GetString().Trim()
                         .Equals("TOTAL", StringComparison.OrdinalIgnoreCase)))
                        {
                            continue;
                        }

                        var name = row.Cell(1).GetString().Trim();
                        var startCell = row.Cell(3);
                        var endCell = row.Cell(4);
                        var hoursCell = row.Cell(6);

                        if (string.IsNullOrWhiteSpace(name))
                            throw new InvalidDataException(
                            "Name is missing at row " + row.RowNumber() +
                            " | Raw value: [" + row.Cell(1).Value + "]");


                        // ---- Infer date ----
                        DateTime startTime = DateTime.MinValue;
                        DateTime endTime = DateTime.MinValue;

                        bool hasStart = TryGetDateTime(startCell, out startTime);
                        bool hasEnd = TryGetDateTime(endCell, out endTime);

                        if (!hasStart && !hasEnd)
                        {
                            throw new InvalidDataException(
                                "Start and End time are both missing or invalid.");
                        }

                        DateTime date = hasStart
                            ? startTime.Date
                            : endTime.Date;


                        // ---- Hours ----
                        double hours;
                        if (!hoursCell.TryGetValue(out hours))
                        {
                            if (!double.TryParse(
                                hoursCell.GetString(),
                                NumberStyles.Any,
                                CultureInfo.InvariantCulture,
                                out hours))
                            {
                                throw new InvalidDataException("Invalid hours value.");
                            }
                        }

                        results.Add(new WorkEntry
                        {
                            Name = name,
                            Date = date,
                            Hours = hours
                        });
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            "Error parsing row " + row.RowNumber(), ex);
                    }
                }
            }

            return results;
        }

        private bool TryGetDateTime(IXLCell cell, out DateTime value)
        {
            value = default;

            if (cell == null || cell.IsEmpty())
                return false;

            if (cell.TryGetValue(out value))
                return true;

            return DateTime.TryParse(
                cell.GetString(),
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out value);
        }
    }
}
