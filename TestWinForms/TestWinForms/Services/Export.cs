using Crotating.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Crotating.Services
{
    public class ExcelExporter
    {
        public void ExportSummary(
            IEnumerable<WorkEntry> entries,
            string outputPath)
        {
            if (entries == null || !entries.Any())
                throw new InvalidOperationException("No data to export.");

            // EPPlus license
            ExcelPackage.License.SetNonCommercialPersonal("Your Name or Organization");

            // ---- Prepare axes ----
            var minDate = entries.Min(e => e.Date.Date);
            var maxDate = entries.Max(e => e.Date.Date);

            var dates = new List<DateTime>();
            for (var d = minDate; d <= maxDate; d = d.AddDays(1))
            {
                dates.Add(d);
            }

            var names = entries
                .Select(e => e.Name)
                .Distinct()
                .OrderBy(n => n)
                .ToList();

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Summary");

                // ---- Header row ----
                ws.Cells[1, 1].Value = "Name";

                for (int col = 0; col < dates.Count; col++)
                {
                    ws.Cells[1, col + 2].Value = dates[col];
                    ws.Cells[1, col + 2].Style.Numberformat.Format = "mm/dd/yyyy";
                }

                // ---- Data rows ----
                for (int row = 0; row < names.Count; row++)
                {
                    string name = names[row];
                    ws.Cells[row + 2, 1].Value = name;

                    for (int col = 0; col < dates.Count; col++)
                    {
                        DateTime date = dates[col];

                        double totalHours = entries
                            .Where(e => e.Name == name && e.Date.Date == date)
                            .Sum(e => e.Hours);

                        if (totalHours > 0)
                            ws.Cells[row + 2, col + 2].Value = totalHours;
                    }
                }

                // ---- Formatting ----
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                ws.View.FreezePanes(2, 2);

                // ---- Save ----
                var file = new FileInfo(outputPath);
                if (file.Exists)
                    file.Delete();

                package.SaveAs(file);
            }
        }
    }
}
