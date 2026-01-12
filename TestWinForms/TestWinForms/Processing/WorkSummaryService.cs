using Crotating.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Crotating.Services
{
    public class WorkSummaryService
    {
        public DataTable BuildExportTable(List<DailySummary> summaries)
        {
            if (summaries == null || summaries.Count == 0)
                throw new InvalidOperationException("No data to export.");

            var table = new DataTable();

            // ---- Distinct dates become columns ----
            var dates = summaries
                .Select(s => s.Date)
                .Distinct()
                .OrderBy(d => d)
                .ToList();

            // First column: Name
            table.Columns.Add("Name", typeof(string));

            // Date columns
            foreach (var date in dates)
            {
                table.Columns.Add(
                    date.ToString("yyyy-MM-dd"),
                    typeof(double));
            }

            // ---- Group by person (row per name) ----
            var groupedByName = summaries.GroupBy(s => s.Name);

            foreach (var personGroup in groupedByName)
            {
                var row = table.NewRow();
                row["Name"] = personGroup.Key;

                foreach (var entry in personGroup)
                {
                    var columnName = entry.Date.ToString("yyyy-MM-dd");
                    row[columnName] = entry.TotalHours;
                }

                table.Rows.Add(row);
            }

            return table;
        }
    }
}
