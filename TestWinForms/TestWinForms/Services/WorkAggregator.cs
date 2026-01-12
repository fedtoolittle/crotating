using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Crotating.Models;

namespace Crotating.Services
{
    public class WorkAggregator
    {
        public List<DailySummary> AggregateByPersonAndDay(
            List<WorkEntry> entries)
        {
            return entries
                .GroupBy(e => new { e.Name, e.Date })
                .Select(g => new DailySummary
                {
                    Name = g.Key.Name,
                    Date = g.Key.Date,
                    TotalHours = g.Sum(x => x.Hours)
                })
                .ToList();
        }
    }
}
