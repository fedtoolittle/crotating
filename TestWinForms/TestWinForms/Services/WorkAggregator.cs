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
        public IList<DailySummary> AggregateByPersonAndDay(
            IList<WorkEntry> entries)
        {
            var map = new Dictionary<string, Dictionary<DateTime, decimal>>();

            foreach (var entry in entries)
            {
                var date = entry.StartTime.Date;

                if (!map.ContainsKey(entry.Name))
                    map[entry.Name] = new Dictionary<DateTime, decimal>();

                if (!map[entry.Name].ContainsKey(date))
                    map[entry.Name][date] = 0m;

                map[entry.Name][date] += entry.HoursWorked;
            }

            var results = new List<DailySummary>();

            foreach (var person in map)
            {
                foreach (var day in person.Value)
                {
                    results.Add(new DailySummary
                    {
                        Name = person.Key,
                        Date = day.Key,
                        TotalHours = day.Value
                    });
                }
            }

            return results;
        }
    }
}
