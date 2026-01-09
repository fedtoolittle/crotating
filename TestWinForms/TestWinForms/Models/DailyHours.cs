using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Crotating.Models
{
    public class DailyHours
    {
        public string Name { get; set; } = string.Empty;
        public DateTime Date { get; set; }
        public decimal Hours { get; set; }
    }
}
