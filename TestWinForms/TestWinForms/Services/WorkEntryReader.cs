using System.Collections.Generic;
using Crotating.Models;

namespace Crotating.Services
{
    public interface IWorkEntryReader
    {
        List<WorkEntry> ReadEntries(string filePath);
    }
}
