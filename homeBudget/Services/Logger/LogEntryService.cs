using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using homeBudget.Models;

namespace homeBudget.Services.Logger
{
    public class LogEntryService : ILogEntryService
    {
        //private ApplicationDbContext _dbContext;
        public void Save(LogEntry logEntry)
        {
            logEntry.Timestamp = DateTime.Now;
            System.Diagnostics.Debug.WriteLine($" message: {logEntry.Message}, Type: {logEntry.Type}, duration in mms:{logEntry.EventDuration}");
        }

    }
}
