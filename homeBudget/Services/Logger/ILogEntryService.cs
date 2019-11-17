using homeBudget.Models;

namespace homeBudget.Services.Logger
{
    public interface ILogEntryService
    {
        void Save(LogEntry logEntry);
    }
}