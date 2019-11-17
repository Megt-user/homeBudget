using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace homeBudget.Models
{
    public class LogEntry
    {
        [Key]
        public int Id { get; set; }

        [DataType(DataType.DateTime)]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy hh:mm:ss:fff }")]
        public DateTime Timestamp { get; set; }
        public string Message { get; set; }
        /// <summary>
        /// Duration of the current event in milliseconds
        /// </summary>
        public long? EventDuration { get; set; }
        public string Type { get; set; }
        public string Method { get; set; }


        public LogEntry()
        {

        }

        // TODO - Legg inn Type på alle LogEntries
        public LogEntry(string message, string method, string type )
        {
            Message = message;
            Method = method;
            Type = type;
            Timestamp = DateTime.Now;
        } 
        public LogEntry(string message, string method,long eventDuration, string type)
        {
            Message = message;
            Method = method;
            EventDuration = eventDuration;
            Type = type;
        }
    }
}
