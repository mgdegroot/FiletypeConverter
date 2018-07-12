using FiletypeConverter.Interfaces;
using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Utils
{
    public class OutputSupplier : IOutputSupplier
    {
        public ILog Log { get; set; } = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public List<string> Journal { get; } = new List<string>();
        public List<string> Errors { get; } = new List<string>();

        public event OutputAdded ErrorAdded;
        public event OutputAdded JournalAdded;

        public void AddJournalEntry(string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                Journal.Add(text);
                Log.Info(text);
                JournalAdded?.Invoke(text);
            }
        }

        public void AddLogEntry(string text, bool isError = false)
        {
            if (!string.IsNullOrEmpty(text))
            {
                if (isError)
                {
                    Errors.Add($"ERROR - {text}");
                    Log.Error(text);
                }
                else
                {
                    Errors.Add($"INFO - {text}");
                    Log.Info(text);
                }
                ErrorAdded?.Invoke(text);
            }
        }

        public void AddJournalAndLog(string journalText, string logText, bool isError = false)
        {
            if (!string.IsNullOrEmpty(journalText))
            {
                Journal.Add(journalText);
                Log.Info(journalText);
                JournalAdded?.Invoke(journalText);
            }

            if (!string.IsNullOrEmpty(logText))
            {
                if (isError)
                {
                    Errors.Add($"ERROR - {logText}");
                    Log.Error(logText);
                }
                else
                {
                    Errors.Add($"INFO - {logText}");
                    Log.Info(logText);
                }
                ErrorAdded?.Invoke(logText);
            }
        }
    }
}
