using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Interfaces
{
    public delegate void OutputAdded(string message);

    public interface IOutputSupplier
    {
        List<string> Journal { get; }
        List<string> Errors { get; }
        event OutputAdded ErrorAdded;
        event OutputAdded JournalAdded;

        void AddJournalEntry(string text);
        void AddLogEntry(string text, bool isError = false);
        void AddJournalAndLog(string journalText, string logText, bool isError = false);
    }
}
