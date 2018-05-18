using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter
{
    public abstract class FileConverter : IFileConverter
    {
        protected log4net.ILog log;

        public List<string> Journal { get; private set; } = new List<string>();
        public List<string> Errors { get; private set; } = new List<string>();

        public FileConverter(log4net.ILog log)
        {
            this.log = log;
        }

        public struct ConvertConfig
        {
            public bool ProcessOutlookMsg { get; set; }
            public bool ProcessWord { get; set; }
            public bool ProcessPowerpoint { get; set; }
            public bool ProcessExcel { get; set; }
            public bool ProcessImages { get; set; }
            public string RootDir { get; set; }
            public string OutputDir { get; set; }
            public string Filter { get; set; }
            public bool ProcessOutlookPst { get; set; }
        }

        public abstract Task processInBackgroundAsync(ConvertConfig config);

        public event EventHandler JournalEntryAdded;

        public virtual void OnJournalEntryAdded(EventArgs args)
        {
            EventHandler handler = JournalEntryAdded;
            if (handler != null)
            {
                handler(this, args);
            }
        }

        public event EventHandler ErrorEntryAdded;

        public virtual void OnErrorEntryAdded(EventArgs args)
        {
            ErrorEntryAdded?.Invoke(this, args);
        }

        public void Dispatch(ConvertConfig convertConfig)
        {
            throw new NotImplementedException("Nog niet");
            //if (convertConfig.ProcessWord || convertConfig.ProcessPowerpoint || convertConfig.ProcessExcel)
            //{
            //    OfficeFileConverter officeFileConverter = new OfficeFileConverter(log);
            //    officeFileConverter.processInBackgroundAsync(convertConfig);
            //}
        }

        protected void updateLogAndJournal(string outputText, string dbgText, bool isError = false)
        {
            if (!string.IsNullOrEmpty(outputText))
            {
                Journal.Add(outputText);
                log.Info(outputText);
                //synchronizationContext.Post(new SendOrPostCallback(o => { txtOutput.Text += (string)o + Environment.NewLine; txtOutput.ScrollToEnd(); }), outputText);
            }

            if (!string.IsNullOrEmpty(dbgText))
            {
                if (isError)
                {
                    Errors.Add($"ERROR - {dbgText}");
                    log.Error(dbgText);
                }
                else
                {
                    Errors.Add($"INFO - {dbgText}");
                    log.Info(dbgText);
                }
                //synchronizationContext.Post(new SendOrPostCallback(o => { txtDebug.Text += (string)o + Environment.NewLine; txtDebug.ScrollToEnd(); }), dbgText);
            }
        }
    }
}
