using FiletypeConverter.Interfaces;
using FiletypeConverter.Utils;
using OfficeConverter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter
{
    public enum ConvertTarget
    {
        PDF,
        TXT,
    }

    public abstract class FileConverter : IFileConverter
    {
        public FileConverter()
        {
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
            public bool KeepIntermediateFiles { get; set; }
        }

        public IOutputSupplier Output { get; set; } = new OutputSupplier();

        public IConverter Converter { get; set; } = new Converter();

        public ConvertTarget TargetFormat { get; set; } = ConvertTarget.PDF;

        public bool KeepIntermediateFiles { get; set; }
        public string Path { get; set; }
        public List<string> Journal { get; private set; } = new List<string>();
        public List<string> Errors { get; private set; } = new List<string>();
        
        public event OutputAdded ErrorAdded;
        public event OutputAdded JournalAdded;


        public abstract Task processInBackgroundAsync(ConvertConfig config);


        //public event EventHandler JournalEntryAdded;

        //public virtual void OnJournalEntryAdded(EventArgs args)
        //{
        //    EventHandler handler = JournalEntryAdded;
        //    if (handler != null)
        //    {
        //        handler(this, args);
        //    }
        //}

        //public event EventHandler ErrorEntryAdded;

        //public virtual void OnErrorEntryAdded(EventArgs args)
        //{
        //    ErrorEntryAdded?.Invoke(this, args);
        //}

        public void Dispatch(ConvertConfig convertConfig)
        {
            throw new NotImplementedException("Nog niet");
            //if (convertConfig.ProcessWord || convertConfig.ProcessPowerpoint || convertConfig.ProcessExcel)
            //{
            //    OfficeFileConverter officeFileConverter = new OfficeFileConverter(log);
            //    officeFileConverter.processInBackgroundAsync(convertConfig);
            //}
        }

        //protected void updateLogAndJournal(string outputText, string dbgText, bool isError = false)
        //{
        //    if (!string.IsNullOrEmpty(outputText))
        //    {
        //        Journal.Add(outputText);
        //        log.Info(outputText);
        //        JournalAdded?.Invoke(outputText);
        //        //synchronizationContext.Post(new SendOrPostCallback(o => { txtOutput.Text += (string)o + Environment.NewLine; txtOutput.ScrollToEnd(); }), outputText);
        //    }

        //    if (!string.IsNullOrEmpty(dbgText))
        //    {
        //        if (isError)
        //        {
        //            Errors.Add($"ERROR - {dbgText}");
        //            log.Error(dbgText);
        //        }
        //        else
        //        {
        //            Errors.Add($"INFO - {dbgText}");
        //            log.Info(dbgText);
        //        }
        //        ErrorAdded?.Invoke(dbgText);
        //        //synchronizationContext.Post(new SendOrPostCallback(o => { txtDebug.Text += (string)o + Environment.NewLine; txtDebug.ScrollToEnd(); }), dbgText);
        //    }
        //}
    }
}
