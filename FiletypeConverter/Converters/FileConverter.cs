using FiletypeConverter.Interfaces;
using FiletypeConverter.Utils;
using OfficeConverter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Converters
{
    public enum ConvertTarget
    {
        PDF,
        TXT,
    }

    [Flags]
    public enum FileType
    {
        NONE = 0,
        OUTLOOK_MSG = 1,
        OUTLOOK_PST = 2,
        WORD = 4,
        POWERPOINT = 8,
        EXCEL = 16,
        IMAGES = 64,
        PDF = 128,
    }

    public abstract class FileConverter : IFileConverter
    {
        protected IFileParser _fileParser;

        public FileConverter(IFileParser fileParser, IOutputSupplier outputSupplier)
        {
            _fileParser = fileParser;
            Output = outputSupplier;
        }

        public struct ConvertConfig
        {
            public FileType SourceFiles { get; set; }

            public string RootDir { get; set; }
            public string OutputDir { get; set; }
            public string Filter { get; set; }
            public bool KeepIntermediateFiles { get; set; }
            public bool ChangeDateTimes { get; set; }
            public DateTime? NewCreatedTime { get; set; }
            public DateTime? NewModifiedTime { get; set; }
        }

        public IOutputSupplier Output { get; set; } = new OutputSupplier();

        public IConverter Converter { get; set; } = new Converter();

        public ConvertTarget TargetFormat { get; set; } = ConvertTarget.PDF;

        public bool KeepIntermediateFiles { get; set; }
        public string Path { get; set; }
        public List<string> Journal { get; private set; } = new List<string>();
        public List<string> Errors { get; private set; } = new List<string>();
        
        public abstract Task ProcessInBackgroundAsync(ConvertConfig config);

        public void Dispatch(ConvertConfig convertConfig)
        {
            throw new NotImplementedException("Nog niet");
        }
    }
}
