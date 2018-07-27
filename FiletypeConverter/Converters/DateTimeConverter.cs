using FiletypeConverter.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FiletypeConverter.Converters
{
    public class DateTimeConverter : FileConverter
    {
        public DateTimeConverter(IFileParser fileParser, IOutputSupplier outputSupplier) : base(fileParser, outputSupplier)
        {
        }

        public override async Task ProcessInBackgroundAsync(ConvertConfig config)
        {
            if (config.ChangeDateTimes)
            {
                await processChangeDateTimes(path: config.OutputDir, filePattern: "*", dtCreated: config.NewCreatedTime, dtModified: config.NewModifiedTime);
            }
        }

        private async Task processChangeDateTimes(string path, string filePattern, DateTime? dtCreated, DateTime? dtModified)
        {
            await Task.Run(async () =>
            {
                List<string> matchingFiles = FileWalker.WalkDir(path, pattern: filePattern);

                foreach(var filename in matchingFiles)
                {
                    if (dtCreated.HasValue)
                    {
                        File.SetCreationTime(filename, dtCreated.Value);
                    }
                    if (dtModified.HasValue)
                    {
                        File.SetLastWriteTime(filename, dtModified.Value);
                    }
                }
            });
        }
    }
}
