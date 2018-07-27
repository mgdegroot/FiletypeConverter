using FiletypeConverter.Interfaces;
using FiletypeConverter.Utils;
using log4net;
using OfficeConverter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Converters
{
    public class OfficeFileConverter : FileConverter, IFileConverter
    {

        public OfficeFileConverter(IFileParser fileParser, IOutputSupplier outputSupplier) : base(fileParser, outputSupplier)
        {
            
        }

        public override async Task ProcessInBackgroundAsync(ConvertConfig config)
        {
            if (!Directory.Exists(config.OutputDir))
            {
                Directory.CreateDirectory(config.OutputDir);
            }

            List<string> filePatternsToProcess = new List<string>();

            if ((config.SourceFiles | FileType.WORD) == FileType.WORD)
            {
                filePatternsToProcess.AddRange(Util.FileExtensions[FileType.WORD]);
            }

            if ((config.SourceFiles | FileType.POWERPOINT) == FileType.POWERPOINT)
            {
                filePatternsToProcess.AddRange(Util.FileExtensions[FileType.POWERPOINT]);
            }

            if ((config.SourceFiles | FileType.EXCEL) == FileType.EXCEL)
            {
                filePatternsToProcess.AddRange(Util.FileExtensions[FileType.EXCEL]);
            }

            await processOfficeFiles(config.RootDir, config.OutputDir, filePatternsToProcess.ToArray<string>());
        }

        private async Task processOfficeFiles(string rootPath, string outputPath, string[] filePatterns)
        {
            // TODO: refactor so all background work is done in one thread instead of multiple async methods-->
            await Task.Run(async () =>
            {
                FileAttributes attr = File.GetAttributes(rootPath);

                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    if (!outputPath.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
                    {
                        outputPath += System.IO.Path.DirectorySeparatorChar;
                    }

                    List<string> matchingFiles = FileWalker.WalkDir(rootPath, filePatterns, true);

                    foreach (var filename in matchingFiles)
                    {
                        Output.AddJournalEntry($"Found matching file: {filename}");

                        string nwFilename = filename.Replace(rootPath, outputPath) + ".pdf";
                        processSingleOfficeFile(filename, nwFilename);
                    }
                }
                else
                {
                    Output.AddJournalEntry($"Processing single file: {rootPath}");
                    string nwFilename = rootPath.Replace(rootPath, outputPath) + ".pdf";
                    processSingleOfficeFile(rootPath, nwFilename);
                }


            });
        }

        private void processSingleOfficeFile(string filename, string nwFilename)
        {
            Output.AddJournalEntry($"Converting {filename} to {nwFilename}");

            FileInfo nwFileInfo = new FileInfo(nwFilename);

            if (!nwFileInfo.Exists)
            {
                Directory.CreateDirectory(nwFileInfo.Directory.FullName);
            }

            
            try
            {
                // no parsing needed, convert directly -->
                Converter.ConvertFromCom(filename, nwFilename);
            }
            catch (Exception ex)
            {
                Output.AddJournalAndLog($"Conversion failed for file {filename}", $"ERROR: {filename}: {ex.Message}", true);
            }
        }
    }
}
