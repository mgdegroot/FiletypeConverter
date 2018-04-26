using log4net;
using OfficeConverter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter
{
    public class OfficeFileConverter : FileConverter, IFileConverter
    {

        private Converter converter = new Converter();

        public OfficeFileConverter(ILog log) : base(log)
        {
        }

        public override async Task processInBackgroundAsync(ConvertConfig config)
        {
            if (!Directory.Exists(config.OutputDir))
            {
                Directory.CreateDirectory(config.OutputDir);
            }

            if (config.ProcessWord)
            {
                updateLogAndJournal("Converting Word documents.", null);
                await processOfficeFiles(config.RootDir, config.OutputDir, "*.doc?");
            }

            if (config.ProcessPowerpoint)
            {
                updateLogAndJournal("Converting Powerpoint documents.", null);
                await processOfficeFiles(config.RootDir, config.OutputDir, "*.ppt?");
            }

            if (config.ProcessExcel)
            {
                updateLogAndJournal("Converting Excel documents.", null);
                await processOfficeFiles(config.RootDir, config.OutputDir, "*.xls?");
            }
        }

        private async Task processOfficeFiles(string rootPath, string outputPath, string extension)
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

                    List<string> matchingFiles = FileWalker.WalkDir(rootPath, extension, true);

                    foreach (var filename in matchingFiles)
                    {
                        updateLogAndJournal($"Found matching file: {filename}", null);

                        string nwFilename = filename.Replace(rootPath, outputPath) + ".pdf";
                        processSingleOfficeFile(filename, nwFilename);
                    }
                }
                else
                {
                    updateLogAndJournal($"Processing single file: {rootPath}", null);
                    string nwFilename = rootPath.Replace(rootPath, outputPath) + ".pdf";
                    processSingleOfficeFile(rootPath, nwFilename);
                }


            });
        }

        private void processSingleOfficeFile(string filename, string nwFilename)
        {

            FileInfo nwFileInfo = new FileInfo(nwFilename);

            if (!nwFileInfo.Exists)
            {
                Directory.CreateDirectory(nwFileInfo.Directory.FullName);
            }

            updateLogAndJournal($"Converting {filename} to {nwFilename}", null);

            try
            {
                converter.Convert(filename, nwFilename);
            }
            catch (Exception ex)
            {

                updateLogAndJournal($"Conversion failed for file {filename}", $"ERROR: {filename}: {ex.Message}", true);
            }
        }
    }
}
