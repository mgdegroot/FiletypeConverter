using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;

namespace FiletypeConverter
{
    public class ImageFileConverter : FileConverter, IFileConverter
    {
        public ImageFileConverter(ILog log) : base(log)
        {
        }

        public override async Task processInBackgroundAsync(ConvertConfig config)
        {
            if (!Directory.Exists(config.OutputDir))
            {
                Directory.CreateDirectory(config.OutputDir);
            }

            if (config.ProcessImages)
            {
                updateLogAndJournal("Converting Outlook message files.", null);
                await processImageFiles(config.RootDir, config.OutputDir);
            }
        }


        private async Task processImageFiles(string rootPath, string outputPath)
        {
            await Task.Run(async () => {
                FileAttributes attr = File.GetAttributes(rootPath);

                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    if (!outputPath.EndsWith(Path.DirectorySeparatorChar.ToString()))
                    {
                        outputPath += Path.DirectorySeparatorChar;
                    }

                    List<string> matchingFiles = FileWalker.WalkDir(rootPath, "*.msg", true);
                    foreach (var filename in matchingFiles)
                    {
                        string nwFilename = filename.Replace(rootPath, outputPath);

                        processSingleImageFile(filename, nwFilename);
                    }
                }
                else
                {
                    string nwFilename = rootPath.Replace(rootPath, outputPath);
                    processSingleImageFile(rootPath, nwFilename);
                }
            });
        }

        private void processSingleImageFile(string inFile, string outFile)
        {
            throw new NotImplementedException("Nog niet");
        }
    }
}
