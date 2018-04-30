using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;

namespace FiletypeConverter
{
    // TODO: rename and refactor to reflect actual functionality, namely
    // copy files without converting them.

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
                updateLogAndJournal("Converting Image files.", null);
                await processImageFiles(config.RootDir, config.OutputDir);
            }
        }


        private async Task processImageFiles(string rootPath, string outputPath)
        {
            string[] extensionsToMatch = { "jpg","jpeg", "png", "gif", "bmp", "pdf"};
            await Task.Run(async () => {
                FileAttributes attr = File.GetAttributes(rootPath);

                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    if (!outputPath.EndsWith(Path.DirectorySeparatorChar.ToString()))
                    {
                        outputPath += Path.DirectorySeparatorChar;
                    }
                    
                    List<string> matchingFiles = FileWalker.WalkDir(rootPath, extensionsToMatch, true);
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
            File.Copy(inFile, outFile, true);
        }
    }
}
