﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FiletypeConverter.Interfaces;
using log4net;

namespace FiletypeConverter.Converters
{
    // TODO: rename and refactor to reflect actual functionality, namely
    // copy files without converting them.

    public class ImageFileConverter : FileConverter
    {
        public ImageFileConverter(IFileParser fileParser, IOutputSupplier outputSupplier) : base(fileParser, outputSupplier)
        {
        }

        public override async Task ProcessInBackgroundAsync(ConvertConfig config)
        {
            if (!Directory.Exists(config.OutputDir))
            {
                Directory.CreateDirectory(config.OutputDir);
            }

            Output.AddJournalEntry("Converting Image files.");
            await processImageFiles(config.RootDir, config.OutputDir);
        }


        private async Task processImageFiles(string rootPath, string outputPath)
        {
            string[] extensionsToMatch = { "*.jpg","*.jpeg", "*.png", "*.gif", "*.bmp", "*.pdf" };
            await Task.Run(async () => {
                FileAttributes attr = File.GetAttributes(rootPath);

                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    if (!outputPath.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
                    {
                        outputPath += System.IO.Path.DirectorySeparatorChar;
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
            FileInfo nwFileInfo = new FileInfo(outFile);
            try
            {
                if (!nwFileInfo.Exists)
                {
                    Directory.CreateDirectory(nwFileInfo.Directory.FullName);
                }
            }
            catch(Exception ex)
            {
                Output.AddLogEntry($"Exception creating dirctory: {ex.Message}", true);
            }

            try
            {
                File.Copy(inFile, outFile, true);
            }
            catch(Exception ex)
            {
                Output.AddLogEntry($"Exception copying file {inFile} to {outFile}: {ex.Message}", true);
            }
        }
    }
}
