﻿using FiletypeConverter.Interfaces;
using FiletypeConverter.ParsedContent;
using pst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Parsers
{
    public class OutlookPstParser : FileParser, IFileParser
    {
        public OutlookPstParser()
        {

        }

        public List<ParsedPstMessage> ParsedMessages { get; set; } = new List<ParsedPstMessage>();

        public override string ContentAsString => throw new NotImplementedException();

        public override bool Parse()
        {
            return Parse(Path);
        }

        public override bool Parse(string path)
        {
            bool retVal = true;
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentNullException("Path must be set");
            }
            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"File at {path} not found.");
            }

            retVal = processSingleOutlookPstFile(path);


            return retVal;
        }


        ///////////////////////////////////////////////////////////////////

        private bool processSingleOutlookPstFile(string path)
        {
            bool result = true;

            PSTFile pst = PSTFile.Open(File.OpenRead(path));
            Folder rootFolder = pst.GetRootMailboxFolder();

            walkPstDir(rootFolder);

            return result;
        }


        private void walkPstDir(Folder folder)
        {

            Message[] messages = folder.GetMessages();
            string folderName = folder.GetProperty(MAPIProperties.PidTagDisplayName).Value.Value.ToUnicode();

            //nwPath = nwPath + System.IO.Path.DirectorySeparatorChar.ToString() + folderName;
            Folder[] subfolders = folder.GetSubFolders();

            foreach (Message message in messages)
            {
                processSingleMessage(folderName, message);

            }

            // recursively walk the hiearchy -->
            foreach (Folder subfolder in subfolders)
            {
                //string subdirPath = nwPath + System.IO.Path.DirectorySeparatorChar.ToString() + folderName;
                walkPstDir(subfolder);
            }
        }

        private void processSingleMessage(string folderName, Message inMessage)
        {
            //DirectoryInfo nwDirInfo = new DirectoryInfo(outDir);
            ////FileInfo nwFileInfo = new FileInfo(outFile);
            //if (!nwDirInfo.Exists)
            //{
            //    Directory.CreateDirectory(nwDirInfo.FullName);
            //}


            //updateLogAndJournal($"Original: {inFile}. New: {outFile}", null);


            ParsedPstMessage parsedMessage = new ParsedPstMessage()
            {
                FolderName = folderName,
            };


            try
            {
                parsedMessage.Subject = inMessage.GetProperty(MAPIProperties.PidTagSubject).Value.Value.ToUnicode().Trim();
            }
            catch { }
            try
            {
                parsedMessage.Sender = inMessage.GetProperty(MAPIProperties.PidTagSenderEmailAddress).Value.Value.ToUnicode().Trim();
            }
            catch { }
            try
            {
                parsedMessage.Body = inMessage.GetProperty(MAPIProperties.PidTagBody).Value.Value.ToUnicode().Trim();
            }
            catch { }

            try
            {
                foreach (Recipient recipient in inMessage.GetRecipients())
                {
                    try
                    {
                        string recipientAddress = recipient.GetProperty(MAPIProperties.PidTagEmailAddress).Value.Value.ToUnicode();
                        parsedMessage.Recipients.Add(recipientAddress);
                    }
                    catch (Exception ex)
                    {
                        Output.AddLogEntry("recipient.GetProperty(): " + ex.Message, true);

                    }
                }
            }
            catch (Exception ex)
            {
                Output.AddLogEntry("message.GetRecipients(): " + ex.Message, true);
            }

            // TODO: reenable fetching of attachements when rest is working properly -->
            //try
            //{
            //    foreach (Attachment attachment in inMessage.GetAttachments())
            //    {
            //        try
            //        {
            //            string attFilename = attachment.GetProperty(MAPIProperties.PidTagAttachLongFilename).Value.Value.ToUnicode();
            //            parsedMessage.AttachmentNames.Add(attFilename);
            //        }
            //        catch (Exception ex)
            //        {
            //            log.Error(ex.Message);
            //            Errors.Add(ex.Message);
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    updateLogAndJournal(null, "message.GetAttachments(): " + ex.Message, true);
            //}

            ParsedMessages.Add(parsedMessage);
            //string filePart = $"{parsedMessage.Sender} - {parsedMessage.Subject}";
            //foreach (char invalidChar in System.IO.Path.GetInvalidFileNameChars())
            //{
            //    filePart = filePart.Replace(invalidChar, '_');
            //}
            //string outFile = $"{outDir}\\{filePart}";



            ////outFile = $"{outDir}\\{msgCounter++.ToString()}";
            //string outFileTxt = outFile + ".txt";
            //string outFilePdf = outFile + ".pdf";


            //string result = parsedMessage.MsgAsString;
            //File.WriteAllText(outFileTxt, result);
            //// TODO: reenable after test -->
            //converter.Convert(outFileTxt, outFilePdf);

            //if (!KeepIntermediateFiles)
            //{
            //    File.Delete(outFileTxt);
            //}

            ////extractAndConvertAttachements(inFile, outFile);

        }
    }
}
