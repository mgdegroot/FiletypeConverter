using log4net;
using OfficeConverter;
using pst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter
{
    public class OutlookPstConverter : FileConverter, IFileConverter
    {
        private Converter converter = new Converter();

        public OutlookPstConverter(ILog log):base(log)
        {
            this.log = log;
        }

        public OutlookPstConverter(string path, ILog log) : this(log)
        {
            this.Path = path;
        }

        public class ParsedMessage
        {
            public string Subject { get; set; }
            public string Sender { get; set; }
            public List<string> Recipients { get; set; } = new List<string>();
            public List<string> AttachmentNames { get; set; } = new List<string>();
            public string Body { get; set; }
            public string FolderName { get; set; }
            public override string ToString()
            {
                string recipients = String.Join<string>(",", Recipients);
                string attachmentNames = String.Join<string>(",", AttachmentNames);
                return $@"[Subject: {Subject};
Sender: {Sender};
Recipients: {recipients};
Foldername: {FolderName};
Body: {Body};
AttachementNames: {attachmentNames}]";
            }

            public string MsgAsString => ToString().Replace("[", "").Replace("]", "");

        }

        public bool KeepMsgTxtFiles { get; set; } = false;
        public string Path { get; set; }
        public List<ParsedMessage> ParsedMessages { get; set; } = new List<ParsedMessage>();

        public List<string> Journal { get; private set; } = new List<string>();
        public List<string> Errors { get; private set; } = new List<string>();

        public override async Task processInBackgroundAsync(ConvertConfig config)
        {
            if (!Directory.Exists(config.OutputDir))
            {
                Directory.CreateDirectory(config.OutputDir);
            }

            if (config.ProcessOutlookPst)
            {
                throw new NotImplementedException("nog niet");
            }
        }

        public bool parse()
        {
            return parse(Path);
        }

        public bool parse(string path)
        {
            bool result = true;
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentNullException("Path is null or empty");
            }

            if (!File.Exists(path))
            {
                throw new FileNotFoundException("File not found.");
            }



            return parseUsingLibPST(path);
        }

        private bool parseUsingLibPST(string path)
        {
            bool result = true;

            //PSTFile pstFile = new PSTFile();
            PSTFile pst = PSTFile.Open(File.OpenRead(path));
            Folder rootFolder = pst.GetRootMailboxFolder();

            walkPstDir(rootFolder);
            //Message[] messages = rootFolder.GetMessages();


            return result;
        }


        private void walkPstDir(Folder folder)
        {

            Message[] messages = folder.GetMessages();
            string folderName = folder.GetProperty(MAPIProperties.PidTagDisplayName).Value.Value.ToUnicode();

            Folder[] subfolders = folder.GetSubFolders();

            foreach(Message message in messages)
            {
                processSingleMessage(folderName, message, "");

            }

            // recursively walk the hiearchy -->
            foreach(Folder subfolder in subfolders)
            {
                walkPstDir(subfolder);
            }
        }

        private void processSingleMessage(string folderName, Message inMessage, string outFile)
        {
            FileInfo nwFileInfo = new FileInfo(outFile);
            if (!nwFileInfo.Exists)
            {
                Directory.CreateDirectory(nwFileInfo.Directory.FullName);
            }

            string outFileTxt = outFile + ".txt",
                outFilePdf = outFile + ".pdf";

            //updateLogAndJournal($"Original: {inFile}. New: {outFile}", null);


            ParsedMessage parsedMessage = new ParsedMessage();


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
                        updateLogAndJournal(null, "recipient.GetProperty(): " + ex.Message, true);

                    }
                }
            }
            catch (Exception ex)
            {
                updateLogAndJournal(null, "message.GetRecipients(): " + ex.Message, true);
            }

            try
            {
                foreach (Attachment attachment in inMessage.GetAttachments())
                {
                    try
                    {
                        string attFilename = attachment.GetProperty(MAPIProperties.PidTagAttachLongFilename).Value.Value.ToUnicode();
                        parsedMessage.AttachmentNames.Add(attFilename);
                    }
                    catch (Exception ex)
                    {
                        log.Error(ex.Message);
                        Errors.Add(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                updateLogAndJournal(null, "message.GetAttachments(): " + ex.Message, true);
            }

            ParsedMessages.Add(parsedMessage);

            string result = parsedMessage.MsgAsString;
            File.WriteAllText(outFileTxt, result);
            converter.Convert(outFileTxt, outFilePdf);

            if (!KeepMsgTxtFiles)
            {
                File.Delete(outFileTxt);
            }

            //extractAndConvertAttachements(inFile, outFile);

        }
    }
}
