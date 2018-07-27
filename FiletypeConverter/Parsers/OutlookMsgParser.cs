using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using EAGetMail;
using FiletypeConverter.Interfaces;
using FiletypeConverter.ParsedContent;
using FiletypeConverter.Parsers;
using log4net;
using MsgReader.Mime;
using MsgReader.Mime.Header;
using MsgReader.Outlook;

namespace FiletypeConverter.Parsers
{
    public class OutlookMsgParser : FileParser
    {

        public OutlookMsgParser()
        {
            
        }

        public OutlookMsgParser(string path) :this()
        {
            Path = path;
        }

        public override bool Parse(string path)
        {
            return parseUsingMsgReader(path);
        }

        private bool parseUsingMsgReader(string path)
        {
            if (!File.Exists(path))
            {
                // TODO: logging
                return false;
            }

            MsgReader.Outlook.Storage.Message msg = null;
            try
            {
                msg = new MsgReader.Outlook.Storage.Message(path);
            }
            catch (Exception ex)
            {
                return false;
            }

            using (msg)
            {
                ParsedMsgMessage parsedMsgMessage = new ParsedMsgMessage();
                
                parsedMsgMessage.Sender = msg.Sender.DisplayName + "<" + msg.Sender.Email + ">";

                parsedMsgMessage.To_formatted = msg.GetEmailRecipients(Storage.Recipient.RecipientType.To, false, false);
                parsedMsgMessage.CC_formatted = msg.GetEmailRecipients(Storage.Recipient.RecipientType.Cc, false, false);
                parsedMsgMessage.Subject = msg.Subject;
                parsedMsgMessage.BodyHtml = msg.BodyHtml;
                parsedMsgMessage.BodyRtf = msg.BodyRtf;
                parsedMsgMessage.BodyText = msg.BodyText;

                parsedMsgMessage.AttachementNames_formatted = msg.GetAttachmentNames();
                
                //this.headers = msg.Headers;

                parsedMsgMessage.SentOn = msg.SentOn;
                parsedMsgMessage.ReceivedOn = msg.ReceivedOn;
                parsedMsgMessage.CreationTime = msg.CreationTime;
                parsedMsgMessage.LastModificationTime = msg.LastModificationTime;

                foreach (var msgAttachment in msg.Attachments)
                {
                    MsgReader.Outlook.Storage.Attachment ss;
                    
                    if (msgAttachment is MsgReader.Outlook.Storage.Attachment)
                    {
                        var attach = (MsgReader.Outlook.Storage.Attachment)msgAttachment;
                        parsedMsgMessage.AttachementNames.Add(attach.FileName);
                    }
                    else if (msgAttachment is MsgReader.Outlook.Storage.Message)
                    {
                        var attach = (MsgReader.Outlook.Storage.Message) msgAttachment;
                        parsedMsgMessage.AttachementNames.Add(attach.FileName);
                    }
                }
                //extractAndConvertAttachements(path);


                //                msgAsText = $@"
                //FROM: {from}
                //SENT ON: {sentOn}
                //TO: {recipientsTo}
                //CC: {recipientsCC}
                //SUBJECT: {subject}
                //HTMLBODY: {htmlBody}
                //RTFBODY: {rtfBody}
                //TXTBODY: {textBody}
                //ATTN: {attachementNames}
                //CREATIONTIME: {creationTime}
                //RECV_ON: {receivedOn}
                //MOD_DATE: {lastModificationTime}";
            }
            return true;
        }

        //public async void extractAndConvertAttachements(string origMsgPath, string destMsgPath)
        //{
        //    MsgReader.Reader reader = new MsgReader.Reader();

        //    string msgFile = origMsgPath;
        //    string attachementDestDir = destMsgPath + "_bijlages";

        //    if (!Directory.Exists(attachementDestDir))
        //    {
        //        Directory.CreateDirectory(attachementDestDir);
        //    }

        //    string[] outputFiles = reader.ExtractToFolder(msgFile, attachementDestDir);

        //    FileConverter.ConvertConfig convertConfig = new FileConverter.ConvertConfig()
        //    {
        //        ProcessOutlookMsg = true,
        //        ProcessWord = true,
        //        ProcessPowerpoint = true,
        //        ProcessExcel = true,
        //        ProcessImages = true,
        //        RootDir = attachementDestDir,
        //        OutputDir = attachementDestDir + "_pdf",
        //        Filter = "*",
        //    };


        //    FileConverter outlookFileConverter = new OutlookFileConverter(log);
        //    await outlookFileConverter.processInBackgroundAsync(convertConfig);

        //    FileConverter officeFileConverter = new OfficeFileConverter(log);

        //    await officeFileConverter.processInBackgroundAsync(convertConfig);

        //    FileConverter fileTransferrer = new ImageFileConverter(log);
        //    await fileTransferrer.processInBackgroundAsync(convertConfig);
        //}

        private bool parseUsingEAMail(string path)
        {
            this.Path = path;
            if (string.IsNullOrEmpty(Path))
            {
                throw new FileNotFoundException("Path not set");
            }

            if (File.Exists(Path) == false)
            {
                throw new FileNotFoundException($"File {Path} not found.");
            }

            ParsedMsgMessage parsedMsgMessage = new ParsedMsgMessage();

            Mail mail = new Mail("TryIt");
            //parsedMsgMessage.Mail = mail;

            try
            {
                mail.LoadOMSG(Path);
            }
            catch (Exception ex)
            {
                return false;
            }

            parsedMsgMessage.Sender = mail.From.ToString();
            parsedMsgMessage.To_formatted = "";

            foreach (var mailAddress in mail.To)
            {
                parsedMsgMessage.To.Add(mailAddress.ToString());
                parsedMsgMessage.To_formatted += mailAddress.ToString() + "; ";
            }

            foreach (var mailAddress in mail.Cc)
            {
                parsedMsgMessage.CC.Add(mailAddress.ToString());
                parsedMsgMessage.CC_formatted += mailAddress.ToString() + "; ";
            }

            parsedMsgMessage.Subject = mail.Subject;
            parsedMsgMessage.BodyText = mail.TextBody;
            parsedMsgMessage.BodyHtml = mail.HtmlBody;

            foreach (var mailAttachment in mail.Attachments)
            {
                parsedMsgMessage.AttachementNames.Add(mailAttachment.Name);
                parsedMsgMessage.AttachementNames_formatted += mailAttachment.Name + "; ";
            }

            return true;
        }

    }
}
