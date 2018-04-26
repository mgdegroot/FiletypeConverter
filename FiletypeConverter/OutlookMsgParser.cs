﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using EAGetMail;
using MsgReader.Mime;
using MsgReader.Mime.Header;
using MsgReader.Outlook;

namespace FiletypeConverter
{
    public class OutlookMsgParser
    {
        public string Path { get; private set; }

        public string From { get; private set; }

        public string To_formatted { get; private set; }
        public List<string> To { get; private set; } = new List<string>();

        public string CC_formatted { get; private set; }
        public List<string> CC { get; private set; } = new List<string>();

        public string Subject { get; private set; }

        public string BodyText { get; private set; }

        public string BodyHtml { get; private set; }

        public string BodyRtf { get; private set; }

        public string AttachementNames_formatted { get; private set; }
        public List<string> AttachementNames { get; private set; } = new List<string>();

        public Mail Mail { get; private set; }

        public DateTime? CreationTime { get; private set; }
        public DateTime? SentOn { get; private set; }
        public DateTime? ReceivedOn { get; private set; }

        public DateTime? LastModificationTime { get; private set; }

        public string MsgAsString => $@"FROM: {From}
TO: {To_formatted}
CC: {CC_formatted}
SUBJECT: {Subject}
CREATED: {CreationTime}
SENT ON: {SentOn}
RECV ON: {ReceivedOn}
MOD DATE: {LastModificationTime}
ATTACHEMENTS: {AttachementNames_formatted}
TEXT: {BodyText}";

        public string AsString => MsgAsString;


        private MessageHeader headers = null;

        public OutlookMsgParser()
        {
        }

        public OutlookMsgParser(string path):this()
        {
            Path = path;
        }

        public bool parse()
        {
            if (string.IsNullOrEmpty(Path))
            {
                throw new ArgumentException("Path not set");
            }
            return parse(Path);
        }

        public bool parse(string path)
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
                From = msg.Sender.DisplayName + "<" + msg.Sender.Email + ">";
                
                To_formatted = msg.GetEmailRecipients(Storage.Recipient.RecipientType.To, false, false);
                CC_formatted = msg.GetEmailRecipients(Storage.Recipient.RecipientType.Cc, false, false);
                Subject = msg.Subject;
                BodyHtml = msg.BodyHtml;
                BodyRtf = msg.BodyRtf;
                BodyText = msg.BodyText;
                
                AttachementNames_formatted = msg.GetAttachmentNames();
                
                this.headers = msg.Headers;

                SentOn = msg.SentOn;
                ReceivedOn = msg.ReceivedOn;
                CreationTime = msg.CreationTime;
                LastModificationTime = msg.LastModificationTime;

                foreach (var msgAttachment in msg.Attachments)
                {
                    MsgReader.Outlook.Storage.Attachment ss;
                    if (msgAttachment is MsgReader.Outlook.Storage.Attachment)
                    {
                        var attach = (MsgReader.Outlook.Storage.Attachment)msgAttachment;
                        AttachementNames.Add(attach.FileName);
                    }
                    else if (msgAttachment is MsgReader.Outlook.Storage.Message)
                    {
                        var attach = (MsgReader.Outlook.Storage.Message) msgAttachment;
                        AttachementNames.Add(attach.FileName);
                    }
                    
                }


                

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

            Mail mail = new Mail("TryIt");
            this.Mail = mail;

            try
            {
                mail.LoadOMSG(Path);
            }
            catch (Exception ex)
            {
                return false;
            }

            From = mail.From.ToString();
            To_formatted = "";

            foreach (var mailAddress in mail.To)
            {
                To.Add(mailAddress.ToString());
                To_formatted += mailAddress.ToString() + "; ";
            }

            foreach (var mailAddress in mail.Cc)
            {
                CC.Add(mailAddress.ToString());
                CC_formatted += mailAddress.ToString() + "; ";
            }

            Subject = mail.Subject;
            BodyText = mail.TextBody;
            BodyHtml = mail.HtmlBody;

            foreach (var mailAttachment in mail.Attachments)
            {
                AttachementNames.Add(mailAttachment.Name);
                AttachementNames_formatted += mailAttachment.Name + "; ";
            }

            return true;
        }

    }
}
