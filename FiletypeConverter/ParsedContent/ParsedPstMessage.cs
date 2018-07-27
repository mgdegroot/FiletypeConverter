using FiletypeConverter.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.ParsedContent
{
    

    public class ParsedPstMessage : IParsedContent
    {
        public string Subject { get; set; }
        public string Sender { get; set; }
        public List<string> Recipients { get; set; } = new List<string>();
        public List<string> AttachmentNames { get; set; } = new List<string>();
        public string Body { get; set; }
        public string FolderName { get; set; }
        public string IdentifyingName => $"{Sender} - {Subject}";
        public DateTime CreationTime { get; set; }
        public DateTime ModificationTime { get; set; }

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

        public string ContentAsString => ToString().Replace("[", "").Replace("]", "").Replace(";\r\n", "\r\n");

    }
}
