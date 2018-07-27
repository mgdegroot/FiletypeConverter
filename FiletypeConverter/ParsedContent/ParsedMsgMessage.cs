using FiletypeConverter.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.ParsedContent
{
    public class ParsedMsgMessage : IParsedContent
    {
        public string Sender { get; set; }

        public string To_formatted { get; set; }
        public List<string> To { get; set; } = new List<string>();

        public string CC_formatted { get; set; }
        public List<string> CC { get; set; } = new List<string>();

        public string Subject { get; set; }

        public string BodyText { get; set; }

        public string BodyHtml { get; set; }

        public string BodyRtf { get; set; }

        public string AttachementNames_formatted { get; set; }
        public List<string> AttachementNames { get; set; } = new List<string>();


        public DateTime? CreationTime { get; set; }
        public DateTime? SentOn { get; set; }
        public DateTime? ReceivedOn { get; set; }

        public DateTime? LastModificationTime { get; set; }

        public string ContentAsString => $@"FROM: {Sender}
TO: {To_formatted}
CC: {CC_formatted}
SUBJECT: {Subject}
CREATED: {CreationTime}
SENT ON: {SentOn}
RECV ON: {ReceivedOn}
MOD DATE: {LastModificationTime}
ATTACHEMENTS: {AttachementNames_formatted}
TEXT: {BodyText}";



        public string FolderName => "<GEEN>";

        public string IdentifyingName => $"{Sender} - {Subject}";

        //private MessageHeader headers = null;
        //public Mail Mail { get; private set; }

    }
}
