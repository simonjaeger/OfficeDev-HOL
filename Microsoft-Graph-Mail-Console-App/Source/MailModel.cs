using System.Collections.Generic;

namespace Microsoft_Graph_Mail_Console_App
{
    public class MailModel
    {
        public MessageModel Message { get; set; }
    }

    public class MessageModel
    {
        public string Subject { get; set; }
        public BodyModel Body { get; set; }
        public List<ToRecipientModel> ToRecipients { get; set; }
    }

    public class BodyModel
    {
        public string ContentType { get; set; }
        public string Content { get; set; }
    }

    public class ToRecipientModel
    {
        public EmailAddressModel EmailAddress { get; set; }
    }

    public class EmailAddressModel
    {
        public string Address { get; set; }
    }
}
