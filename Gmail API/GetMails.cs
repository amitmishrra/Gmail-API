using GmailAPI.APIHelper;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace Gmail_API
{
    [TestClass]
    public class GetMails
    {
        [TestMethod]
        public void GetMail()
        {
            string mail = ConfigurationManager.AppSettings["HostAddress"];
            GmailService GmailService = GmailAPIHelper.GetService();
            List<Gmail> EmailList = new List<Gmail>();
            UsersResource.MessagesResource.ListRequest ListRequest = GmailService.Users.Messages.List(mail);
            ListRequest.LabelIds = "INBOX";
            ListRequest.IncludeSpamTrash = false;
            ListRequest.Q = "is:unread";

            //GET ALL EMAILS
            ListMessagesResponse ListResponse = ListRequest.Execute();

            Console.WriteLine("Length : " + ListResponse.Messages.Count);
            foreach (Message Msg in ListResponse.Messages)
            {
                UsersResource.MessagesResource.GetRequest Message = GmailService.Users.Messages.Get(mail, Msg.Id);
                Message MsgContent = Message.Execute();

                string MailBody = string.Empty;

                if (MsgContent.Payload.Parts == null && MsgContent.Payload.Body != null)
                {
                    // If there are no parts, get the body data directly
                    MailBody = MsgContent.Payload.Body.Data;
                }
                else if (MsgContent.Payload.Parts != null)
                {
                    // Extract the data from the parts
                    MailBody = GmailAPIHelper.MsgNestedParts(MsgContent.Payload.Parts);
                }

                // Decode the base64-encoded body
                string ReadableText = DecodeBase64String(MailBody);
                Console.WriteLine(ReadableText);
            }
        }

        // Decode the Base64 string
        private string DecodeBase64String(string base64Input)
        {
            if (string.IsNullOrEmpty(base64Input))
                return string.Empty;

            // Gmail API Base64 URL encoding replacements
            base64Input = base64Input.Replace("-", "+").Replace("_", "/");

            // Ensure the string length is a multiple of 4
            switch (base64Input.Length % 4)
            {
                case 2: base64Input += "=="; break;
                case 3: base64Input += "="; break;
            }

            byte[] base64EncodedBytes = Convert.FromBase64String(base64Input);
            return Encoding.UTF8.GetString(base64EncodedBytes);
        }
    }
}
