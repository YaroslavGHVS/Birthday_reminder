using MailKit.Net.Smtp;
using MimeKit;
using System.Collections.Generic;


namespace MailCLient
{
    class GmailClient
    {
        private string senderName = "employees.birthdays.kpmg@gmail.com";
        private string subject = "🎉Birthday Notification";

        private string smtpServer = "smtp.gmail.com";
        private int port = 587;
        private bool useSsl = false;

        private string login = "employees.birthdays.kpmg@gmail.com";
        private string pass = "aaeufjpwvrtdogwf";

        private SmtpClient smtpInstance;

        public GmailClient()
        {
            smtpInstance = new SmtpClient();
            smtpInstance.Connect(smtpServer, port, useSsl);
            smtpInstance.Authenticate(login, pass);
        }

        ~GmailClient()
        {
            smtpInstance.Disconnect(true);
        }

        public void Send(List<string> receivers, string htmlText)
        {
            smtpInstance.Send( generateMail(receivers, htmlText) );
        }

        private MimeMessage generateMail(List<string> receivers, string htmlText)
        {
            var email = new MimeMessage();

            email.From.Add(new MailboxAddress("Sender Name", senderName));

            foreach(string adress in receivers)
                email.To.Add(new MailboxAddress("Receiver Name", adress));
            
            email.Subject = subject;
            email.Body = new TextPart(MimeKit.Text.TextFormat.Html)
            {
                Text = htmlText
            };

            return email;

        }
    }
    
}
