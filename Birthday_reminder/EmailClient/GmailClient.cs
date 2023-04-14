using MailKit.Net.Smtp;
using MimeKit;
using System.Collections.Generic;
using Birthday_reminder;


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

        public void Send(List<Record> recordList, string htmlText)
        {
            smtpInstance.Send(generateMail(recordList, htmlText) );
        }

        private MimeMessage generateMail(List<Record> recordList, string htmlText)
        {
            var email = new MimeMessage();

            email.From.Add(new MailboxAddress("Birthday Reminder KPMG", senderName));

            foreach(Record record in recordList)
                email.To.Add(new MailboxAddress(record.Name, record.Email));
            
            email.Subject = subject;
            email.Body = new TextPart(MimeKit.Text.TextFormat.Html)
            {
                Text = htmlText
            };

            return email;

        }
    }
    
}
