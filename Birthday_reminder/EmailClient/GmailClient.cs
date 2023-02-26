using System;
using System.Net;
using System.Net.Mail;

namespace Birthday_reminder.EmailClient
{
/// <summary>
/// Класс, который реализует процесс отправки письма пользователю. 
/// </summary>
    public class GmailClient : IEmailClient
    {
        public string FromMail = "voinn.andrey@gmail.com";
        public string Subject = "Birthday Reminder";

        private string smtpServer = "smtp.gmail.com";
        private bool isSSLEnabled;
        private NetworkCredential credential;

        public GmailClient(string usename, string pass, bool isSSLEnabled = true)
        {
            credential = new NetworkCredential(usename, pass);
            this.isSSLEnabled = isSSLEnabled;
        }

        public bool SendMail(string text, string receipients)
        {
            try
            {
                var client = new SmtpClient(smtpServer, 587)
                {
                    Credentials = credential,
                    EnableSsl = this.isSSLEnabled
                };
                client.Send(this.FromMail, receipients, this.Subject, text);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

        }
    }

}
