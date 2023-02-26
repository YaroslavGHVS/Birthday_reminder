namespace Birthday_reminder.EmailClient
{
    public interface IEmailClient
    {
        bool SendMail(string text, string receipients);
    }
}
