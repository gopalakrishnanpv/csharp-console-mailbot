using System.Configuration;
using System.Collections.Generic;
using MimeKit;
using MimeKit.Text;

namespace csharp_console_mailbot
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadEmail();
        }
        public static void SendEmail()
        {
            string toAddress = ConfigurationManager.AppSettings["to"];
            EmailHelper.SendEmail(toAddress);
        }

        public static void ReadEmail()
        {
            List< MimeMessage> emails = EmailHelper.GetUnreadMails();
            foreach (MimeMessage email in emails)
            {
                System.Console.WriteLine(email.Subject);
                System.Console.WriteLine(email.GetTextBody(TextFormat.Plain));
                System.Console.WriteLine(email.HtmlBody);
            }
            
        }
    }
}
