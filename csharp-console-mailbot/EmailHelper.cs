using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;

public class EmailHelper
{
    private static readonly string mailServer, login, password;
    private static readonly int port;
    private static readonly bool ssl;
    private static string projectPath;
    private static string assetsPath;
    private static string emailTemplatePath;
    static EmailHelper()
    {
        mailServer = ConfigurationManager.AppSettings["host_imap"];
        port = Convert.ToInt32(ConfigurationManager.AppSettings["port"]);
        ssl = true;
        login = ConfigurationManager.AppSettings["username"];
        password = ConfigurationManager.AppSettings["password"];
        projectPath = Path.Combine(Directory.GetCurrentDirectory().ToString(), @"..\..\");
        assetsPath = Path.Combine(projectPath, @"Data");
        emailTemplatePath = Path.GetFullPath(Path.Combine(assetsPath, @"EmailTemplate.html"));
    }

    public static void SendEmail(string sendTo)
    {
        MailAddress from = new MailAddress(ConfigurationManager.AppSettings["from"]);
        MailAddress to = new MailAddress(sendTo);
        MailMessage message = new MailMessage(from, to);
        message.Subject = ConfigurationManager.AppSettings["subject"];
        message.IsBodyHtml = true;
        StreamReader reader = new StreamReader(emailTemplatePath);
        message.Body = reader.ReadToEnd();
        reader.Close();

        SmtpClient client = new SmtpClient(Convert.ToString(ConfigurationManager.AppSettings["host_smtp"]))
        {
            Credentials = new NetworkCredential(ConfigurationManager.AppSettings["username"], ConfigurationManager.AppSettings["password"]),
            Port = Convert.ToInt32(ConfigurationManager.AppSettings["port_smtp"]),
            EnableSsl = true
        };

        try
        {
            client.Send(message);
        }
        catch (SmtpException ex)
        {
            Console.WriteLine(ex.Message.ToString());
        }
    }

    public static List<MimeMessage> GetUnreadMails(string subject = null)
    {
        List<MimeMessage> messages = new List<MimeMessage>();
        using (var client = new ImapClient())
        {
            client.Connect(mailServer, port, ssl);

            // Note: since we don't have an OAuth2 token, disable
            // the XOAUTH2 authentication mechanism.
            client.AuthenticationMechanisms.Remove("XOAUTH2");

            client.Authenticate(login, password);

            var inbox = client.Inbox;
            inbox.Open(FolderAccess.ReadWrite);
            SearchResults results;
            if (subject != null)
            {
                results = inbox.Search(SearchOptions.All, SearchQuery.SubjectContains(subject).And(SearchQuery.NotSeen));
            }
            else
            {
                results = inbox.Search(SearchOptions.All, SearchQuery.NotSeen);
            }
            foreach (var uniqueId in results.UniqueIds)
            {
                MimeMessage message = inbox.GetMessage(uniqueId);
                messages.Add(message);
                //inbox.AddFlags(uniqueId, MessageFlags.Seen, false);
            }

            client.Disconnect(true);
        }

        return messages;
    }

    public static IEnumerable<string> GetAllMails()
    {
        var messages = new List<string>();

        using (var client = new ImapClient())
        {
            client.Connect(mailServer, port, ssl);

            // Note: since we don't have an OAuth2 token, disable
            // the XOAUTH2 authentication mechanism.
            client.AuthenticationMechanisms.Remove("XOAUTH2");

            client.Authenticate(login, password);

            // The Inbox folder is always available on all IMAP servers...
            var inbox = client.Inbox;
            inbox.Open(FolderAccess.ReadOnly);
            var results = inbox.Search(SearchOptions.All, SearchQuery.All);
            foreach (var uniqueId in results.UniqueIds)
            {
                var message = inbox.GetMessage(uniqueId);

                messages.Add(message.HtmlBody);

                //Mark message as read
                //inbox.AddFlags(uniqueId, MessageFlags.Seen, true);
            }

            client.Disconnect(true);
        }

        return messages;
    }
}