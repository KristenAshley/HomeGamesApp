using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Configuration;
using System.Net;

namespace HomeGames
{
    public class alternateEmail
    {
        public string Username { get; set; }
        public string Password { get; set; }

        public alternateEmail(string username, string password)
        {
            Username = username;
            Password = password;
        }

        public void Send(MailMessage msg)
        {
            SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["host"].ToString(), Convert.ToInt32(ConfigurationManager.AppSettings["port"]));
            client.EnableSsl = true;
            client.Timeout = 10000;
            client.UseDefaultCredentials = false;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.Credentials = new NetworkCredential(Username, Password);
            client.Send(msg);
        }
    }
}
