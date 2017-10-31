using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace DutiesSendService
{
    public class EmailService
    {
        public SmtpConfiguration Smtp { get; set; }
        public EmailConfiguration Sender { get; set; }

        public void SendEmail(string toEmail, string subject, string body, string attachment)
        {
            var from = new MailAddress(Sender.email);
            var to = new MailAddress(toEmail);

            using (var message = new MailMessage(from, to))
            {
                message.Subject = subject;
                message.Body = body;
                message.IsBodyHtml = false;
                if (attachment != null)
                {
                    message.Attachments.Add(new Attachment(attachment));
                }

                using (var smtp = new SmtpClient(Smtp.host, Smtp.port))
                {
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = new NetworkCredential(Sender.email, Sender.pasword);
                    smtp.EnableSsl = true;

                    smtp.Send(message);
                }
            }
        }
    }
}
