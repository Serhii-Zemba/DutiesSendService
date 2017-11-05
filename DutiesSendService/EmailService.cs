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

        public void SendEmail(string toEmail, string subject, string body, string attachment, EmailType type)
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

                    switch (type)
                    {
                            case EmailType.Duty:
                                smtp.Credentials = new NetworkCredential(Sender.email, Sender.pasword);
                            break;
                            case EmailType.Error:
                                smtp.Credentials = new NetworkCredential("andrey.arshavintut@gmail.com", "Qwerty_12345");
                            break;
                    }
                    
                    smtp.EnableSsl = true;

                    smtp.Send(message);
                }
            }
        }

        public enum EmailType
        {
            Duty,
            Error
        }
    }
}
