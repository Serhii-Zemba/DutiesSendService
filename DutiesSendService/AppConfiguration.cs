using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DutiesSendService
{
    public class AppConfiguration
    {
        public List<string> errorReceivers { get; set; }
        public SmtpConfiguration smtp { get; set; }
        public EmailConfiguration sender { get; set; }
        public List<AccessAccountConfiguration> accessAccounts { get; set; }
        public List<SheetConfiguration> sheets { get; set; }
    }

    public class SmtpConfiguration
    {
        public string host { get; set; }
        public int port { get; set; }
    }

    public class EmailConfiguration
    {
        public string email { get; set; }
        public string pasword { get; set; }
    }

    public class AccessAccountConfiguration
    {
        public int id { get; set; }
        public string authTokenFolder { get; set; }
        public string applicationName { get; set; }
    }

    public class SheetConfiguration
    {
        public string name { get; set; }
        public string receiver { get; set; }
        public string sheetId { get; set; }
        public string sheetName { get; set; }
        public string tableHeader { get; set; }
        public int accessAccountId { get; set; }
        public string bodyAndSubjectStart { get; set; }
        public string fullName { get; set; }
    }
}
