using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace DutiesSendService
{
    public class Duty
    {
        public int RowNumber { get; set; }
        public string CustomerName { get; set; }
        public string PhoneLine { get; set; }
        public string TaskId { get; set; }
        public string SubscriberLineType { get; set; }
        public string Account { get; set; }
        public string Service { get; set; }
        public string ADSLType { get; set; }
        public string ADSLPort { get; set; }
        public string B6Network { get; set; }
        public string Region { get; set; }
        public string City { get; set; }
        public string Term { get; set; }
        public string CustomerAddress { get; set; }
        public string ContactName { get; set; }
        public string ContactPhone { get; set; }
        public string DutyCreator { get; set; }
        public string CreationDate { get; set; }
        public string ExpirationDate { get; set; }
        public string LinearDataForOTA { get; set; }
        public string NotesFromUkt { get; set; }
        public string Status { get; set; }
    }
}
