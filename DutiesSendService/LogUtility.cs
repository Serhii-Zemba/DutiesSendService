using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DutiesSendService
{
    public class LogUtility
    {
        public string BuildExceptionMessage(Exception ex)
        {
            var exception = ex;
            while (exception.InnerException != null)
            {
                exception = exception.InnerException;
            }

            var exceptionMsg = Environment.NewLine + "Message :" + exception.Message;
            
            exceptionMsg += Environment.NewLine + "Source :" + exception.Source;
            
            exceptionMsg += Environment.NewLine + "Stack Trace :" + exception.StackTrace;

            return exceptionMsg;
        }
    }
}
