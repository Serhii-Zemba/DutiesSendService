using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace DutiesSendService
{
    public class Worker
    {
        private AppConfiguration config;

        public void ReadSettings()
        {
            using (var stream = new StreamReader("app_settings.json"))
            {
                var json = stream.ReadToEnd();
                config = JsonConvert.DeserializeObject<AppConfiguration>(json);
            }
        }

        public void Start()
        {
            while (true)
            {
                var emailService = new EmailService
                {
                    Smtp = config.smtp,
                    Sender = config.sender
                };

                try
                {
                    var pdf = new PDFSharp();

                    foreach (var sheet in config.sheets)
                    {
                        var accessAccount = config.accessAccounts.FirstOrDefault(x => x.id == sheet.accessAccountId);

                        var sheetsApi = new GoogleSheets
                        {
                            AccessAccount = accessAccount
                        };
                        
                        var newDuties = sheetsApi.GetNewDuties(sheet.sheetId, sheet.sheetName);

                        var dirManager = new DirectoryManager();
                        var tempDirectory = dirManager.CreateTemporaryDutiesDirectory();

                        foreach (var duty in newDuties)
                        {
                            var document = pdf.CreateDutyPDF(duty, sheet.tableHeader);
                            var pdfDuty =
                                pdf.SaveDutyPDFLocally(document, tempDirectory, $"{duty.PhoneLine}_{duty.Service}");

                            var bodyAndSubject = string.Empty;

                            switch (sheet.name)
                            {
                                case "tks":
                                    bodyAndSubject = "ТКС";
                                    break;
                                case "tis":
                                    bodyAndSubject = "ТИС";
                                    break;
                                case "b2c":
                                    bodyAndSubject = "ЮКОМ";
                                    break;
                            }

                            bodyAndSubject +=
                                $" {duty.PhoneLine}; {duty.CustomerAddress}; {duty.Service}; {duty.NotesFromUkt}; Включить до {duty.ExpirationDate}";

                            emailService.SendEmail(sheet.receiver, bodyAndSubject, bodyAndSubject, pdfDuty);

                            sheetsApi.UpdateDutyStatus(sheet.sheetId, sheet.sheetName, $"V{duty.RowNumber}");
                        }

                        dirManager.RemoveTemporaryDutiesDirectory(tempDirectory);
                    }

                    Thread.Sleep(900000);
                }
                catch (Exception ex)
                {
                    var exceptionMsg = new LogUtility().BuildExceptionMessage(ex);
                    emailService.SendEmail(config.errorReceiver, "Google Sheets App Error", exceptionMsg, null);
                }
            }
        } 
    }
}
