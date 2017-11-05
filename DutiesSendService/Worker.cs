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
            var errorCount = 0;

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

                            var notesForSubject = duty.NotesFromUkt.Length > 60
                                ? duty.NotesFromUkt.Substring(0, 60)
                                : duty.NotesFromUkt;
                            var customerId = $"{duty.PhoneLine}";
                            if (duty.PhoneLine.Trim().ToLower() == "нс")
                            {
                                customerId += $" Л/С {duty.Account}";
                            }
                            var subject =
                                $"{sheet.bodyAndSubjectStart} {customerId}; {duty.CustomerAddress}; {duty.Service}; {notesForSubject}; Включить до {duty.ExpirationDate}";

                            var body =
                                $"{sheet.bodyAndSubjectStart} {customerId}; {duty.CustomerAddress}; {duty.Service}; {duty.NotesFromUkt}; Включить до {duty.ExpirationDate}";

                            emailService.SendEmail(sheet.receiver, subject, body, pdfDuty, EmailService.EmailType.Duty);

                            sheetsApi.UpdateDutyStatus(sheet.sheetId, sheet.sheetName, $"V{duty.RowNumber}");
                        }

                        dirManager.RemoveTemporaryDutiesDirectory(tempDirectory);
                    }

                    Thread.Sleep(900000);
                }
                catch (Exception ex)
                {
                    var exceptionMsg = new LogUtility().BuildExceptionMessage(ex);

                    foreach (var receiver in config.errorReceivers)
                    {
                        emailService.SendEmail(receiver, "Google Sheets App Error", exceptionMsg, null, EmailService.EmailType.Error);
                    }

                    errorCount++;
                    if (errorCount == 10)
                    {
                        throw ex;
                    }
                }
            }
        } 
    }
}
