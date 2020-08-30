using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;

namespace DutiesSendService
{
    public class PDFSharp
    {
        public Document CreateDutyPDF(Duty duty, string tableHeader)
        {
            var document = new Document();
            var pageSetup = document.DefaultPageSetup.Clone();
            pageSetup.PageFormat = PageFormat.A4;
            pageSetup.Orientation = Orientation.Portrait;
            pageSetup.LeftMargin = "2cm";
            pageSetup.TopMargin = "1.5cm";

            var section = document.AddSection();
            section.PageSetup = pageSetup;

            var header = section.AddParagraph("Наряд на виконання робіт");
            header.Format.Font.Name = "Times New Roman";
            header.Format.Font.Size = "12pt";
            header.Format.Font.Bold = true;
            header.Format.Alignment = ParagraphAlignment.Center;

            section.AddParagraph();

            var table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "11pt");
            table.Format.Font.Bold = true;
            table.AddColumn("14cm");
            table.AddColumn("4cm");

            var row = table.AddRow();

            row.Cells[0].AddParagraph($"№ {duty.TaskId}\tвід {duty.CreationDate}");

            row.Cells[1].AddParagraph(tableHeader);
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("5cm");
            table.AddColumn("13cm");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Абонент");

            row.Cells[1].AddParagraph($"{duty.CustomerName}");
            row.Cells[1].Format.Font.Bold = true;

            row = table.AddRow();

            row.Cells[0].AddParagraph("Особовий рахунок");

            row.Cells[1].AddParagraph($"{duty.Account}");

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("5cm");
            table.AddColumn("4cm");
            table.AddColumn("4.5cm");
            table.AddColumn("4.5cm");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Договір");

            row.Cells[2].AddParagraph("Юридичний статус");

            row.Cells[3].AddParagraph("Фіз.особа");
            row.Cells[3].Format.Font.Italic = true;

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("5cm");
            table.AddColumn("13cm");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Юридична адреса");

            row.Cells[1].AddParagraph($"{duty.CustomerAddress}");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Адреса підключення");

            row.Cells[1].AddParagraph($"{duty.CustomerAddress}");
            row.Cells[1].Format.Font.Bold = true;

            row = table.AddRow();

            row.Cells[0].AddParagraph("П.І.Б. контактної особи");

            row.Cells[1].AddParagraph($"{duty.ContactName}");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Контактні телефони");

            row.Cells[1].AddParagraph($"{duty.ContactPhone}");

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("5cm");
            table.AddColumn("4cm");
            table.AddColumn("4.5cm");
            table.AddColumn("4.5cm");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Абонентське обладнання");

            row.Cells[1].AddParagraph("□ ADSL-модем");

            row.Cells[2].AddParagraph("□ STB");

            row.Cells[3].AddParagraph("□");

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("5cm");
            table.AddColumn("13cm");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Дані абонентської  лінії для підключення ОТА");

            row.Cells[1].AddParagraph($"{duty.LinearDataForOTA}");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Примітки");

            row.Cells[1].AddParagraph($"{duty.NotesFromUkt}");
            row.Cells[1].Format.Font.Italic = true;

            row = table.AddRow();

            row.Cells[0].AddParagraph("Примітки для ADSL");

            row.Cells[1].AddParagraph($"Абон.лінія: {duty.PhoneLine}, Тип лінії: {duty.SubscriberLineType}, Тип підключ ADSL: {duty.ADSLType}");
            row.Cells[1].Format.Font.Italic = true;

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "11pt");
            table.AddColumn("11.5cm");
            table.AddColumn("6.5cm");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Найменування робіт з підключення Абонента до телекомунікаційної мережі ПАТ «Укртелеком»:");
            row.Cells[0].Format.Font.Bold = true;

            row.Cells[1].AddParagraph("Дата, час виконання");

            row = table.AddRow();

            var paragraph = row.Cells[0].AddParagraph();
            paragraph.Format.Font.Size = "12pt";
            paragraph = row.Cells[0].AddParagraph($"Підключення до послуги «{duty.Service}»");
            paragraph.Format.Font.Size = "12pt";
            paragraph.Format.Font.Bold = true;
            paragraph.Format.Font.Italic = true;

            paragraph = row.Cells[0].AddParagraph("Примітки:");
            paragraph.Format.Font.Size = "10pt";
            paragraph.Format.Font.Italic = true;
            paragraph.Format.Font.Underline = Underline.Single;
            var text = paragraph.AddFormattedText($" Включить не позже {duty.ExpirationDate}");
            text.Font.Size = "12pt";
            text.Font.Bold = true;
            paragraph = row.Cells[0].AddParagraph();
            paragraph.Format.Font.Size = "10pt";

            paragraph = row.Cells[1].AddParagraph("____.____.____");
            paragraph.Format.Font.Size = "14pt";
            paragraph = row.Cells[1].AddParagraph("____:____");
            paragraph.Format.Font.Size = "14pt";
            paragraph = row.Cells[1].AddParagraph("узгоджена з Абонентом");
            paragraph.Format.Font.Size = "10pt";
            paragraph.Format.Font.Italic = true;

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "11pt");
            table.AddColumn("5.5cm");
            table.AddColumn("5.5cm");
            table.AddColumn("7cm");

            row = table.AddRow();
            row.BottomPadding = new Unit(2);

            row.Cells[0].AddParagraph("Наряд видав");
            paragraph = row.Cells[0].AddParagraph("(ПІБ співробітника ПАТ «Укртелеком»)");
            paragraph.Format.Font.Size = "8pt";
            paragraph.Format.Font.Italic = true;
            paragraph.Format.Font.Underline = Underline.Single;

            row.Cells[1].AddParagraph($"{duty.DutyCreator}");

            paragraph = row.Cells[2].AddParagraph($"Дата: {duty.CreationDate} ____:_____");
            paragraph.Format.Font.Size = "10pt";

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("1.5cm");
            table.AddColumn("7cm");
            table.AddColumn("2.5cm");
            table.AddColumn("2.5cm");
            table.AddColumn("4.5cm");

            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.BottomPadding = new Unit(5);

            row.Cells[0].AddParagraph("Статус");

            row.Cells[1].AddParagraph("Найменування робіт");

            row.Cells[2].AddParagraph("Виконавець");

            row.Cells[3].AddParagraph("Відповідальни");
            paragraph = row.Cells[3].AddParagraph("й");
            text = paragraph.AddFormattedText(" (ПІБ)");
            text.Font.Bold = false;

            row.Cells[4].AddParagraph("Дата : Час виконання");

            row = table.AddRow();
            row.Cells[1].AddParagraph("Реконструкція лінії, встановлення, підключення та налаштування абонентського обладнання");

            row = table.AddRow();
            row.Cells[1].AddParagraph("Реконструкция існуючої абонентської лінії (заміна або перевлаштування проводки та\\або схеми підключення)");

            table.AddRow();

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("5cm");
            table.AddColumn("6cm");
            table.AddColumn("7cm");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Вузол");
            row.Cells[0].AddParagraph();

            row.Cells[1].AddParagraph("DSLAM ");

            row.Cells[2].AddParagraph("Плата/Порт");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Показники якості каналу зв’язку  / Дані вимірювання абонентської лінії");

            row.Cells[1].AddParagraph("Швидкість доступу до мережі Інтернет, Мбіт/сек");

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("5cm");
            table.AddColumn("3cm");
            table.AddColumn("3cm");
            table.AddColumn("3cm");
            table.AddColumn("4cm");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Rшл, Ом");

            row.Cells[1].AddParagraph("Rиз_a-b, МОм");

            row.Cells[2].AddParagraph("Rиз_a-0, МОм");

            row.Cells[3].AddParagraph("Rиз_b-0, МОм");

            row.Cells[4].AddParagraph("Ca-b, мкФ");

            table.AddRow();

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("5cm");
            table.AddColumn("13cm");

            row = table.AddRow();

            row.Cells[0].AddParagraph("Зауваження Абонента щодо виконаних робіт:");

            paragraph = row.Cells[1].AddParagraph();
            paragraph.Format.Font.Size = "12pt";
            for (int i = 0; i < 3; i++)
            {
                row.Cells[1].AddParagraph("…………………………………………………………………………");
            }

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "12pt");
            table.AddColumn("18cm");

            row = table.AddRow();

            row.Cells[0].AddParagraph();
            row.Cells[0].AddParagraph();

            paragraph = row.Cells[0].AddParagraph("Підпис Абонента\t\t\t\t__________________\t\t__________________");
            paragraph.Format.Font.Size = "11pt";

            paragraph = row.Cells[0].AddParagraph("\t\t\t\t\t\t\tпідпис\t\t\t\tП.І.Б.");
            paragraph.Format.Font.Size = "10pt";
            paragraph.Format.Font.Italic = true;

            paragraph = row.Cells[0].AddParagraph();
            paragraph.Format.Font.Size = "10pt";
            paragraph.Format.Font.Italic = true;

            paragraph = row.Cells[0].AddParagraph("Претензій та зауважень щодо прокладання кабелю,");
            paragraph.Format.Font.Size = "9pt";
            paragraph = row.Cells[0].AddParagraph("підключення послуги та демонстрації сервіса  не маю");
            paragraph.Format.Font.Size = "9pt";
            text = paragraph.AddFormattedText("\t__________________\t\t__________________");
            text.Font.Size = "11pt";

            paragraph = row.Cells[0].AddParagraph("(");
            paragraph.Format.Font.Size = "10pt";
            text = paragraph.AddFormattedText("Підпис уповноваженої особи");
            text.Font.Size = "9pt";
            text.Font.Italic = true;
            text = paragraph.AddFormattedText(")\t\t\t\t");
            text.Font.Size = "10pt";
            text = paragraph.AddFormattedText("підпис\t\t\t\tП.І.Б.");
            text.Font.Size = "10pt";
            text.Font.Italic = true;

            for (int i = 0; i < 2; i++)
            {
                paragraph = row.Cells[0].AddParagraph();
                paragraph.Format.Font.Size = "10pt";
                paragraph.Format.Font.Italic = true;
            }

            paragraph = row.Cells[0].AddParagraph("Підпис Виконавця\t\t\t\t__________________\t\t__________________");
            paragraph.Format.Font.Size = "11pt";

            paragraph = row.Cells[0].AddParagraph("\t\t\t\t\t\t\tпідпис\t\t\t\tП.І.Б.");
            paragraph.Format.Font.Size = "10pt";
            paragraph.Format.Font.Italic = true;

            row.Cells[0].AddParagraph();

            return document;
        }
        public Document CreateDutyPDF2(Duty duty, SheetConfiguration sheet)
        {
            var document = new Document();
            var pageSetup = document.DefaultPageSetup.Clone();
            pageSetup.PageFormat = PageFormat.A4;
            pageSetup.Orientation = Orientation.Portrait;
            pageSetup.LeftMargin = "1cm";
            pageSetup.RightMargin = "1cm";
            pageSetup.TopMargin = "0.5cm";
            pageSetup.BottomMargin = "0cm";

            var section = document.AddSection();
            section.PageSetup = pageSetup;

            var headerTable = AddTable(section, "white", "0.5pt", ParagraphAlignment.Right, "Times New Roman", "11pt");
            headerTable.AddColumn("2cm");
            headerTable.AddColumn("9cm");
            headerTable.AddColumn("8.5cm");

            var row = headerTable.AddRow();
            var text = "Наряд на виконання робіт з підключення ADSL";

            row.Cells[0].Format.Font.Italic = true;
            row.Cells[0].Format.Font.Size = "12pt";
            row.Cells[0].AddParagraph(sheet.tableHeader);
            row.Cells[1].Format.Font.Underline = Underline.Single;
            row.Cells[1].Format.Font.Size = "10pt";
            row.Cells[1].AddParagraph($"Наряд на виконання робіт з підключення {duty.Service}");
            row.Cells[2].Format.Font.Underline = Underline.Single;
            row.Cells[2].Format.Font.Size = "10pt";
            row.Cells[2].AddParagraph($"Включити до {duty.ExpirationDate}");

            var table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("19.5cm");
            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Shading.Color = Colors.LightGray;
            row.Cells[0].AddParagraph("Абонент (фізична особа)");

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("6cm");
            table.AddColumn("13.5cm");
            row = table.AddRow();
            row.Cells[0].AddParagraph("Номер телефону ПАТ «Укртелеком»");
            row.Cells[1].AddParagraph(duty.PhoneLine);
            row = table.AddRow();
            row.Cells[0].AddParagraph("Номер особового рахунку Абонента");
            row.Cells[1].AddParagraph(duty.Account);
            row = table.AddRow();
            row.Cells[0].AddParagraph("Прізвище Ім'я По-батькові Абонента");
            row.Cells[1].AddParagraph(duty.CustomerName);

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("2.5cm");
            table.AddColumn("5.5cm");
            table.AddColumn("6cm");
            table.AddColumn("5.5cm");
            row = table.AddRow();
            row.Cells[0].AddParagraph("Паспорт, серія та номер");
            row.Cells[1].Shading.Color = Colors.LightGray;
            row.Cells[2].Format.Font.Size = "8pt";
            row.Cells[2].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].AddParagraph("Реєстраційний номер облікової картки платника податків (Ідентифікаційний номер)");
            row.Cells[3].Shading.Color = Colors.LightGray;

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("6.5cm");
            table.AddColumn("13cm");
            row = table.AddRow();
            row.Cells[0].AddParagraph("Адреса надання Послуги");
            row.Cells[1].AddParagraph(duty.CustomerAddress);

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("2cm");
            table.AddColumn("4.5cm");
            table.AddColumn("13cm");
            row = table.AddRow();
            row.Cells[0].AddParagraph("Контакти");
            row.Cells[1].AddParagraph("Мобільний телефон");
            row.Cells[2].AddParagraph($"{duty.ContactName} | {duty.ContactPhone}");

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Left, "Times New Roman", "10pt");
            table.AddColumn("6.5cm");
            table.AddColumn("13cm");
            row = table.AddRow();
            row.Cells[0].AddParagraph("Підключення  послуги");
            row.Cells[1].AddParagraph(duty.Service);
            row = table.AddRow();
            row.Cells[0].AddParagraph("Лінійні данні ОТА/найменування MSAN");
            row.Cells[1].AddParagraph(duty.LinearDataForOTA);
            row = table.AddRow();
            row.Cells[0].AddParagraph("Зауваження від Укртелекому");
            row.Cells[1].AddParagraph(duty.NotesFromUkt);

            var paragraph = section.AddParagraph();
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);
            paragraph.AddText("Роботи виконані в повному обсязі");
            var str = paragraph.AddFormattedText("\t\t\t\t\tSPEEDTEST: Download______________ Upload______________");
            str.Font.Size = Unit.FromPoint(8);
            section.AddParagraph();
            paragraph = section.AddParagraph("Претензій та зауважень щодо прокладання кабелю,\tАбонент_________________/_______________________/");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);
            paragraph = section.AddParagraph("підключення послуги та демонстрації сервісів не маю");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);

            str = paragraph.AddFormattedText("\t\t\t(Підпис)");
            str.Font.Name = "Times New Roman";
            str.Font.Size = Unit.FromPoint(9);
            str.Font.Italic = true;

            paragraph = section.AddParagraph("Роботи виконанувались з використанням ЗІЗ\t\t«________» _____________________ 20____ року");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);
            section.AddParagraph();
            paragraph = section.AddParagraph("ЗАЯВКА");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(11);
            paragraph.Format.Font.Bold = true;
            paragraph.Format.Font.Underline = Underline.Single;
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            section.AddParagraph();
            paragraph = section.AddParagraph("Бажаю з «________» ___________________ 20____ року");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Center, "Times New Roman", "10pt");
            table.AddColumn("9.75cm");
            table.AddColumn("9.75cm");
            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("1. Отримувати послугу «Інтернет / Інтерактивне ТВ» від Укртелеком»");
            row.Cells[1].AddParagraph("2. Також отримати абонентське обладнання для користування послугою");

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Center, "Times New Roman", "10pt");
            table.AddColumn("4cm");
            table.AddColumn("5.75cm");
            table.AddColumn("4.875cm");
            table.AddColumn("4.875cm");
            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Shading.Color = Colors.LightGray;
            row.Cells[0].AddParagraph("Підпис Абонента");
            row.Cells[1].AddParagraph("Назва послуги");
            paragraph = row.Cells[1].AddParagraph("(тарифний план згідно договору оферти my.ukrtelecom.ua)");
            paragraph.Format.Font.Bold = false;
            row.Cells[2].AddParagraph("Виробник та модель обладнання");
            row.Cells[3].AddParagraph("Умови отримання обладнання"); ;
            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].AddParagraph("□ ADSL");
            row.Cells[1].AddParagraph("□ VDSL ");
            row.Cells[1].AddParagraph("□ Інтерактивне TV");
            row.Cells[1].AddParagraph("□ FTTH (xPON)");
            row.Cells[1].AddParagraph("□ FTTB");
            row.Cells[3].AddParagraph("□ оренда за акційною вартістю");
            row.Cells[3].AddParagraph("□ обладнання видано в ЦОА");

            section.AddParagraph();

            paragraph = section.AddParagraph("Акт приймання-передачі");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(11);
            paragraph.Format.Font.Bold = true;
            paragraph.Format.Font.Underline = Underline.Single;
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph = section.AddParagraph("обладнання у тимчасове платне користування");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(11);
            paragraph.Format.Font.Bold = true;
            paragraph.Format.Font.Underline = Underline.Single;
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            section.AddParagraph();

            paragraph = section.AddParagraph("Цей Акт складено на виконання умов Додаткової угоди укладеної між ПАТ «Укртелеком» та  абонентом");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);

            paragraph = section.AddParagraph("\tМи, що нижче підписалися, представник ");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);
            str = paragraph.AddFormattedText($"\t\t{sheet.fullName}\t\t", TextFormat.Underline);
            str.Font.Size = Unit.FromPoint(12);

            paragraph = section.AddParagraph();
            paragraph.AddFormattedText(
                "що діє на підставі Договору підряду №\t\t\tвід\t\t\tр., з однієї сторони,");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);

            paragraph = section.AddParagraph("та користувач послуг ");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);
            str = paragraph.AddFormattedText($"\t\t{duty.CustomerName}\t\t,", TextFormat.Underline);
            str.Font.Size = Unit.FromPoint(12);

            paragraph = section.AddParagraph("\t\t\t\t\t(П. І. Б. абонента)");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(8);
            paragraph.Format.Font.Italic = true;

            paragraph = section.AddParagraph(
                "з іншої сторони, склали цей Акт про те, що відповідно до Додаткової угоди, ПАТ «Укртелеком» передав, а Абонент прийняв Обладнання у тимчасове платне користування наступне обладнання:");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);

            section.AddParagraph();

            table = AddTable(section, "black", "0.5pt", ParagraphAlignment.Center, "Times New Roman", "10pt");
            table.AddColumn("5cm");
            table.AddColumn("8.5cm");
            table.AddColumn("6cm");
            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Модель (модем / STB приставка)");
            row.Cells[1].AddParagraph("Серійний номер");
            row.Cells[2].AddParagraph("Вартість, грн.");
            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Cells[2].AddParagraph("оренда за акційною вартістю");

            paragraph = section.AddParagraph();
            paragraph.Format.Font.Name = "Times New Roman";
            str = paragraph.AddFormattedText("Комплектація обладнання (I.net):   ", TextFormat.Bold);
            str.Font.Size = Unit.FromPoint(6);
            str = paragraph.AddFormattedText("Модем- 1 шт., Розгалужувач (Splitter)- 1 шт., Блок живлення- 1шт., Телефонний кабель- 2 шт.,  Ethernet кабель-1 шт., Інструкція з використання-1 шт.");
            str.Font.Size = Unit.FromPoint(6);

            paragraph = section.AddParagraph();
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Underline = Underline.Single;
            str = paragraph.AddFormattedText("Комплектація обладнання (Інтерактивне ТВ):   ", TextFormat.Bold);
            str.Font.Size = Unit.FromPoint(6);
            str = paragraph.AddFormattedText("приставка MAG-255 - 1 шт., Пульт дистанційного керування- 1 шт., Блок живлення- 1шт., Елементи живлення (батарейки)- 2 шт.,  Кабель RCA -1 шт.");
            str.Font.Size = Unit.FromPoint(6);

            paragraph = section.AddParagraph("Обладнання є працездатним та перевірено в присутності споживача.");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(9);
            paragraph.Format.Font.Bold = true;
            paragraph.Format.Font.Italic = true;

            section.AddParagraph();

            paragraph = section.AddParagraph("\tВід ПАТ «Укртелеком» (Представник)\t\t\t\t\tАбонент");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);

            section.AddParagraph();
            section.AddParagraph();

            paragraph = section.AddParagraph("________________/___________________________/\t\t________________/___________________________/");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(10);
            paragraph.Format.Font.Italic = true;

            paragraph = section.AddParagraph("\t(Підпис)\t\t(П. І. Б. представника)\t\t\t\t(Підпис)\t\t(П. І. Б. абонента)");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(9);
            paragraph.Format.Font.Italic = true;

            section.AddParagraph();

            paragraph = section.AddParagraph("“_______”______________________20\tр.\t\t\t“_______”____________________20\t\tр.");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(9);

            section.AddParagraph();
            section.AddParagraph();

            paragraph = section.AddParagraph("Шляхом підпису цього документу Абонент:");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(7);

            paragraph = section.AddParagraph(
                "- погоджується з тарифами/цінами та умовами отримання Сервісу, Обладнання, що визначено цім документом (зокрема зі спеціальними умовами отримання послуги «Оренда обладнання») та публічним абонентським договором про надання послуги до договору про надання телекомунікаційних послуг (текст угоди розміщено на сайтах ukrtelecom.ua, my.ukrtelecom.ua);");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(7);

            paragraph = section.AddParagraph("- підтверджує правильність даних, вказаних у цій Заявці;");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(7);

            paragraph = section.AddParagraph(
                "- у випадку замовлення отримання обладнання для користування послугою на умовах оренди за акційної вартістю Абонент погоджується з умовами Додаткової угоди щодо надання обладнання на умовах оренди за акційною вартістю до договору/угоди про надання послуги доступу до мережі Інтернет за технологіями xDSL (текст Додаткової угоди розміщено на сайтах ukrtelecom.ua, my.ukrtelecom.ua).");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = Unit.FromPoint(7);
            paragraph.Format.Font.Bold = true;

            return document;
        }

        public string SaveDutyPDFLocally(Document document, string path, string fileName)
        {
            var filePath = $"{path}{fileName}.pdf";

            var pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();

            using (var pdfDocument = pdfRenderer.PdfDocument)
            {
                using (var fs = new FileStream(filePath, FileMode.Create))
                {
                    pdfDocument.Save(fs);
                }
            }

            return filePath;
        }

        private Table AddTable(Section section, string bordersColor, string bordersWidth, ParagraphAlignment alignment,
            string fontName, string fontSize)
        {
            var table = section.AddTable();
            table.Borders.Color = Color.Parse(bordersColor);
            table.Borders.Width = bordersWidth;
            table.Format.Alignment = alignment;
            table.Format.Font.Name = fontName;
            table.Format.Font.Size = fontSize;
            table.Format.LineSpacingRule = LineSpacingRule.Single;
            return table;
        }
    }
}
