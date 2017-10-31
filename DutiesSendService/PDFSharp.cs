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

            paragraph = row.Cells[0].AddParagraph("Підпис Абонента\t\t__________________\t\t");
            paragraph.Format.Font.Size = "11pt";
            text = paragraph.AddFormattedText($"{duty.CustomerName}");
            text.Font.Underline = Underline.Single;

            paragraph = row.Cells[0].AddParagraph("\t\t\t\t\tпідпис\t\t\t\t\tП.І.Б.");
            paragraph.Format.Font.Size = "10pt";
            paragraph.Format.Font.Italic = true;

            paragraph = row.Cells[0].AddParagraph();
            paragraph.Format.Font.Size = "10pt";
            paragraph.Format.Font.Italic = true;

            paragraph = row.Cells[0].AddParagraph("Претензій та зауважень не маю");
            paragraph.Format.Font.Size = "10pt";
            text = paragraph.AddFormattedText("\t__________________\t\t");
            text.Font.Size = "11pt";
            text = paragraph.AddFormattedText($"{duty.CustomerName}");
            text.Font.Size = "11pt";
            text.Font.Underline = Underline.Single;

            paragraph = row.Cells[0].AddParagraph("(");
            paragraph.Format.Font.Size = "10pt";
            text = paragraph.AddFormattedText("Підпис уповноваженої особи");
            text.Font.Size = "9pt";
            text.Font.Italic = true;
            text = paragraph.AddFormattedText(")\t\t");
            text.Font.Size = "10pt";
            text = paragraph.AddFormattedText("підпис\t\t\t\t\tП.І.Б.");
            text.Font.Size = "10pt";
            text.Font.Italic = true;

            for (int i = 0; i < 2; i++)
            {
                paragraph = row.Cells[0].AddParagraph();
                paragraph.Format.Font.Size = "10pt";
                paragraph.Format.Font.Italic = true;
            }

            paragraph = row.Cells[0].AddParagraph("Підпис Виконавця\t\t__________________\t\t(__________________)");
            paragraph.Format.Font.Size = "11pt";

            paragraph = row.Cells[0].AddParagraph("\t\t\t\t\tпідпис\t\t\t\t\tП.І.Б.");
            paragraph.Format.Font.Size = "10pt";
            paragraph.Format.Font.Italic = true;

            row.Cells[0].AddParagraph();

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
