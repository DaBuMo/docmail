using System.Diagnostics;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using MigraDoc.DocumentObjectModel.Shapes;
using System.IO;
using System;

namespace DOCMAIL.Services.Business
{
    public class InvoiceService
    {
        private string rutaLogo = "C:\\Users\\st2burgoda\\Desktop\\Real Docmail\\docmail\\DOCMAIL\\DOCMAIL\\Content\\Styles\\logo_invoices.png";

        public void RetornarInvoice(string CICAFECH = "dd.MM.yyyy", string CICANCIN = "CICANCIN")
        {
            Document document = new Document();
            Style style = document.Styles["Normal"];

            style.Font.Name = "Verdana";
            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("1cm", TabAlignment.Right);
            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);
            style = document.Styles.AddStyle("Table", "Normal");
            style.Font.Name = "Verdana";
            style.Font.Size = 9;
            style = document.Styles.AddStyle("Reference", "Normal");
            style.ParagraphFormat.SpaceBefore = "5mm";
            style.ParagraphFormat.SpaceAfter = "5mm";
            style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);

            Section section = document.AddSection();

            // Añadir la imagen del logo a la izquierda image.Left = "-1.5cm"; // Ajustar a un valor negativo para moverlo más a la izquierda
            // Logo
            Image logo = section.Headers.Primary.AddImage(rutaLogo);
            logo.Height = "1cm";
            logo.LockAspectRatio = true;
            logo.RelativeVertical = RelativeVertical.Line;
            logo.RelativeHorizontal = RelativeHorizontal.Margin;
            logo.Top = ShapePosition.Top;
            logo.Left = ShapePosition.Left;
            logo.WrapFormat.Style = WrapStyle.Through;

            // Encabezado
            Paragraph cuerpo_logo = section.AddParagraph();
            cuerpo_logo.Style = "Reference";
            cuerpo_logo.AddFormattedText("NEUMATICOS S.A.I.C.", TextFormat.Bold);
            cuerpo_logo.AddTab();
            cuerpo_logo.AddFormattedText("Date: ", TextFormat.Bold);
            cuerpo_logo.AddDateField(CICAFECH);
            cuerpo_logo.AddLineBreak();
            cuerpo_logo.AddFormattedText("Cervantes 1901", TextFormat.Bold);
            cuerpo_logo.AddLineBreak();
            cuerpo_logo.AddFormattedText("1722  - Merlo - Argentina", TextFormat.Bold);
            cuerpo_logo.AddLineBreak();
            cuerpo_logo.AddFormattedText("Tel : 4489-6660", TextFormat.Bold);

            // Nro invoice
            Paragraph nro = section.AddParagraph();
            nro.Format.SpaceBefore = "1cm";
            nro.Style = "Reference";
            nro.AddFormattedText("PROFORMA COMMERCIAL INVOICE N°: ", TextFormat.Bold);
            nro.AddText(CICANCIN);
            nro.Format.Alignment = ParagraphAlignment.Center;
            
            // 1/3 seccion
            Paragraph destinatarios = section.AddParagraph();
            destinatarios.Format.SpaceBefore = "1cm";
            destinatarios.Style = "Reference";
            destinatarios.AddFormattedText("TO", TextFormat.Bold);
            destinatarios.AddTab();
            destinatarios.AddFormattedText("CONSIGNEE", TextFormat.Bold);
            destinatarios.AddLineBreak();
            destinatarios.AddText("CICANOMD");
            destinatarios.AddTab();
            destinatarios.AddText("CICANOMC");
            // 2/3 seccion
            destinatarios = section.AddParagraph();
            destinatarios.Format.SpaceBefore = "1cm";
            destinatarios.Style = "Reference";
            destinatarios.AddText("CICADIRD");
            destinatarios.AddTab();
            destinatarios.AddText("CICADIRC");
            destinatarios.AddLineBreak();
            destinatarios.AddText("CICALOCD");
            destinatarios.AddTab();
            destinatarios.AddText("CICALOCC");
            // 3/3 seccion
            destinatarios = section.AddParagraph();
            destinatarios.Style = "Reference";
            destinatarios.AddFormattedText("CICACOPD · CICAPAID", TextFormat.Bold);
            destinatarios.AddTab();
            destinatarios.AddFormattedText("CICACOPC · CICAPAIC", TextFormat.Bold);

            // Detalle 
            Table table = section.AddTable();
            table.Format.SpaceBefore = "1cm";
            table.Style = "Table";
            table.Borders.Color = Color.FromRgb(128, 128, 128);
            table.Borders.Width = 0.1;
            table.Borders.Left.Width = 0.1;
            table.Borders.Right.Width = 0.1;
            table.Rows.LeftIndent = 0;
            table.SetEdge(0, 0, 6, 2, Edge.Box, BorderStyle.Single, 0.75, Color.Empty);

            Column column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("4cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("1cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = Color.FromRgb(0, 0, 255);
            row.Cells[0].AddParagraph("Quantity");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[1].AddParagraph("Code");
            row.Cells[1].Format.Font.Bold = false;
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[2].AddParagraph("Description");
            row.Cells[2].Format.Font.Bold = false;
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[3].AddParagraph("");
            row.Cells[3].Format.Font.Bold = false;
            row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[4].AddParagraph("Un Weight KG");
            row.Cells[4].Format.Font.Bold = false;
            row.Cells[4].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[4].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[5].AddParagraph("Tot Weight KG");
            row.Cells[5].Format.Font.Bold = false;
            row.Cells[5].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[5].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[6].AddParagraph("Unit Price");
            row.Cells[6].Format.Font.Bold = false;
            row.Cells[6].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[6].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[7].AddParagraph("Amount U$S");
            row.Cells[7].Format.Font.Bold = false;
            row.Cells[7].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[7].VerticalAlignment = VerticalAlignment.Bottom;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = Color.FromRgb(0, 0, 155);
            row.Cells[1].AddParagraph("Quantity");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].AddParagraph("Unit Price");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].AddParagraph("Discount (%)");
            row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[4].AddParagraph("Taxable");
            row.Cells[4].Format.Alignment = ParagraphAlignment.Left;
            // Footer
            Paragraph footer = section.Footers.Primary.AddParagraph();
            footer.AddText("PIRELLI NEUMATICOS S.A.I.C.· Cervantes 1901 ·1722 - Merlo · Argentina\n");
            footer.AddText("Invoice Nro: Test");
            footer.Format.Font.Size = 9;
            footer.Format.Alignment = ParagraphAlignment.Center;

            // Renderizamos el documento
            var pdfRenderer = new PdfDocumentRenderer();
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();

            // Guardardamos
            string filename = "C:\\Users\\st2burgoda\\Desktop\\TablaEjemploConLogo.pdf";
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }
            pdfRenderer.PdfDocument.Save(filename);

            // Abrir el archivo PDF
            Process.Start(new ProcessStartInfo(filename) { UseShellExecute = true });
        }

    }
}
