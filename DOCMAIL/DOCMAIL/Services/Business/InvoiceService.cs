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
            style.Font.Size = 7;
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
            destinatarios.Format.SpaceBefore = "0.5cm";
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
            table.Style = "Table";
            table.Borders.Color = Color.FromRgb(128, 128, 128);
            table.Borders.Width = 0.75;
            table.Rows.LeftIndent = 0;
            table.Rows.Alignment = RowAlignment.Right;

            Column column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("3.5cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("0.2cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn("2.8cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Right;


            agregarEncabezados(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);
            agregarRegistro(table);


            table.SetEdge(0, 0, 6, 2, Edge.Box, BorderStyle.Single, 0.75, Color.Empty);

            section.AddParagraph();

            table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = Color.FromRgb(128, 128, 128);
            table.Borders.Width = 0.75;
            table.Rows.LeftIndent = 0;
            table.Rows.Alignment = RowAlignment.Right;

            // Definir columnas
            Column column1 = table.AddColumn(Unit.FromCentimeter(11)); // Primera columna
            Column column2 = table.AddColumn(Unit.FromCentimeter(2.5));  // Segunda columna (CIDEPRTO)
            column1.Format.Alignment = ParagraphAlignment.Left;
            column2.Format.Alignment = ParagraphAlignment.Right; // Alinear contenido a la izquierda

            // Agregar filas con contenido
            string[] items = { "Subtotal", "Flete Internacional", "Seguro", "Importe Final" };
            foreach (string item in items)
            {
                Row row = table.AddRow();
                row.Height = Unit.FromCentimeter(0.2);
                row.HeadingFormat = true;

                // Agregar contenido a la celda
                row.Cells[0].AddParagraph(item);
                row.Cells[0].Borders.Width = 0.5; // Ancho del borde de la celda
                row.Cells[0].Borders.Color = Colors.Black; // Color del borde

                // Agregar "CIDEPRTO" a la segunda columna
                row.Cells[1].AddParagraph("CIDEPRTO");
                row.Cells[1].Borders.Width = 0.5;
                row.Cells[1].Borders.Color = Colors.Black;
            }
            // Alinear la tabla a la derecha
            table.Format.Alignment = ParagraphAlignment.Right;

            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "0.5cm";
            paragraph.AddFormattedText("PACKING :", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddFormattedText("Net Weight (Kg): ", TextFormat.Bold);
            paragraph.AddText("CIPINETW");
            paragraph.AddFormattedText(" · ", TextFormat.Bold);
            paragraph.AddFormattedText("Gross Weight (Kg): ", TextFormat.Bold);
            paragraph.AddText("CIPIGROW");
            paragraph.AddLineBreak();
            paragraph.AddTab();
            paragraph.AddTab();
            paragraph.AddFormattedText("Measurement (M3): ", TextFormat.Bold);
            paragraph.AddText("CIPIMEAS");
            paragraph.AddFormattedText(" · ", TextFormat.Bold);
            paragraph.AddFormattedText("Container: ", TextFormat.Bold);
            paragraph.AddText("CIPICONT");

            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "0.5cm";
            paragraph.AddFormattedText("VESSEL :", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddTab();

            paragraph.AddText("CIPIVESS");
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("MARITIME :", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddTab();

            paragraph.AddText("CIPIMARI");

            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "0.5cm";
            paragraph.AddFormattedText("OBSERVATION :", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddText("CIPIOBS1");

            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "0.5cm";
            paragraph.AddFormattedText("Payment Terms :", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddText("CIPIPAYM");
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("Conditions :", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddTab();

            paragraph.AddText("CIPICOND");

            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "0.5cm";
            paragraph.AddFormattedText("NOTIFY (1) :", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddText("CIPINO11");

            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "0.5cm";
            paragraph.AddFormattedText("NOTIFY (2) :", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddText("CIPINO12");

            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "0.5cm";
            paragraph.AddFormattedText("DOCUMENT :", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddText("CIPIDOC1");



            // Footer
            Paragraph footer = section.Footers.Primary.AddParagraph();
            footer.AddText("PIRELLI NEUMATICOS S.A.I.C.· Cervantes 1901 ·1722 - Merlo · Argentina\n");
            footer.AddText("Invoice Nro: CICANCIN");
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
        private void agregarEncabezados(Table table)
        {
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Font.Bold = true;
            row.Height = Unit.FromCentimeter(0.2);
            row.Format.Alignment = ParagraphAlignment.Center;
            //row.Shading.Color = Color.FromRgb(0, 0, 255);
            string[] items = { "Quantity", "Code", "Description", "", "Un Weight KG", "Tot Weight KG", "Unit Price", "Amount U$S" };
            int i = 0;
            foreach(string item in items)
            {
                row.Cells[i].AddParagraph(item);
                row.Cells[i].Format.Alignment = ParagraphAlignment.Left;
                row.Cells[i].VerticalAlignment = VerticalAlignment.Bottom;
                i++;
            }
        }
        private void agregarRegistro(Table table)
        {
            
            Row row = table.AddRow();
            string[] items = { "CIDECANT", "CIDEMATE", "CIDEDESC", "CIDEPEUN", "CIDEPETO", "CIDEPRUN", "CIDEPRTO" };
            row.Height = Unit.FromCentimeter(0.2);
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            int i = 0;
            foreach (string item in items)
            {
                if (i == 3)
                {
                    i++;
                }
                row.Cells[i].AddParagraph(item);
                if (item == "CIDEPRTO")
                {
                    row.Cells[i].Format.Alignment = ParagraphAlignment.Right;
                }
                else
                {
                    row.Cells[i].Format.Alignment = ParagraphAlignment.Left;

                }
                i++;

            }
        }
    }
}
