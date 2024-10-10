using System.Diagnostics;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using MigraDoc.DocumentObjectModel.Shapes;
using System.IO;
using System;
using DOCMAIL.Services.Data;
using DOCMAIL.Models;
using System.Collections.Generic;
using System.Data.Common;
using PdfSharp.Pdf;

namespace DOCMAIL.Services.Business
{
    public class InvoiceService
    {
        private InvoiceDAO invoiceDAO = new InvoiceDAO();
        private CabeceraModel cabecera;
        private List<RegistroModel> registros;
        private SubtotalModel subtotal;
        private PieModel pie;

        private string rutaLogo = "C:\\Users\\st2burgoda\\Desktop\\Real Docmail\\docmail\\DOCMAIL\\DOCMAIL\\Content\\Styles\\logo_invoices.png";
        public PdfDocument RetornarInvoice(string nroInvoice = "9730014445")
        {
            Document document = new Document();
            aplicarStyle(document);
            Section section = document.AddSection();

            cabecera = invoiceDAO.ConseguirCabecera(nroInvoice);
            registros = invoiceDAO.ConseguirRegistros(nroInvoice);
            subtotal = invoiceDAO.ConseguirSubtotal(nroInvoice);
            pie = invoiceDAO.ConseguirPie(nroInvoice);

            // Header
            añadirLogo(section);
            añadirCuerpoLogo(section);
            añadirNroInvoice(section);
            añadirCuerpoCabecera(section);

            // Body
            añadirTablaDetalle(section);
            agregarSubtotal(section);
            
            // Footer
            agregarPie(section);

            // Renderizamos el documento
            var pdfRenderer = new PdfDocumentRenderer();
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();

            return pdfRenderer.PdfDocument;

            // Guardardamos
            //string filename = $"C:\\Users\\st2burgoda\\Desktop\\{nroInvoice}.pdf";
            //if (File.Exists(filename))
            //{
                //File.Delete(filename);
            //}
            //pdfRenderer.PdfDocument.Save(filename);

            // Abrir el archivo PDF
            //Process.Start(new ProcessStartInfo(filename) { UseShellExecute = true });
        }
        
        /// <summary>
        /// Predefine los Styles que tendra nuestro documento
        /// </summary>
        /// <param name="document">Documento al cual aplicaremos los Styles definidos</param>
        private void aplicarStyle(Document document)
        {
            // Style el resto del documento
            Style style = document.Styles["Normal"];
            style.Font.Name = "Verdana";
            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("1cm", TabAlignment.Right);
            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

            // Style para la tabla
            style = document.Styles.AddStyle("Table", "Normal");
            style.Font.Name = "Verdana";
            style.Font.Size = 8;
            
            // Style para hacer tabs hasta el final de la hoja
            style = document.Styles.AddStyle("Reference", "Normal");
            style.ParagraphFormat.SpaceBefore = "5mm";
            style.ParagraphFormat.SpaceAfter = "5mm";
            //style.Font.Size = 8;
            style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);
        }

        private void añadirLogo(Section section)
        {
            Image logo = section.Headers.Primary.AddImage(rutaLogo);
            logo.Height = "1cm";
            logo.LockAspectRatio = true;
            logo.RelativeVertical = RelativeVertical.Line;
            logo.RelativeHorizontal = RelativeHorizontal.Margin;
            logo.Top = ShapePosition.Top;
            logo.Left = ShapePosition.Left;
            logo.WrapFormat.Style = WrapStyle.Through;

        }

        private void añadirCuerpoLogo(Section section)
        {
            var cuerpo_logo = agregarParrafo(section, "NEUMATICOS S.A.I.C.", "", 0, "Reference");
            agregarExtralinea(cuerpo_logo, "Date: ", cabecera.CICAFECH, 1, true);
            agregarExtralinea(cuerpo_logo, "Cervantes 1901", "", 0, false);
            agregarExtralinea(cuerpo_logo, "1722  - Merlo - Argentina", "", 0, false);
            agregarExtralinea(cuerpo_logo, "Tel : 4489-6660 ","", 0, false);
        }

        private void añadirNroInvoice(Section section)
        {
            Paragraph nro = section.AddParagraph();

            nro.Format.SpaceBefore = "0.5cm";
            nro.Style = "Reference";
            nro.AddFormattedText("PROFORMA COMMERCIAL INVOICE N°: ", TextFormat.Bold);
            nro.AddText(cabecera.CICANCIN);
            nro.Format.Alignment = ParagraphAlignment.Center;
        }

        private void añadirCuerpoCabecera(Section section)
        {
            Paragraph destinatarios = section.AddParagraph();

            destinatarios.Format.SpaceBefore = "0.5cm";
            destinatarios.Style = "Reference";

            destinatarios.AddFormattedText("TO", TextFormat.Bold);
            destinatarios.AddTab();
            destinatarios.AddFormattedText("CONSIGNEE", TextFormat.Bold);
            destinatarios.AddLineBreak();
            destinatarios.AddText($"{cabecera.CICANOMD}");
            destinatarios.AddTab();
            destinatarios.AddText($"{cabecera.CICANOMC}");

            destinatarios = section.AddParagraph();
            destinatarios.Style = "Reference";
            destinatarios.Format.SpaceBefore = "0.2cm";

            destinatarios.AddText($"{cabecera.CICADIRD}");
            destinatarios.AddTab();
            destinatarios.AddText($"{cabecera.CICADIRC}");
            destinatarios.AddLineBreak();
            destinatarios.AddText($"{cabecera.CICALOCD}");
            destinatarios.AddTab();
            destinatarios.AddText($"{cabecera.CICALOCC}");

            destinatarios = section.AddParagraph();
            destinatarios.Style = "Reference";

            destinatarios.AddFormattedText($"{cabecera.CICACOPD} · {cabecera.CICAPAID}", TextFormat.Bold);
            destinatarios.AddTab();
            destinatarios.AddFormattedText($"{cabecera.CICACOPC} · {cabecera.CICAPAIC}", TextFormat.Bold);

        }

        private void añadirTablaDetalle(Section section)
        {
            Table table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = Color.FromRgb(128, 128, 128);
            table.Borders.Width = 0.75;
            table.Rows.LeftIndent = 0;
            table.Rows.Alignment = RowAlignment.Right;

            Column column = table.AddColumn("1.7cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn("1.5cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("4.8cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("0.2cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn("2.6cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("1.9cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("2.3cm");
            column.Format.Alignment = ParagraphAlignment.Right;
            column.Shading.Color = Color.FromRgb(240, 240, 240);

            agregarEncabezados(table);
            foreach (RegistroModel registro in registros)
            {
                agregarRegistro(table, registro);
            }
        }
        private void agregarEncabezados(Table table)
        {
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Font.Bold = true;
            row.Height = Unit.FromCentimeter(0.2);
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Shading.Color = Color.FromRgb(207, 207, 207);
            string[] items = { "Quantity", "Code", "Description", "", "Un Weight KG", "Tot Weight KG", "Unit Price", "Amount U$S" };
            int i = 0;
            foreach(string item in items)
            {
                row.Cells[i].AddParagraph(item);
                row.Cells[i].Borders.Color = Colors.Black;
                row.Cells[i].Format.Alignment = ParagraphAlignment.Left;
                row.Cells[i].VerticalAlignment = VerticalAlignment.Bottom;
                i++;
            }
        }
        private void agregarRegistro(Table table, RegistroModel registro)
        {
            Row row = table.AddRow();
            string[] items = { registro.CIDECANT, registro.CIDEMATE, registro.CIDEDESC, registro.CIDEPEUN, registro.CIDEPETO, registro.CIDEPRUN, registro.CIDEPRTO };
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
                if (item == registro.CIDEPRTO)
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
        private void agregarSubtotal(Section section)
        {
            section.AddParagraph();

            Table table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = Color.FromRgb(128, 128, 128);
            table.Borders.Width = 0.75;
            table.Rows.LeftIndent = 0;
            table.Rows.Alignment = RowAlignment.Right;

            Column column1 = table.AddColumn(Unit.FromCentimeter(12)); 
            Column column2 = table.AddColumn(Unit.FromCentimeter(2.3));
            column2.Shading.Color = Color.FromRgb(240, 240, 240);

            column1.Format.Alignment = ParagraphAlignment.Left;
            column2.Format.Alignment = ParagraphAlignment.Right; 

            
            var propertyNames = new[] { "FCA", "FLETE", "SEGURO", "TOTAL" };
            foreach (string propertyName in propertyNames)
            {
                // Usar reflexión para obtener el valor de la propiedad
                var propertyValue = subtotal.GetType().GetProperty(propertyName)?.GetValue(subtotal, null)?.ToString();

                if (propertyValue != null)
                {
                    Row row = table.AddRow();
                    row.Height = Unit.FromCentimeter(0.2);
                    row.HeadingFormat = true;

                    row.Cells[0].AddParagraph(propertyName); 
                    row.Cells[0].Borders.Width = 0.5; 
                    row.Cells[0].Borders.Color = Colors.Black; 

                    row.Cells[1].AddParagraph(propertyValue);
                    row.Cells[1].Borders.Width = 0.5;
                    row.Cells[1].Borders.Color = Colors.Black;
                }
            }
            table.Format.Alignment = ParagraphAlignment.Right;
        }


        private void agregarPie(Section section)
        {
            Paragraph paragraph = section.AddParagraph();

            paragraph.Format.SpaceBefore = "0.25cm";
            paragraph.AddFormattedText("PACKING :", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddFormattedText("Net Weight (Kg): ", TextFormat.Bold);
            paragraph.AddText(pie.CIPINETW);
            paragraph.AddFormattedText("  ·  ", TextFormat.Bold);
            paragraph.AddFormattedText("Gross Weight (Kg): ", TextFormat.Bold);
            paragraph.AddText(pie.CIPIGROW);
            paragraph.AddLineBreak();
            paragraph.AddTab();
            paragraph.AddTab();
            paragraph.AddFormattedText("Measurement (M3): ", TextFormat.Bold);
            paragraph.AddText(pie.CIPIMEAS);
            paragraph.AddFormattedText("  ·  ", TextFormat.Bold);
            paragraph.AddFormattedText("Container: ", TextFormat.Bold);
            paragraph.AddText(pie.CIPICONT);

            var parrafoAnt = agregarParrafo(section, "VESSEL :", pie.CIPIVESS,2);
            agregarExtralinea(parrafoAnt, "MARITIME :", pie.CIPIMARI, 2,false);

            agregarParrafo(section, "OBSERVATION :", pie.CIPIOBS1, 1);
            agregarParrafo(section, "MARITIME :", pie.CIPIMARI, 2);

            parrafoAnt = agregarParrafo(section, "Payment Terms :",pie.CIPIPAYM, 1);
            agregarExtralinea(parrafoAnt, "Conditions :", pie.CIPICOND, 2, false);

            agregarParrafo(section, "NOTIFY (1) :", pie.CIPINO11, 1);

            parrafoAnt = agregarParrafo(section, "NOTIFY (2) :", pie.CIPINO12, 1);
            agregarExtralinea(parrafoAnt, "DOCUMENT :", pie.CIPIDOC1, 3, true);

            Paragraph footer = section.Footers.Primary.AddParagraph();
            footer.AddText("PIRELLI NEUMATICOS S.A.I.C.· Cervantes 1901 ·1722 - Merlo · Argentina\n");
            footer.AddText($"Invoice Nro: {pie.CIPINCIN}");
            footer.Format.Font.Size = 7;
            footer.Format.Alignment = ParagraphAlignment.Center;
        }

        private Paragraph agregarParrafo(Section section,string tipo, string contenido,int tabs,string style = "")
        {
            Paragraph paragraph = section.AddParagraph();
            if (style != "")
            {
                paragraph.Style = style;
            }
            paragraph.Format.SpaceBefore = "0.25cm";
            paragraph.AddFormattedText(tipo, TextFormat.Bold);
            agregarTabs(paragraph, tabs);
            paragraph.AddText(contenido);
            return paragraph;        
        }

        private void agregarExtralinea(Paragraph paragraph,string tipo,string contenido,int tabs,bool mismaLinea)
        {
            if (!mismaLinea)
            {
                paragraph.AddLineBreak();
            }
            else
            {
                agregarTabs(paragraph, tabs);
                tabs--;
            }
            paragraph.AddFormattedText(tipo, TextFormat.Bold);
            agregarTabs(paragraph, tabs);
            paragraph.AddText(contenido);
        }

        /// <summary>
        /// Agrega tabulaciones en el parrafo indicado
        /// </summary>
        /// <param name="paragraph"> Parrafo en el que queremos tabular </param>
        /// <param name="tabs"> Cantidad tabulaciones </param>
        private void agregarTabs(Paragraph paragraph,int tabs)
        {
            for (int i = 0; i < tabs; i++)
            {
                paragraph.AddTab();
            }
        }
    }
}
