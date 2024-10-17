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
using System.Web;
using PdfSharp.Charting;
using PdfSharp.Diagnostics;

namespace DOCMAIL.Services.Business
{
    public class InvoiceService
    {
        private InvoiceDAO invoiceDAO = new InvoiceDAO();
        private CabeceraModel cabecera;
        private List<RegistroModel> registros;
        private SubtotalModel subtotal;
        private PieModel pie;


        string rutaLogo = HttpContext.Current.Server.MapPath("~/Content/Styles/logo_invoices.png");
        /// <summary>
        /// Se encarga de construir y modelar un PDF de la Invoice en base a los datos conseguidos en la BD 
        /// </summary>
        /// <param name="nroInvoice">Nro de invoce del cual buscaremos los datos y formaremos el diseño del PDF</param>
        /// <returns>PdfDocument</returns>
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

            var corte = section.AddParagraph();
            corte.Format.SpaceBefore = "1mm";

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

        /// <summary>
        /// Añade un logo en la parte superior de la pantalla
        /// </summary>
        /// <param name="section">Seccion en la cual aplicaremos nuestro logo</param>
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
        /// <summary>
        /// El cuerpo del logo en la parte superior de la pantalla
        /// </summary>
        /// <param name="section">Seccion en la cual aplicaremos el cuerpo de nuestro logo</param>
        private void añadirCuerpoLogo(Section section)
        {
            var cuerpo_logo = agregarParrafo(section, "NEUMATICOS S.A.I.C.", "", 0, "Reference");
            agregarExtralinea(cuerpo_logo, "Date: ", cabecera.CICAFECH, 1, true);
            agregarExtralinea(cuerpo_logo, "Cervantes 1901", "", 0, false);
            agregarExtralinea(cuerpo_logo, "1722  - Merlo - Argentina", "", 0, false);
            agregarExtralinea(cuerpo_logo, "Tel : 4489-6660 ","", 0, false);
        }

        /// <summary>
        /// Añade el Nro del Invoice en la parte superior de la pantalla
        /// </summary>
        /// <param name="section">Seccion en la cual aplicaremos nuestro Nro de Invoice</param>
        private void añadirNroInvoice(Section section)
        {
            FormattedText linea;
            Paragraph nro = section.AddParagraph();


            nro.Format.SpaceBefore = "1mm";
            nro.Style = "Reference";
            linea = nro.AddFormattedText("PROFORMA COMMERCIAL INVOICE N°: ", TextFormat.Bold);
            linea.Font.Size = 9;
            linea = nro.AddFormattedText(cabecera.CICANCIN);
            linea.Font.Size = 10;
            nro.Format.Alignment = ParagraphAlignment.Center;
        }

        /// <summary>
        /// Añade los detalles de cabecera
        /// </summary>
        /// <param name="section">Seccion en la cual aplicaremos nuestra cabecera</param>
        private void añadirCuerpoCabecera(Section section)
        {

            Table table = section.AddTable();
            table.Borders.Width = 0.75; // Grosor del borde
            table.Borders.Color = Colors.Black; // Color del borde

            Column column = table.AddColumn();
            column.Width = Unit.FromCentimeter(16.3);

            table.Format.Alignment = ParagraphAlignment.Center;

            Row row = table.AddRow();

            Paragraph destinatarios = row.Cells[0].AddParagraph();
            destinatarios.Format.SpaceBefore = "1mm";
            destinatarios.Style = "Reference";

            FormattedText linea = destinatarios.AddFormattedText("TO", TextFormat.Bold);
            linea.Font.Size = 9;
            destinatarios.AddTab();
            linea = destinatarios.AddFormattedText("CONSIGNEE", TextFormat.Bold);
            linea.Font.Size = 9;
            destinatarios.AddLineBreak();

            linea = destinatarios.AddFormattedText($"{cabecera.CICANOMD}");
            linea.Font.Size = 10;
            destinatarios.AddTab();
            linea = destinatarios.AddFormattedText($"{cabecera.CICANOMC}");
            linea.Font.Size = 10;

            destinatarios.AddLineBreak();

            linea = destinatarios.AddFormattedText($"{cabecera.CICADIRD}");
            linea.Font.Size = 10;
            destinatarios.AddTab();
            linea = destinatarios.AddFormattedText($"{cabecera.CICADIRC}");
            linea.Font.Size = 10;

            destinatarios.AddLineBreak();
            linea = destinatarios.AddFormattedText($"{cabecera.CICALOCD}");
            linea.Font.Size = 10;
            destinatarios.AddTab();
            linea = destinatarios.AddFormattedText($"{cabecera.CICALOCC}");
            linea.Font.Size = 10;

            destinatarios.AddLineBreak();
            linea = destinatarios.AddFormattedText($"{cabecera.CICACOPD} · {cabecera.CICAPAID}");
            linea.Font.Size = 10;
            destinatarios.AddTab();
            linea = destinatarios.AddFormattedText($"{cabecera.CICACOPC} · {cabecera.CICAPAIC}");
            linea.Font.Size = 10;


            destinatarios.Format.SpaceAfter = "0.3cm"; // Margen inferior

            // Finalizar la tabla
            row.Cells[0].Borders.Width = 0.75; // Grosor del borde de la celda
            row.Cells[0].Borders.Color = Colors.Black; // Color del borde de la celda
        }

        /// <summary>
        /// Crea la tabla en la cual ingresaremos los registros
        /// </summary>
        /// <param name="section">Seccion en la cual aplicaremos nuestra tabla</param>
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

            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("4.8cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("0.1cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("1.4cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            column = table.AddColumn("2.1cm");
            column.Format.Alignment = ParagraphAlignment.Right;
            column.Shading.Color = Color.FromRgb(240, 240, 240);

            agregarEncabezados(table);
            foreach (RegistroModel registro in registros)
            {
                agregarRegistro(table, registro);
            }
        }

        /// <summary>
        /// Agrega los encabezados de los registros
        /// </summary>
        /// <param name="table">Tabla en la cual aplicaremos los encabezados</param>
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
                row.Cells[i].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[i].VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                i++;
            }
        }

        /// <summary>
        /// Agrega un registro a la tabla respetando el formato dado a esta
        /// </summary>
        /// <param name="table">Tabla en la cual aplicaremos los registros</param>
        /// <param name="registro">Registro modelado a partir de RegistroModel</param>
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

        /// <summary>
        /// Agrega la seccion de subtotal
        /// </summary>
        /// <param name="section">Seccion en la cual aplicaremos la seccion de subtotal</param>
        private void agregarSubtotal(Section section)
        {
            section.AddParagraph();

            Table table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = Color.FromRgb(128, 128, 128);
            table.Borders.Width = 0.75;
            table.Rows.LeftIndent = 0;
            table.Rows.Alignment = RowAlignment.Right;

            Column column1 = table.AddColumn(Unit.FromCentimeter(10.3)); 
            Column column2 = table.AddColumn(Unit.FromCentimeter(2.1));
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

        /// <summary>
        ///  Agrega la seccion de pie
        /// </summary>
        /// <param name="section">Seccion en la cual aplicaremos nuestro footer</param>
        private void agregarPie(Section section)
        {
            FormattedText linea;
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "0.25cm";

            linea = paragraph.AddFormattedText("PACKING :", TextFormat.Bold);
            linea.Font.Size = 8;
            paragraph.AddTab();
            linea = paragraph.AddFormattedText("Net Weight (Kg): ", TextFormat.Bold);
            linea.Font.Size = 8;
            linea = paragraph.AddFormattedText(pie.CIPINETW);
            linea.Font.Size = 9;
            paragraph.AddFormattedText("  ·  ", TextFormat.Bold);
            linea = paragraph.AddFormattedText("Gross Weight (Kg): ", TextFormat.Bold);
            linea.Font.Size = 8;
            linea = paragraph.AddFormattedText(pie.CIPIGROW);
            linea.Font.Size = 9;
            paragraph.AddLineBreak();
            paragraph.AddTab();
            paragraph.AddTab();
            linea = paragraph.AddFormattedText("Measurement (M3): ", TextFormat.Bold);
            linea.Font.Size = 8;
            linea = paragraph.AddFormattedText(pie.CIPIMEAS);
            linea.Font.Size = 9;
            paragraph.AddFormattedText("  ·  ", TextFormat.Bold);
            linea = paragraph.AddFormattedText("Container: ", TextFormat.Bold);
            linea.Font.Size = 8;
            linea = paragraph.AddFormattedText(pie.CIPICONT);
            linea.Font.Size = 9;

            var parrafoAnt = agregarParrafo(section, "VESSEL :", pie.CIPIVESS,2,tamaño:8);
            agregarExtralinea(parrafoAnt, "MARITIME :", pie.CIPIMARI, 2,false, tamaño: 8);

            agregarParrafo(section, "OBSERVATION :", pie.CIPIOBS1, 1, tamaño: 8);
            agregarParrafo(section, "MARITIME :", pie.CIPIMARI, 2, tamaño: 8);

            parrafoAnt = agregarParrafo(section, "Payment Terms :",pie.CIPIPAYM, 1, tamaño: 8);
            agregarExtralinea(parrafoAnt, "Conditions :", pie.CIPICOND, 2, false, tamaño: 8);

            agregarParrafo(section, "NOTIFY (1) :", pie.CIPINO11, 1, tamaño: 8);

            parrafoAnt = agregarParrafo(section, "NOTIFY (2) :", pie.CIPINO12, 1, tamaño: 8);
            agregarExtralinea(parrafoAnt, "DOCUMENT :", pie.CIPIDOC1, 3, true, tamaño: 8);

            Paragraph footer = section.Footers.Primary.AddParagraph();
            footer.AddText("PIRELLI NEUMATICOS S.A.I.C.· Cervantes 1901 ·1722 - Merlo · Argentina\n");
            footer.AddText($"Invoice Nro: {pie.CIPINCIN}");
            footer.Format.Font.Size = 8;
            footer.Format.Alignment = ParagraphAlignment.Center;
        }

        /// <summary>
        /// Agrega un parrafo en la seccion indicada
        /// </summary>
        /// <param name="section"> Seccion donde se agregara los parrafos </param>
        /// <param name="tipo"> Definicion del contenido </param>
        /// <param name="contenido"> </param>
        /// <param name="tabs"> Cantidad tabulaciones </param>
        /// <param name="style"> Style que se le aplicara al parrafo </param>
        private Paragraph agregarParrafo(Section section,string tipo, string contenido,int tabs,string style = "", int tamaño = 9,bool negrilla = true)
        {
            FormattedText linea;
            Paragraph paragraph = section.AddParagraph();
            if (style != "")
            {
                paragraph.Style = style;
            }
            paragraph.Format.SpaceBefore = "0.25cm";
            if (negrilla)
            {
                linea = paragraph.AddFormattedText(tipo, TextFormat.Bold);
            }
            else
            {
                linea = paragraph.AddFormattedText(tipo);
            }
            linea.Font.Size = tamaño;
            agregarTabs(paragraph, tabs);
            paragraph.AddText(contenido);
            return paragraph;        
        }

        /// <summary>
        /// Agrega una linea extra en un parrafo dado
        /// </summary>
        /// <param name="paragraph"> Parrafo donde se agregara la linea extra </param>
        /// <param name="tipo"> Definicion del contenido </param>
        /// <param name="contenido"> </param>
        /// <param name="tabs"> Cantidad tabulaciones </param>
        /// <param name="mismaLinea"> Indica si se agrega o no un salto de linea antes de agregar la nueva linea </param>
        private void agregarExtralinea(Paragraph paragraph,string tipo,string contenido,int tabs,bool mismaLinea,int tamaño = 9,bool negrilla = true)
        {
            FormattedText linea;
            if (!mismaLinea)
            {
                paragraph.AddLineBreak();
            }
            else
            {
                agregarTabs(paragraph, tabs);
                tabs--;
            }

            if (negrilla)
            {
                linea = paragraph.AddFormattedText(tipo, TextFormat.Bold);
            }
            else
            {
                linea = paragraph.AddFormattedText(tipo);
            }

            linea.Font.Size = tamaño;
            agregarTabs(paragraph, tabs);
            linea = paragraph.AddFormattedText(contenido);
            linea.Font.Size = tamaño+1;

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
