using DOCMAIL.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Web;

namespace DOCMAIL.Services.Data
{
    public class InvoiceDAO
    {

        /// <summary>
        /// Se encarga de conseguir los datos relacionados a la cabecera del Invoice y devolvera un CabeceraModel con dichos datos.
        /// </summary>
        /// <param name="nroInvoice">Nro del Commercial Invoice del cual se daran formato los datos</param>
        /// <returns>CabeceraModel</returns>
        public CabeceraModel ConseguirCabecera(string nroInvoice)
        {
            DataTable result;
            CabeceraModel cabecera = new CabeceraModel();

            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro("@NCIN", nroInvoice);
                result = conn.Executar_proc("DOCMAIL_REPORTE_CI_CABECERA", "COMEX");
            }

            if (result.Rows.Count == 0)
            {
                cabecera.CICAFECH = "";
                cabecera.CICANCIN = "";
                cabecera.CICANOMD = "";
                cabecera.CICADIRD = "";
                cabecera.CICALOCD = "";
                cabecera.CICACOPD = "";
                cabecera.CICAPAID = "";
                cabecera.CICANOMC = "";
                cabecera.CICADIRC = "";
                cabecera.CICALOCC = "";
                cabecera.CICACOPC = "";
                cabecera.CICAPAIC = "";
            }
            else
            {
                foreach (DataRow row in result.Rows)
                {
                    cabecera.CICAFECH = row["CICAFECH"] != DBNull.Value
                        ? DateTime.Parse(row["CICAFECH"].ToString()).ToString("d/M/yyyy")
                        : "";

                    cabecera.CICANCIN = row["CICANCIN"] != DBNull.Value ? row["CICANCIN"].ToString() : "";
                    cabecera.CICANOMD = row["CICANOMD"] != DBNull.Value ? row["CICANOMD"].ToString() : "";
                    cabecera.CICADIRD = row["CICADIRD"] != DBNull.Value ? row["CICADIRD"].ToString() : "";
                    cabecera.CICALOCD = row["CICALOCD"] != DBNull.Value ? row["CICALOCD"].ToString() : "";
                    cabecera.CICACOPD = row["CICACOPD"] != DBNull.Value ? row["CICACOPD"].ToString() : "";
                    cabecera.CICAPAID = row["CICAPAID"] != DBNull.Value ? row["CICAPAID"].ToString() : "";
                    cabecera.CICANOMC = row["CICANOMC"] != DBNull.Value ? row["CICANOMC"].ToString() : "";
                    cabecera.CICADIRC = row["CICADIRC"] != DBNull.Value ? row["CICADIRC"].ToString() : "";
                    cabecera.CICALOCC = row["CICALOCC"] != DBNull.Value ? row["CICALOCC"].ToString() : "";
                    cabecera.CICACOPC = row["CICACOPC"] != DBNull.Value ? row["CICACOPC"].ToString() : "";
                    cabecera.CICAPAIC = row["CICAPAIC"] != DBNull.Value ? row["CICAPAIC"].ToString() : "";
                }
            }

            return cabecera;
        }

        /// <summary>
        /// Se encarga de conseguir los datos relacionados a los registros del Invoice y devolvera una lista de RegistroModel con dichos datos.
        /// </summary>
        /// <param name="nroInvoice">Nro del Commercial Invoice del cual se daran formato los datos</param>
        /// <returns>List<RegistroModel></returns>
        public List<RegistroModel> ConseguirRegistros(string nroInvoice)
        {
            DataTable result;
            List<RegistroModel> registros = new List<RegistroModel>();

            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro("@NCIN", nroInvoice);
                result = conn.Executar_proc("DOCMAIL_REPORTE_CI_DETALLE", "COMEX");
            }

            foreach (DataRow row in result.Rows)
            {
                RegistroModel registro = new RegistroModel();

                if (decimal.TryParse(row["CIDEPEUN"].ToString(), out decimal unidad))
                {
                    registro.CIDEPEUN = unidad.ToString("#,0.00", CultureInfo.InvariantCulture).Replace(',', 'X').Replace('.', ',').Replace('X', '.');
                }
                else
                {
                    registro.CIDEPEUN = "0,00"; 
                }

                if (decimal.TryParse(row["CIDEPETO"].ToString(), out decimal precio))
                {
                    registro.CIDEPETO = precio.ToString("#,0.00", CultureInfo.InvariantCulture).Replace(',', 'X').Replace('.', ',').Replace('X', '.');
                }
                else
                {
                    registro.CIDEPETO = "0,00";
                }

                if (decimal.TryParse(row["CIDEPRUN"].ToString(), out decimal cantidadUni))
                {
                    registro.CIDEPRUN = cantidadUni.ToString("#,0.00", CultureInfo.InvariantCulture).Replace(',', 'X').Replace('.', ',').Replace('X', '.');
                }
                else
                {
                    registro.CIDEPRUN = "0,00"; 
                }

                if (decimal.TryParse(row["CIDEPRTO"].ToString(), out decimal total))
                {
                    registro.CIDEPRTO = total.ToString("#,0.00", CultureInfo.InvariantCulture).Replace(',', 'X').Replace('.', ',').Replace('X', '.');
                }
                else
                {
                    registro.CIDEPRTO = "0,00"; 
                }

                registro.CIDEMATE = row["CIDEMATE"].ToString();
                registro.CIDEDESC = row["CIDEDESC"].ToString();
                registro.CIDECANT = row["CIDECANT"].ToString(); 

                registros.Add(registro);
            }
            return registros;
        }

        /// <summary>
        /// Se encarga de conseguir los datos relacionados al subtotal del Invoice y devolvera un SubtotalModel con dichos datos.
        /// </summary>
        /// <param name="nroInvoice">Nro del Commercial Invoice del cual se daran formato los datos</param>
        /// <returns>SubtotalModel</returns>
        public SubtotalModel ConseguirSubtotal(string nroInvoice)
        {
            DataTable result;
            SubtotalModel subtotal = new SubtotalModel();

            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro("@NCIN", nroInvoice);
                result = conn.Executar_proc("DOCMAIL_REPORTE_CI_SUBTOTAL", "COMEX");
            }

            if (result.Rows.Count == 0)
            {
                // Asignar "" a las propiedades en caso de que no haya filas
                subtotal.FCA = "";
                subtotal.FLETE = "";
                subtotal.SEGURO = "";
                subtotal.TOTAL = "";
            }
            else
            {
                foreach (DataRow row in result.Rows)
                {
                    // Formatear FCA
                    if (decimal.TryParse(row["SUBTOTAL"].ToString(), out decimal subtotalValue))
                    {
                        subtotal.FCA = subtotalValue.ToString("#,0.00", CultureInfo.InvariantCulture).Replace(',', 'X').Replace('.', ',').Replace('X', '.');
                    }
                    else
                    {
                        subtotal.FCA = "0,00"; // Manejo de error
                    }

                    // Formatear FLETE
                    if (decimal.TryParse(row["FLETE"].ToString(), out decimal fleteValue))
                    {
                        subtotal.FLETE = fleteValue.ToString("#,0.00", CultureInfo.InvariantCulture).Replace(',', 'X').Replace('.', ',').Replace('X', '.');
                    }
                    else
                    {
                        subtotal.FLETE = "0,00"; // Manejo de error
                    }

                    // Formatear SEGURO
                    if (decimal.TryParse(row["SEGURO"].ToString(), out decimal seguroValue))
                    {
                        subtotal.SEGURO = seguroValue.ToString("#,0.00", CultureInfo.InvariantCulture).Replace(',', 'X').Replace('.', ',').Replace('X', '.');
                    }
                    else
                    {
                        subtotal.SEGURO = "0,00"; // Manejo de error
                    }

                    // Formatear TOTAL
                    if (decimal.TryParse(row["TOTAL"].ToString(), out decimal totalValue))
                    {
                        subtotal.TOTAL = totalValue.ToString("#,0.00", CultureInfo.InvariantCulture).Replace(',', 'X').Replace('.', ',').Replace('X', '.');
                    }
                    else
                    {
                        subtotal.TOTAL = "0,00"; // Manejo de error
                    }
                }
            }

            return subtotal;
        }


        /// <summary>
        /// Se encarga de conseguir los datos relacionados al pie del Invoice y devolvera un PieModel con dichos datos.
        /// </summary>
        /// <param name="nroInvoice">Nro del Commercial Invoice del cual se daran formato los datos</param>
        /// <returns>PieModel</returns>
        public PieModel ConseguirPie(string nroInvoice)
        {
            DataTable result;
            PieModel pie = new PieModel();

            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro("@NCIN", nroInvoice);
                result = conn.Executar_proc("DOCMAIL_REPORTE_CI_PIE", "COMEX");
            }

            if (result.Rows.Count == 0)
            {
                pie.CIPINETW = "";
                pie.CIPIGROW = "";
                pie.CIPIMEAS = "";
                pie.CIPICONT = "";
                pie.CIPIVESS = "";
                pie.CIPIOBS1 = "";
                pie.CIPIPAYM = "";
                pie.CIPICOND = "";
                pie.CIPINO11 = "";
                pie.CIPINO12 = "";
                pie.CIPIDOC1 = "";
                pie.CIPINCIN = "";
                pie.CIPIMARI = "";
            }
            else
            {
                foreach (DataRow row in result.Rows)
                {
                    pie.CIPINETW = row["CIPINETW"] != DBNull.Value ? row["CIPINETW"].ToString() : "";
                    pie.CIPIGROW = row["CIPIGROW"] != DBNull.Value ? row["CIPIGROW"].ToString() : "";
                    pie.CIPIMEAS = row["CIPIMEAS"] != DBNull.Value ? row["CIPIMEAS"].ToString() : "";
                    pie.CIPICONT = row["CIPICONT"] != DBNull.Value ? row["CIPICONT"].ToString() : "";
                    pie.CIPIVESS = row["CIPIVESS"] != DBNull.Value ? row["CIPIVESS"].ToString() : "";
                    pie.CIPIOBS1 = row["CIPIOBS1"] != DBNull.Value ? row["CIPIOBS1"].ToString() : "";
                    pie.CIPIPAYM = row["CIPIPAYM"] != DBNull.Value ? row["CIPIPAYM"].ToString() : "";
                    pie.CIPICOND = row["CIPICOND"] != DBNull.Value ? row["CIPICOND"].ToString() : "";
                    pie.CIPINO11 = row["CIPINO11"] != DBNull.Value ? row["CIPINO11"].ToString() : "";
                    pie.CIPINO12 = row["CIPINO12"] != DBNull.Value ? row["CIPINO12"].ToString() : "";
                    pie.CIPIDOC1 = row["CIPIDOC1"] != DBNull.Value ? row["CIPIDOC1"].ToString() : "";
                    pie.CIPINCIN = row["CIPINCIN"] != DBNull.Value ? row["CIPINCIN"].ToString() : "";
                    pie.CIPIMARI = row["CIPIMARI"] != DBNull.Value ? row["CIPIMARI"].ToString() : "";
                }
            }

            return pie;
        }
    }
}