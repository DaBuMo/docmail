using DOCMAIL.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace DOCMAIL.Services.Data
{
    public class InvoiceDAO
    {
        public CabeceraModel ConseguirCabecera(string nroInvoice)
        {
            DataTable result;
            CabeceraModel cabecera = new CabeceraModel();
            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro("@NCIN", nroInvoice);

                result = conn.Executar_proc("DOCMAIL_REPORTE_CI_CABECERA", "COMEX");
            }
            foreach (DataRow row in result.Rows)
            {
                cabecera.CICAFECH = DateTime.Parse(row["CICAFECH"].ToString()).ToString("M/d/yyyy");
                cabecera.CICANCIN = row["CICANCIN"].ToString();

                cabecera.CICANOMD = row["CICANOMD"].ToString();
                cabecera.CICADIRD = row["CICADIRD"].ToString();
                cabecera.CICALOCD = row["CICALOCD"].ToString();
                cabecera.CICACOPD = row["CICACOPD"].ToString();
                cabecera.CICAPAID = row["CICAPAID"].ToString();

                cabecera.CICANOMC = row["CICANOMC"].ToString();
                cabecera.CICADIRC = row["CICADIRC"].ToString();
                cabecera.CICALOCC = row["CICALOCC"].ToString();
                cabecera.CICACOPC = row["CICACOPC"].ToString();
                cabecera.CICAPAIC = row["CICAPAIC"].ToString();
            }

            return cabecera;
        }

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

                registro.CIDECANT = row["CIDECANT"].ToString();
                registro.CIDEMATE = row["CIDEMATE"].ToString();
                registro.CIDEDESC = row["CIDEDESC"].ToString();
                registro.CIDEPEUN = row["CIDEPEUN"].ToString();
                registro.CIDEPETO = row["CIDEPETO"].ToString();
                registro.CIDEPRUN = row["CIDEPRUN"].ToString();
                registro.CIDEPRTO = row["CIDEPRTO"].ToString();

                registros.Add(registro);
            }

            return registros;
        }

        public SubtotalModel ConseguirSubtotal(string nroInvoice)
        {
            DataTable result;
            SubtotalModel subtotal = new SubtotalModel();

            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro("@NCIN", nroInvoice);

                result = conn.Executar_proc("DOCMAIL_REPORTE_CI_SUBTOTAL", "COMEX");
            }

            foreach (DataRow row in result.Rows)
            {
                subtotal.FCA = row["SUBTOTAL"].ToString();
                subtotal.FLETE = row["FLETE"].ToString();
                subtotal.SEGURO = row["SEGURO"].ToString();
                subtotal.TOTAL = row["TOTAL"].ToString();

            }

            return subtotal;
        }

        public PieModel ConseguirPie(string nroInvoice)
        {
            DataTable result;
            PieModel pie = new PieModel();

            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro("@NCIN", nroInvoice);

                result = conn.Executar_proc("DOCMAIL_REPORTE_CI_PIE", "COMEX");
            }

            foreach (DataRow row in result.Rows)
            {
                pie.CIPINETW = row["CIPINETW"].ToString();
                pie.CIPIGROW = row["CIPIGROW"].ToString();
                pie.CIPIMEAS = row["CIPIMEAS"].ToString();
                pie.CIPICONT = row["CIPICONT"].ToString();
                pie.CIPIVESS = row["CIPIVESS"].ToString();
                pie.CIPIOBS1 = row["CIPIOBS1"].ToString();
                pie.CIPIPAYM = row["CIPIPAYM"].ToString();
                pie.CIPICOND = row["CIPICOND"].ToString();
                pie.CIPINO11 = row["CIPINO11"].ToString();
                pie.CIPINO12 = row["CIPINO12"].ToString();
                pie.CIPIDOC1 = row["CIPIDOC1"].ToString();
                pie.CIPINCIN = row["CIPINCIN"].ToString();
                pie.CIPIMARI = row["CIPIMARI"].ToString();
            }
            return pie;  
        }
    }
}