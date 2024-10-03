using DOCMAIL.Services.Data;
using System;
using System.Data;
using System.IO;


namespace DOCMAIL.Services.Business
{
    public class ProcessingService
    {
        private readonly ProcessingDAO processingDAO = new ProcessingDAO();

        /// <summary>
        /// Se comunica con processingDAO para llamar a la busqueda de todos los registros que cumplan con el mismo tipo de registro y tengan el mismo o posean un Nro de registro similiar.
        /// </summary>
        /// <param name="tipoRegistro">Tipo de registro a buscar (0 = PENDIENTES // 1 = ENVIADOS )</param>
        /// <param name="nroInvoice">Nro parcial o completo del Commercial Invoice a buscar. Si este valor es "" entonces se listaran todos los invoices segun el valor de tipoRegistro</param>
        /// <returns>DataTable</returns>
        public DataTable Listar(string tipoRegistro,string nroInvoice)
        {
            return processingDAO.ListarCommercialInvoices(tipoRegistro, nroInvoice);
        }

        /// <summary>
        /// Se comunica con el ProcessingDAO para llamar al procesamiento y dado de alta de Commercial Invoices traidos desde SAP.
        /// </summary>
        /// <returns></returns>
        public void Procesar()
        {
            string origen = processingDAO.recuperarCarpetaOrigenInvoices();
            string destino = RecuperarCarpetaDestino(origen);
            processingDAO.ProcessComercialInvoices(origen,destino);
        }

        /// <summary>
        /// Establece la carpeta en la cual se guardaran los Commercial Invoices procesados, esta suele ser la primera carpeta que lea en el directorio en el que se almacenen los Commercial Invoices.
        /// </summary>
        /// <param name="origen">Carpeta de almacenamiento de los Commercial Invoices ( Para conseguir directorio de esta carpeta ejecutar SP = "DOCMAIL_CARPETA_SAP" )</param>
        /// <returns>String</returns>
        private string RecuperarCarpetaDestino(string origen)
        {
            string carpeta_procesados = Path.GetFileName(Directory.GetDirectories(origen)[0]);
            return Path.Combine(origen, carpeta_procesados);
        }


    }

    
}