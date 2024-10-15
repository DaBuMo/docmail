using DOCMAIL.Models;
using DOCMAIL.Services.Business;
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;


namespace DOCMAIL.Services.Data
{
    public class ProcessingDAO
    {
        private string linea;

        /// <summary>
        /// Esta funcion busca en la BD todos los registros que cumplan con el tipo de registro dado y posean, si es que se indica, el mismo numero o uno similar al brindado en nroRegistro.
        /// </summary>
        /// <param name="tipo">Descripción del primer parámetro.</param>
        /// <param name="tipoRegistro">Tipo de registro a buscar (0 = PENDIENTES // 1 = ENVIADOS )</param>
        /// <param name="nroInvoice">Nro parcial o completo del Commercial Invoice a buscar. Si este valor es "" entonces se listaran todos los invoices segun el valor de tipoRegistro</param>
        /// <returns>DataTable</returns>
        internal DataTable ListarCommercialInvoices(string tipoRegistro, string nroInvoice)
        {
            DataTable result = new DataTable();

            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro("@ENVIADA", tipoRegistro);
                conn.Criar_parametro("@NCIN", nroInvoice);
                result = conn.Executar_proc("DOCMAIL_CONSULTA_PENDIENTES", "COMEX");

            }
            return result;
        }

        /// <summary>
        /// Esta funcion se encarga de la obtencion de la carpeta en donde se almacenaran las Commercial Invoices por procesar" 
        /// </summary>
        /// <returns>String</returns>
        internal string recuperarCarpetaOrigenInvoices()
        {
            using (Conexao conn = new Conexao())
            {
                return conn.Executar_proc("DOCMAIL_CARPETA_SAP", "COMEX").Rows[0]["CARPETA_ACCESO"].ToString();
            }
        }


        /// <summary>
        /// Esta funcion se encarga de la lectura, procesamiento y dada de alta de las Commercial Invoices traidas desde SAP y almacenadas en la carpeta dada por el startprocedure = "DOCMAIL_CARPETA_SAP" 
        /// </summary>
        /// <param name="origen">Carpeta de almacenamiento de los Commercial Invoices ( Para conseguir directorio de esta carpeta ejecutar SP = "DOCMAIL_CARPETA_SAP" )</param>
        /// <param name="destino">Carpeta de almacenamiento de los Commercial Invoices procesados y dados de alta</param>
        /// <returns></returns>
        internal void ProcessComercialInvoices(string origen, string destino)
        {
            using (Conexao conn = new Conexao())
            {
                // Obtención de archivos pendientes a procesar
                var carpeta = Directory.GetFiles(origen);

                foreach (string archivo in carpeta)
                {
                    using (var lector = new StreamReader(archivo, Encoding.UTF8))
                    {

                        while ((linea = lector.ReadLine()) != null)
                        {
                            conn.Executar_comando($"INSERT INTO DMAUIMPO(IMPOTEXT) SELECT '{linea}'", "COMEX");
                        }
                    }
                    reubicacionProcesado(archivo, destino);
                }
                // Damos de alta todos los archivos procesados
                conn.Executar_proc("DOCMAIL_ALTA_CI", "COMEX");
            }

        }

        /// <summary>
        /// Esta funcion se encarga reubicar el archivo procesado a la carpeta correspondiende" 
        /// <param name="archivo">Nombre original del archivo procesado</param>
        /// <param name="destino">Carpeta de almacenamiento de los Commercial Invoices procesados y dados de alta</param>
        /// <returns></returns>
        internal void reubicacionProcesado(string archivo, string destino)
        {
            string archivoDestino = aplicarFormatoGuardado(archivo, destino);
            try
            {
                if (File.Exists(archivoDestino))
                {
                    File.Delete(archivoDestino);
                }
                File.Move(archivo, archivoDestino);
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error de I/O: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Esta funcion se encarga de aplicar el formato de guardado en un archivo procesado: Indicando el dia, mes y año en el que este archivo fue procesado y ademas agrega el destino al que este archivo sera enviado ya que fue procesado" 
        /// </summary>
        /// <param name="archivo">Nombre original del archivo procesado</param>
        /// <param name="destino">Carpeta de almacenamiento de los Commercial Invoices procesados y dados de alta</param>
        /// <returns>String</returns>
        private string aplicarFormatoGuardado(string archivo, string destino)
        {
            string fechaHora = DateTime.Now.ToString("yyyyMMddHHmm");
            string nuevoNombre = $"{Path.GetFileNameWithoutExtension(archivo)}_{fechaHora}{Path.GetExtension(archivo)}";
            return Path.Combine(destino, nuevoNombre);
        }

        /// <summary>
        /// Esta funcion se encarga aumentar la cantidad de envios de un Invoice espeficio" 
        /// </summary>
        /// <param name="nroInvoice">Nro del Invoice descargado</param>
        /// <returns></returns>
        internal void CambiarNroEnvios(string nroInvoice)
        {
            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro("@NCIN", nroInvoice);
                conn.Executar_proc("DOCMAIL_CI_UPDATE_ENVIOS", "COMEX");
            }
        }
    }
}