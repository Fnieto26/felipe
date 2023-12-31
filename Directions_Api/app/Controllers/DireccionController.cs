using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualBasic;
using Directions_Api.Bussines_Logic;
using Directions_Api.Helpers;
using Directions_Api.Models.Requests;
using Directions_Api.Models.Responses;
using System.Data;
using Microsoft.AspNetCore.Http;

namespace Directions_Api.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class DireccionController : ControllerBase
    {


        private readonly ValidaDireccion oValida; // lleva cadena de conexion
        private readonly Seguridad oSeguridad; // lleva cadena de conexion
        private readonly Archivos oArchivo; // lleva cadena de conexion
        private readonly Coordenadas oCoordenada; // lleva cadena de conexion;
        private readonly Configuration config;

        public DireccionController(ValidaDireccion oValida, Seguridad oSeguridad, Archivos oArchivo, Coordenadas oCoordenada, Configuration config)
        {
            this.oValida = oValida;
            this.oSeguridad = oSeguridad;
            this.oArchivo = oArchivo;
            this.oCoordenada = oCoordenada;
            this.config = config;
        }


        /// <summary>
        /// Retorna el c�digo de una direcci�n
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult<ResponseBase<string>> obtenerCodDireccion(string Direccion, string usuario, string clave)
        {
            ResponseObtenerCodDireccionDTO response = new ResponseObtenerCodDireccionDTO();
            response.ObtenerCodDireccion = oValida.limpia(Direccion);
            response.ObtenerCodDireccion = oValida.ejes(response.ObtenerCodDireccion);
            response.ObtenerCodDireccion = oValida.validaEjePrincipal(response.ObtenerCodDireccion);
            response.ObtenerCodDireccion = response.ObtenerCodDireccion + ";" + oCoordenada.coordenadas(response.ObtenerCodDireccion);
            return Ok(response);
        }

        /// <summary>
        /// Retorna la direcci�n de un numero de telefono dado.
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult<ResponseBase<string>> obtenerDireccionPorTelefono(string telefono, string usuario, string clave)
        {
            ResponseObtenerDireccionPorTelefonoDTO result = new ResponseObtenerDireccionPorTelefonoDTO();
            result.ObtenerDireccionPorTelefono = oValida.telefono(telefono);
            return Ok(result);
        }

        /// <summary>
        /// Valida un usuario. 
        /// </summary>
        /// <returns>Retorna True si el usuario es valido. False en caso contrario</returns>
        [HttpPost]
        public ActionResult<ResponseBase<bool>> validarUsuario(ValidarUsuarioRequest validarUsuarioRequest)
        {
            ResponseValidarUsuarioDTO result = new ResponseValidarUsuarioDTO();

            switch (oSeguridad.validarUsuario(validarUsuarioRequest.Usuario, validarUsuarioRequest.Contrasena))
            {
                case 0:
                    {
                        return Ok(result);
                    }

                case 1:
                    {
                        result.ValidarUsuario = true;
                        return Ok(result);
                    }

                default:
                    {
                        return Ok(result);
                    }
            }
        }

        /// <summary>
        /// Consulta los datos de un usuario.
        /// </summary>
        /// <param name="usuario"></param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult<ResponseBase<string>> consultarUsuario(string usuario)
        {
            ResponseConsultarUsuarioDTO result = new ResponseConsultarUsuarioDTO();
            result.ConsultarUsuario = oSeguridad.consultarUsuario(usuario);
            return Ok(result);
        }

        /// <summary>
        /// Ingresa el registro de control de un archivo enviado. 
        /// </summary>
        /// <returns>Retorna un valor booleano si se debe procesar el archivo en l�nea.</returns>
        [HttpPost]
        public ActionResult<ResponseBase<bool>> ingresarControlArchivo(IngresarControlArchivoRequest ingresarControlArchivoRequest)
        {
            ResponseIngresarControlArchivoDTO result = new ResponseIngresarControlArchivoDTO();
            // Ingresa los datos a la tabla.
            oArchivo.ingresarControlArchivo(ingresarControlArchivoRequest.Usuario, ingresarControlArchivoRequest.NombreArchivo, ingresarControlArchivoRequest.TamanoArchivo, ingresarControlArchivoRequest.Aproximacion);
            if (Information.IsNumeric(config.tamanoArchivoBytes) && ingresarControlArchivoRequest.TamanoArchivo <= System.Convert.ToInt64(config.tamanoArchivoBytes))
            {
                // Procesa el archivo autom�ticamente ya que cumple con el tama�o especifico.
                result.IngresarControlArchivo = true;
                oArchivo.procesoArchivos(config.rutaRecibido, config.rutaarchivo, config.rutaDescarga, config.rutaGenerar, ingresarControlArchivoRequest.NombreArchivo, ingresarControlArchivoRequest.Usuario, true, true);
                oArchivo.actualizarEstadoArchivo(ingresarControlArchivoRequest.Usuario, ingresarControlArchivoRequest.NombreArchivo, ingresarControlArchivoRequest.NombreArchivo + ".txt", 2);
            }
            return Ok(result);
        }

        /// <summary>
        /// "Consulta los archivos cuyo estado sea procesado para un usuario dado. 
        /// Retorna una cadena con los nombres de los archivos separado por coma.
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult<ResponseBase<string>> consultarArchivosProcesado(string usuario)
        {
            ResponseConsultarArchivosProcesadosDTO result = new ResponseConsultarArchivosProcesadosDTO();
            result.ConsultarArchivosProcesados = "";
            System.Data.DataSet dsUsuario = oArchivo.consultarArchivosProcesados(usuario);
            if (dsUsuario.Tables[0].Rows.Count > 0)
            {
                foreach (System.Data.DataRow dr in dsUsuario.Tables[0].Rows)
                    result.ConsultarArchivosProcesados += dr["archivo_descarga"].ToString() + ";";
                result.ConsultarArchivosProcesados = result.ConsultarArchivosProcesados.Substring(0, result.ConsultarArchivosProcesados.Length - 1);
            }
            return Ok(result);
        }

        /// <summary>
        /// Consulta los parametros de configuraci�n del servicio FTP para descarga de archivos.
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult<ResponseBase<List<FTPConfigurationResponse>>> consultarConfiguracionFTP()
        {
            var result = new ResponseBase<List<FTPConfigurationResponse>>();
            result.consultarConfiguracionFTP = new List<FTPConfigurationResponse>();
            var data = SqlHelper.ExecuteDataset(config.cadenaConexion, "spConConfiguracion", "FTP");
            foreach (DataRow row in data.Tables[0].Rows)
            {
                var ftpConfigurationResponse = new FTPConfigurationResponse();
                ftpConfigurationResponse.Llave = row["llave"].ToString();
                ftpConfigurationResponse.Valor = row["valor"].ToString();
                result.consultarConfiguracionFTP.Add(ftpConfigurationResponse);
            }
            return Ok(result);
        }

        /// <summary>
        /// Actualiza el estado de archivo a estado descargado.
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult<ResponseBase<bool>> actualizarEstadoArchivo(ActualizarEstadoArchivoRequest actualizarEstadoArchivoRequest)
        {
            ResponseActualizarEstadoArchivoDTO result = new ResponseActualizarEstadoArchivoDTO();
            result.ActualizarEstadoArchivo = oArchivo.actualizarEstadoArchivo(actualizarEstadoArchivoRequest.NombreUsuario, actualizarEstadoArchivoRequest.ArchivoDescarga, 3);
            return Ok(result);
        }
    }
}