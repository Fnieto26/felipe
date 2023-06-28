using Directions_Api.Helpers;
using System.Data;
using System.Diagnostics;

namespace Directions_Api.Bussines_Logic
{
    public class Seguridad
    {
        private string cadenaConexion; // = ConfigurationManager.ConnectionStrings.Item("conexion").ConnectionString()

        public Seguridad(string cadenaConexion)
        {
            this.cadenaConexion = cadenaConexion;
        }



        /// <summary>
        ///     ''' Valida el ingreso de usuario contra base de datos. Retorna el número de registros 
        ///     ''' que resulten de la validación.
        ///     ''' </summary>
        ///     ''' <param name="nombre_usuario"></param>
        ///     ''' <param name="contrasena"></param>
        ///     ''' <returns></returns>
        ///     ''' <remarks></remarks>
        public int validarUsuario(string nombre_usuario, string contrasena)
        {
            return System.Convert.ToInt32(SqlHelper.ExecuteScalar(cadenaConexion, "spConValidarUsuario", nombre_usuario, contrasena));
        }

        /// <summary>
        ///     ''' Consulta los datos de login, nombre de usuario y permiso y los retorna en un string separados por coma.
        ///     ''' </summary>
        ///     ''' <param name="nombre_usuario"></param>
        ///     ''' <returns></returns>
        ///     ''' <remarks></remarks>
        public string consultarUsuario(string nombre_usuario)
        {
            DataSet dsUsuario = SqlHelper.ExecuteDataset(cadenaConexion, "conDatosUsuario", nombre_usuario);
            
            string strUsuario = "";
            if (dsUsuario.Tables[0].Rows.Count == 1)
            {
                strUsuario += nombre_usuario + ",";
                strUsuario += dsUsuario.Tables[0].Rows[0]["usu_nombre"].ToString() + ",";
                strUsuario += dsUsuario.Tables[0].Rows[0]["usu_permiso"].ToString();
            }

            return strUsuario;
        }
    }

}
