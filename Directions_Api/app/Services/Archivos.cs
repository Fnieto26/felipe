
using ICSharpCode.SharpZipLib.Zip;
using System.Data.SqlTypes;
using System.Data;
using Directions_Api.Helpers;

namespace Directions_Api.Bussines_Logic
{

    public class Archivos
    {

        private readonly string cadenaConexion; // = ConfigurationManager.ConnectionStrings.Item("conexion").ConnectionString()
        private readonly Coordenadas oCordenada; // = ConfigurationManager.ConnectionStrings.Item("conexion").ConnectionString()

        public Archivos(string cadenaConexion, Coordenadas oCordenada)
        {
            this.cadenaConexion = cadenaConexion;
            this.oCordenada = oCordenada;
        }

        /// <summary>
        ///     ''' Funcionalidad para el proceso de descompresión de un UNICO archivo que integra todo el proceso
        ///     ''' para generar un archivo con todas las variables.
        ///     ''' </summary>
        ///     ''' <param name="rutaRecibido"></param>
        ///     ''' <param name="rutaarchivo"></param>
        ///     ''' <param name="rutaDescarga"></param>
        ///     ''' <param name="rutaGenerar"></param>
        ///     ''' <param name="eliminar"></param>
        ///     ''' <param name="renombrar"></param>
        ///     ''' <remarks></remarks>
        public void procesoArchivos(string rutaRecibido, string rutaarchivo, string rutaDescarga, string rutaGenerar, string archivo, string usuario, bool eliminar = false, bool renombrar = false)
        {
            string[] archivos;
            int i;

            // Descomprime el archivo de texto en zip dado
            descomprimirArchivo(rutaRecibido, rutaarchivo, archivo, true, true);
            archivos = Directory.GetFiles(rutaarchivo, "*.txt");
            // Este proceso se repite por cada archivo TXT existente
            for (i = 0; i <= archivos.Length - 1; i++)
            {
                // Cargar uno a uno los archivos de texto descomprimidos a la BD
                cargarArchivo(rutaarchivo, archivos[i], usuario, archivo, true, true);
                // Completar las otras columnas
                completarInformacion(usuario, archivo);
                // Sacar los datos de la BD y generar archivo .TXT
                string[] nomArch;
                nomArch = archivos[i].Split(System.Convert.ToChar(@"\"));
                generarTexto(rutaGenerar, nomArch[nomArch.Length - 1], usuario, archivo);
                // Comprimir el archivo en la ruta de salida
                comprimirArchivo(rutaGenerar, rutaDescarga, nomArch[nomArch.Length - 1], true);
            }
        }


        /// <summary>
        ///     ''' Permite comprimir un archivo de texto, en una ruta especificada.
        ///     ''' </summary>
        ///     ''' <param name="rutaGenerar"></param>
        ///     ''' <param name="rutaDescarga"></param>
        ///     ''' <param name="archivo"></param>
        ///     ''' <param name="eliminar"></param>
        ///     ''' <remarks></remarks>
        private void comprimirArchivo(string rutaGenerar, string rutaDescarga, string archivo, bool eliminar = false)
        {
            try
            {
                // Elimina el archivo si existe
                if (File.Exists(rutaDescarga + @"\" + archivo + ".zip"))
                    File.Delete(rutaDescarga + @"\" + archivo + ".zip");
                // Crea el archivo a comprimir
                string targetZipFileName = rutaDescarga + @"\" + archivo + ".zip";

                using (ZipOutputStream strmZipOutputStream = new ZipOutputStream(System.IO.File.Create(targetZipFileName)))
                {
                    strmZipOutputStream.SetLevel(9);

                    bool hasFiles = false;
                    string gRIPSZipFiles = "";

                    if (archivo.ToString().EndsWith(".TXT", StringComparison.OrdinalIgnoreCase))
                    {
                        FileStream strmFile = System.IO.File.OpenRead(rutaGenerar + @"\" + archivo);
                        byte[] abyBuffer = new byte[System.Convert.ToInt32(strmFile.Length - 1) + 1];

                        strmFile.Read(abyBuffer, 0, abyBuffer.Length);
                        ZipEntry objZipEntry = new ZipEntry(ZipEntry.CleanName(Path.GetFileName(archivo)));

                        objZipEntry.DateTime = DateTime.Now;
                        objZipEntry.Size = strmFile.Length;
                        strmFile.Close();
                        strmZipOutputStream.PutNextEntry(objZipEntry);
                        strmZipOutputStream.Write(abyBuffer, 0, abyBuffer.Length);

                        gRIPSZipFiles += Path.GetFileName(archivo) + "|";

                        hasFiles = true;
                    }
                    if (hasFiles)
                        gRIPSZipFiles = gRIPSZipFiles.Substring(0, gRIPSZipFiles.Length - 1);
                    strmZipOutputStream.Finish();
                    strmZipOutputStream.Close();
                    if (eliminar)
                        File.Delete(rutaGenerar + @"\" + archivo);
                }
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        ///     ''' Genera archivo de texto a ser comprimido.
        ///     ''' </summary>
        ///     ''' <param name="rutaarchivo"></param>
        ///     ''' <param name="archivo"></param>
        ///     ''' <remarks></remarks>
        private void generarTexto(string rutaarchivo, string archivo, string usuario, string nombre_archivo)
        {
            // Borra el archivo si existe
            if (File.Exists(rutaarchivo + @"\" + archivo))
                File.Delete(rutaarchivo + @"\" + archivo);

            System.Data.DataSet dsDireccion = consultarDireccion(usuario, nombre_archivo);
            // Crea un nuevo archivo
            StreamWriter strStreamWrite = new StreamWriter(rutaarchivo + @"\" + archivo);
            string linea = "";

            foreach (System.Data.DataRow dr in dsDireccion.Tables[0].Rows)
            {
                linea = dr["consecutivo"].ToString() + ";" + dr["identificador_direccion"].ToString() + ";" + dr["codigo_direccion"].ToString() + ";" + dr["direccion_cargada"].ToString() + ";" + dr["localidad"].ToString() + ";" + dr["upz"].ToString() + ";" + dr["barrio"].ToString() + ";" + dr["coordenada_x"].ToString() + ";" + dr["coordenada_y"].ToString() + ";" + dr["estrato"].ToString() + ";" + dr["codigo_estado"].ToString();

                strStreamWrite.WriteLine(linea);
            }
            strStreamWrite.Close();
        }

        /// <summary>
        ///     ''' Llama a la funcionalidad "coordenadas" para agregar la información con base en su consecutivo 
        ///     ''' y codigo de direccion.
        ///     ''' </summary>
        ///     ''' <remarks></remarks>
        private void completarInformacion(string nombre_usuario, string nombre_archivo)
        {
            string[] datos;
            string[] tmp;

            // Llamar a la funcion coordenadas x cada datos de la tabla TblDirecciones
            System.Data.DataSet dsDireccion = consultarDireccion(nombre_usuario, nombre_archivo);
            foreach (System.Data.DataRow dr in dsDireccion.Tables[0].Rows)
            {
                string coordenada = oCordenada.coordenadas(dr["codigo_direccion"].ToString());
                tmp = coordenada.Split(System.Convert.ToChar(";"));
                if (tmp.Length == 7)
                {
                    coordenada = dr["id_consecutivo"].ToString() + ";" + coordenada;
                    datos = coordenada.Split(System.Convert.ToChar(";"));
                    if (datos.Length == 8)
                        actualizarDireccion(datos);
                }
            }
        }

        /// <summary>
        ///     ''' Descomprime todos los archivos ZIP de una carpeta dada.
        ///     ''' </summary>
        ///     ''' <param name="rutaRecibido"></param>
        ///     ''' <param name="rutaarchivo"></param>
        ///     ''' <param name="eliminar"></param>
        ///     ''' <param name="renombrar"></param>
        ///     ''' <remarks></remarks>
        private void descomprimirArchivo(string rutaRecibido, string rutaarchivo, bool eliminar = false, bool renombrar = false)
        {
            string[] zipFic;
            int i;
            zipFic = Directory.GetFiles(rutaRecibido, "*.zip");

            for (i = 0; i <= zipFic.Length - 1; i++)
            {
                ZipInputStream z = new ZipInputStream(File.OpenRead(zipFic[i]));
                ZipEntry theEntry;

                do
                {
                    theEntry = z.GetNextEntry();
                    if (theEntry != null)
                    {
                        string fileName = rutaarchivo + @"\" + Path.GetFileName(theEntry.Name);

                        // dará error si no existe el path
                        FileStream streamWriter;
                        try
                        {
                            streamWriter = File.Create(fileName);
                        }
                        catch (DirectoryNotFoundException)
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(fileName));
                            streamWriter = File.Create(fileName);
                        }
                        // 
                        int size;
                        byte[] data = new byte[2049];
                        do
                        {
                            size = z.Read(data, 0, data.Length);
                            if ((size > 0))
                                streamWriter.Write(data, 0, size);
                            else
                                break;
                        }
                        while (true);
                        streamWriter.Close();
                    }
                    else
                        break;
                }
                while (true);
                z.Close();

                // cuando se hayan extraído los ficheros, renombrarlo
                if (renombrar)
                {
                    if (File.Exists(zipFic[i] + ".descomprimido"))
                        File.Delete(zipFic[i] + ".descomprimido");
                    File.Copy(zipFic[i], zipFic[i] + ".descomprimido");
                }
                if (eliminar)
                    File.Delete(zipFic[i]);
            }
        }

        /// <summary>
        ///     ''' Descomprime un archivo ZIP dado en una carpeta dada.
        ///     ''' </summary>
        ///     ''' <param name="rutaRecibido"></param>
        ///     ''' <param name="rutaarchivo"></param>
        ///     ''' <param name="archivo"></param>
        ///     ''' <param name="eliminar"></param>
        ///     ''' <param name="renombrar"></param>
        ///     ''' <remarks></remarks>
        private void descomprimirArchivo(string rutaRecibido, string rutaarchivo, string archivo, bool eliminar = false, bool renombrar = false)
        {
            ZipInputStream z = new ZipInputStream(File.OpenRead(rutaRecibido + @"\" + archivo + ".zip"));
            ZipEntry theEntry;

            do
            {
                theEntry = z.GetNextEntry();
                if (theEntry != null)
                {
                    string fileName = rutaarchivo + @"\" + Path.GetFileName(theEntry.Name);

                    // Dará error si no existe el path
                    FileStream streamWriter;
                    try
                    {
                        streamWriter = File.Create(fileName);
                    }
                    catch (DirectoryNotFoundException)
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(fileName));
                        streamWriter = File.Create(fileName);
                    }
                    // 
                    int size;
                    byte[] data = new byte[2049];
                    do
                    {
                        size = z.Read(data, 0, data.Length);
                        if ((size > 0))
                            streamWriter.Write(data, 0, size);
                        else
                            break;
                    }
                    while (true);
                    streamWriter.Close();
                }
                else
                    break;
            }
            while (true);
            z.Close();

            // cuando se hayan extraído los ficheros, renombrarlo
            if (renombrar)
            {
                if (File.Exists(rutaRecibido + @"\" + archivo + ".descomprimido"))
                    File.Delete(rutaRecibido + @"\" + archivo + ".descomprimido");
                File.Copy(rutaRecibido + @"\" + archivo + ".zip", rutaRecibido + @"\" + archivo + ".descomprimido");
            }
            if (eliminar)
                File.Delete(rutaRecibido + @"\" + archivo + ".zip");
        }

        /// <summary>
        ///     ''' Carga un archivo a la base de datos en la tabla "TblDirecciones". Únicamente carga llave y codigo de dirección.
        ///     ''' </summary>
        ///     ''' <param name="rutaarchivo"></param>
        ///     ''' <param name="archivo"></param>
        ///     ''' <param name="eliminar"></param>
        ///     ''' <param name="renombrar"></param>
        ///     ''' <remarks></remarks>
        private void cargarArchivo(string rutaarchivo, string archivo, string usuario, string nombre_archivo_control, bool eliminar = false, bool renombrar = false)
        {
            string linea = "";
            string[] param;
            StreamReader strStreamRead;

            // Limpia la tabla para poder ingresar datos de un nuevo archivo
            borrarDireccion(usuario);

            // Abrir el archivo y leerlo
            strStreamRead = new StreamReader(archivo);

            while (linea != null)
            {
                linea = strStreamRead.ReadLine();
                if (linea != null)
                {
                    // Procesar la carga hacia la BD
                    param = linea.Split(System.Convert.ToChar(";"));
                    if (param.Length == 5)
                        // TODO: Inserta con el nombre de archivo y usuario.
                        insertarDireccion(param[0], param[1], param[2], param[3], usuario, nombre_archivo_control);
                }
            }
            strStreamRead.Close();
            // cuando se hayan extraído los ficheros, renombrarlo
            if (renombrar)
            {
                if (File.Exists(archivo + ".cargado"))
                    File.Delete(archivo + ".cargado");
                File.Copy(archivo, archivo + ".cargado");
            }
            if (eliminar)
                File.Delete(archivo);
        }

        public bool ingresarControlArchivo(string usuario, string nombre_archivo, long tamano_archivo, int aproximacion)
        {
            if (SqlHelper.ExecuteNonQuery(cadenaConexion, "spInsControl", usuario, nombre_archivo, tamano_archivo, aproximacion) > 0)
                return true;
            else
                return false;
        }

        public bool actualizarEstadoArchivo(string usuario, string nombre_archivo, string archivo_descarga, int estado)
        {
            if (SqlHelper.ExecuteNonQuery(cadenaConexion, "spUpdControl", usuario, nombre_archivo, archivo_descarga, estado) > 0)
                return true;
            else
                return false;
        }

        public bool actualizarEstadoArchivo(string usuario, string archivo_descarga, int estado)
        {
            if (SqlHelper.ExecuteNonQuery(cadenaConexion, "spUpdEstadoControl", usuario, archivo_descarga, estado) > 0)
                return true;
            else
                return false;
        }

        public DataSet consultarArchivosProcesados(string usuario)
        {
            return SqlHelper.ExecuteDataset(cadenaConexion, "spConArchivosProcesados", usuario);
        }

        /// <summary>
        ///     ''' Retorna un data set con todas las direcciones cargadas en la tabla "TblDirecciones".
        ///     ''' </summary>
        ///     ''' <returns></returns>
        ///     ''' <remarks></remarks>
        private System.Data.DataSet consultarDireccion(string nombre_usuario, string nombre_archivo)
        {
            return SqlHelper.ExecuteDataset(cadenaConexion, "spConDireccion", nombre_usuario, nombre_archivo);
        }

        /// <summary>
        ///     ''' Insertar los datos de llave y código de direccion en la tabla de direcciones.
        ///     ''' </summary>
        ///     ''' <param name="parametros"></param>
        ///     ''' <returns></returns>
        ///     ''' <remarks></remarks>
        private bool insertarDireccion(params string[] parametros)
        {
            if (SqlHelper.ExecuteNonQuery(cadenaConexion, "spInsArchivo", parametros) > 0)
                return true;
            else
                return false;
        }

        /// <summary>
        ///     ''' Actualizar los datos adicionales de la dirección como Localidad, UPZ, Código de Barrio, 
        ///     ''' Código Estado.
        ///     ''' </summary>
        ///     ''' <param name="parametros"></param>
        ///     ''' <returns></returns>
        ///     ''' <remarks></remarks>
        private bool actualizarDireccion(params string[] parametros)
        {
            if (SqlHelper.ExecuteNonQuery(cadenaConexion, "spUpdDireccion", parametros) > 0)
                return true;
            else
                return false;
        }

        /// <summary>
        ///     ''' Borrar la tabla "TblDirecciones" para poder generar cargar las nuevas direcciones.
        ///     ''' </summary>
        ///     ''' <remarks></remarks>
        private void borrarDireccion(string nombre_usuario)
        {
            SqlHelper.ExecuteNonQuery(cadenaConexion, "spDelDireccion", nombre_usuario);
        }

        /// <summary>
        ///     ''' Consultar los archivos en estado 1 - Cargados de la tabla de control.
        ///     ''' </summary>
        ///     ''' <returns></returns>
        ///     ''' <remarks></remarks>
        public System.Data.DataSet consultarArchivosCargados()
        {
            return SqlHelper.ExecuteDataset(cadenaConexion, "spConControlEstado", 1);
        }
    }

}
