Imports ICSharpCode.SharpZipLib
Imports ICSharpCode.SharpZipLib.Checksums
Imports System.IO
Imports System.Data

Public Class archivos

    Private _cadenaConexion As String '= ConfigurationManager.ConnectionStrings.Item("conexion").ConnectionString()

    Public WriteOnly Property cadenaConexion() As String
        Set(ByVal value As String)
            _cadenaConexion = value
        End Set
    End Property

    ''' <summary>
    ''' Funcionalidad para el proceso de descompresión de un UNICO archivo que integra todo el proceso
    ''' para generar un archivo con todas las variables.
    ''' </summary>
    ''' <param name="rutaRecibido"></param>
    ''' <param name="rutaarchivo"></param>
    ''' <param name="rutaDescarga"></param>
    ''' <param name="rutaGenerar"></param>
    ''' <param name="eliminar"></param>
    ''' <param name="renombrar"></param>
    ''' <remarks></remarks>
    Public Sub procesoArchivos(ByVal rutaRecibido As String, _
                                ByVal rutaarchivo As String, _
                                ByVal rutaDescarga As String, _
                                ByVal rutaGenerar As String, _
                                ByVal archivo As String, _
                                ByVal usuario As String, _
                                Optional ByVal eliminar As Boolean = False, _
                                Optional ByVal renombrar As Boolean = False)

        Dim archivos() As String
        Dim i As Integer

        'Descomprime el archivo de texto en zip dado
        descomprimirArchivo(rutaRecibido, rutaarchivo, archivo, True, True)
        archivos = Directory.GetFiles(rutaarchivo, "*.txt")
        'Este proceso se repite por cada archivo TXT existente
        For i = 0 To archivos.Length - 1
            'Cargar uno a uno los archivos de texto descomprimidos a la BD
            cargarArchivo(rutaarchivo, archivos(i), usuario, archivo, True, True)
            'Completar las otras columnas
            completarInformacion(usuario, archivo)
            'Sacar los datos de la BD y generar archivo .TXT
            Dim nomArch() As String
            nomArch = archivos(i).Split(CChar("\"))
            generarTexto(rutaGenerar, nomArch(nomArch.Length - 1), usuario, archivo)
            'Comprimir el archivo en la ruta de salida
            comprimirArchivo(rutaGenerar, rutaDescarga, nomArch(nomArch.Length - 1), True)
        Next
    End Sub


    ''' <summary>
    ''' Permite comprimir un archivo de texto, en una ruta especificada.
    ''' </summary>
    ''' <param name="rutaGenerar"></param>
    ''' <param name="rutaDescarga"></param>
    ''' <param name="archivo"></param>
    ''' <param name="eliminar"></param>
    ''' <remarks></remarks>
    Private Sub comprimirArchivo(ByVal rutaGenerar As String, _
                                ByVal rutaDescarga As String, _
                                ByVal archivo As String, _
                                Optional ByVal eliminar As Boolean = False)

        Try
            'Elimina el archivo si existe
            If File.Exists(rutaDescarga & "\" & archivo & ".zip") Then
                File.Delete(rutaDescarga & "\" & archivo & ".zip")
            End If
            'Crea el archivo a comprimir
            Dim targetZipFileName As String = rutaDescarga & "\" & archivo & ".zip"

            Using strmZipOutputStream As New Zip.ZipOutputStream(System.IO.File.Create(targetZipFileName))
                strmZipOutputStream.SetLevel(9)

                Dim hasFiles As Boolean = False
                Dim gRIPSZipFiles As String = ""

                If archivo.ToString.EndsWith(".TXT", StringComparison.OrdinalIgnoreCase) Then
                    Dim strmFile As FileStream = System.IO.File.OpenRead(rutaGenerar & "\" & archivo)
                    Dim abyBuffer(CInt(strmFile.Length - 1)) As Byte

                    strmFile.Read(abyBuffer, 0, abyBuffer.Length)
                    Dim objZipEntry As New Zip.ZipEntry(Zip.ZipEntry.CleanName(Path.GetFileName(archivo)))

                    objZipEntry.DateTime = DateTime.Now
                    objZipEntry.Size = strmFile.Length
                    strmFile.Close()
                    strmZipOutputStream.PutNextEntry(objZipEntry)
                    strmZipOutputStream.Write(abyBuffer, 0, abyBuffer.Length)

                    gRIPSZipFiles &= Path.GetFileName(archivo) & "|"

                    hasFiles = True
                End If
                If hasFiles Then
                    gRIPSZipFiles = gRIPSZipFiles.Substring(0, gRIPSZipFiles.Length - 1)
                End If
                strmZipOutputStream.Finish()
                strmZipOutputStream.Close()
                If eliminar Then
                    File.Delete(rutaGenerar & "\" & archivo)
                End If
            End Using
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Genera archivo de texto a ser comprimido.
    ''' </summary>
    ''' <param name="rutaarchivo"></param>
    ''' <param name="archivo"></param>
    ''' <remarks></remarks>
    Private Sub generarTexto(ByVal rutaarchivo As String, _
                                ByVal archivo As String, ByVal usuario As String, ByVal nombre_archivo As String)
        'Borra el archivo si existe
        If File.Exists(rutaarchivo & "\" & archivo) Then
            File.Delete(rutaarchivo & "\" & archivo)
        End If

        Dim dsDireccion As Data.DataSet = consultarDireccion(usuario, nombre_archivo)
        'Crea un nuevo archivo
        Dim strStreamWrite As New StreamWriter(rutaarchivo & "\" & archivo)
        Dim linea As String = ""

        For Each dr As Data.DataRow In dsDireccion.Tables(0).Rows
            linea = dr("consecutivo").ToString & ";" & dr("identificador_direccion").ToString & ";" & _
                    dr("codigo_direccion").ToString & ";" & dr("direccion_cargada").ToString & ";" & _
                    dr("localidad").ToString & ";" & dr("upz").ToString & ";" & _
                    dr("barrio").ToString & ";" & dr("coordenada_x").ToString & ";" & _
                    dr("coordenada_y").ToString & ";" & dr("estrato").ToString & ";" & dr("codigo_estado").ToString

            strStreamWrite.WriteLine(linea)
        Next
        strStreamWrite.Close()
    End Sub

    ''' <summary>
    ''' Llama a la funcionalidad "coordenadas" para agregar la información con base en su consecutivo 
    ''' y codigo de direccion.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub completarInformacion(ByVal nombre_usuario As String, ByVal nombre_archivo As String)
        Dim datos() As String
        Dim tmp() As String

        'Llamar a la funcion coordenadas x cada datos de la tabla TblDirecciones
        Dim dsDireccion As Data.DataSet = consultarDireccion(nombre_usuario, nombre_archivo)
        For Each dr As Data.DataRow In dsDireccion.Tables(0).Rows            
            Dim oCordenada As New coordenadas
            oCordenada.cadenaConexion = _cadenaConexion
            Dim coordenada As String = oCordenada.coordenadas(dr("codigo_direccion").ToString)
            tmp = coordenada.Split(CChar(";"))
            If tmp.Length = 7 Then
                coordenada = dr("id_consecutivo").ToString & ";" & coordenada
                datos = coordenada.Split(CChar(";"))
                If datos.Length = 8 Then
                    actualizarDireccion(datos)
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Descomprime todos los archivos ZIP de una carpeta dada.
    ''' </summary>
    ''' <param name="rutaRecibido"></param>
    ''' <param name="rutaarchivo"></param>
    ''' <param name="eliminar"></param>
    ''' <param name="renombrar"></param>
    ''' <remarks></remarks>
    Private Sub descomprimirArchivo(ByVal rutaRecibido As String, _
                                ByVal rutaarchivo As String, _
                                Optional ByVal eliminar As Boolean = False, _
                                Optional ByVal renombrar As Boolean = False)

        Dim zipFic() As String
        Dim i As Integer
        zipFic = Directory.GetFiles(rutaRecibido, "*.zip")

        For i = 0 To zipFic.Length - 1
            Dim z As New Zip.ZipInputStream(File.OpenRead(zipFic(i)))
            Dim theEntry As Zip.ZipEntry

            Do
                theEntry = z.GetNextEntry()
                If Not theEntry Is Nothing Then
                    Dim fileName As String = rutaarchivo & "\" & Path.GetFileName(theEntry.Name)

                    ' dará error si no existe el path
                    Dim streamWriter As FileStream
                    Try
                        streamWriter = File.Create(fileName)
                    Catch ex As DirectoryNotFoundException
                        Directory.CreateDirectory(Path.GetDirectoryName(fileName))
                        streamWriter = File.Create(fileName)
                    End Try
                    '
                    Dim size As Integer
                    Dim data(2048) As Byte
                    Do
                        size = z.Read(data, 0, data.Length)
                        If (size > 0) Then
                            streamWriter.Write(data, 0, size)
                        Else
                            Exit Do
                        End If
                    Loop
                    streamWriter.Close()
                Else
                    Exit Do
                End If
            Loop
            z.Close()

            ' cuando se hayan extraído los ficheros, renombrarlo
            If renombrar Then
                If File.Exists(zipFic(i) & ".descomprimido") Then
                    File.Delete(zipFic(i) & ".descomprimido")
                End If
                File.Copy(zipFic(i), zipFic(i) & ".descomprimido")
            End If
            If eliminar Then
                File.Delete(zipFic(i))
            End If
        Next
    End Sub

    ''' <summary>
    ''' Descomprime un archivo ZIP dado en una carpeta dada.
    ''' </summary>
    ''' <param name="rutaRecibido"></param>
    ''' <param name="rutaarchivo"></param>
    ''' <param name="archivo"></param>
    ''' <param name="eliminar"></param>
    ''' <param name="renombrar"></param>
    ''' <remarks></remarks>
    Private Sub descomprimirArchivo(ByVal rutaRecibido As String, _
                                ByVal rutaarchivo As String, _
                                ByVal archivo As String, _
                                Optional ByVal eliminar As Boolean = False, _
                                Optional ByVal renombrar As Boolean = False)

        Dim z As New Zip.ZipInputStream(File.OpenRead(rutaRecibido & "\" & archivo & ".zip"))
        Dim theEntry As Zip.ZipEntry

        Do
            theEntry = z.GetNextEntry()
            If Not theEntry Is Nothing Then
                Dim fileName As String = rutaarchivo & "\" & Path.GetFileName(theEntry.Name)

                ' Dará error si no existe el path
                Dim streamWriter As FileStream
                Try
                    streamWriter = File.Create(fileName)
                Catch ex As DirectoryNotFoundException
                    Directory.CreateDirectory(Path.GetDirectoryName(fileName))
                    streamWriter = File.Create(fileName)
                End Try
                '
                Dim size As Integer
                Dim data(2048) As Byte
                Do
                    size = z.Read(data, 0, data.Length)
                    If (size > 0) Then
                        streamWriter.Write(data, 0, size)
                    Else
                        Exit Do
                    End If
                Loop
                streamWriter.Close()
            Else
                Exit Do
            End If
        Loop
        z.Close()

        ' cuando se hayan extraído los ficheros, renombrarlo
        If renombrar Then
            If File.Exists(rutaRecibido & "\" & archivo & ".descomprimido") Then
                File.Delete(rutaRecibido & "\" & archivo & ".descomprimido")
            End If
            File.Copy(rutaRecibido & "\" & archivo & ".zip", rutaRecibido & "\" & archivo & ".descomprimido")
        End If
        If eliminar Then
            File.Delete(rutaRecibido & "\" & archivo & ".zip")
        End If
    End Sub

    ''' <summary>
    ''' Carga un archivo a la base de datos en la tabla "TblDirecciones". Únicamente carga llave y codigo de dirección.
    ''' </summary>
    ''' <param name="rutaarchivo"></param>
    ''' <param name="archivo"></param>
    ''' <param name="eliminar"></param>
    ''' <param name="renombrar"></param>
    ''' <remarks></remarks>
    Private Sub cargarArchivo(ByVal rutaarchivo As String, _
                                    ByVal archivo As String, _
                                    ByVal usuario As String, _
                                    ByVal nombre_archivo_control As String, _
                                    Optional ByVal eliminar As Boolean = False, _
                                    Optional ByVal renombrar As Boolean = False)

        Dim linea As String = ""
        Dim param() As String
        Dim strStreamRead As StreamReader

        'Limpia la tabla para poder ingresar datos de un nuevo archivo
        borrarDireccion(usuario)

        'Abrir el archivo y leerlo
        strStreamRead = New StreamReader(archivo)

        While Not linea Is Nothing
            linea = strStreamRead.ReadLine()
            If Not linea Is Nothing Then
                'Procesar la carga hacia la BD
                param = linea.Split(CChar(";"))
                If param.Length = 5 Then
                    'TODO: Inserta con el nombre de archivo y usuario.
                    insertarDireccion(param(0), param(1), param(2), param(3), usuario, nombre_archivo_control)
                End If
            End If
        End While
        strStreamRead.Close()
        ' cuando se hayan extraído los ficheros, renombrarlo
        If renombrar Then
            If File.Exists(archivo & ".cargado") Then
                File.Delete(archivo & ".cargado")
            End If
            File.Copy(archivo, archivo & ".cargado")
        End If
        If eliminar Then
            File.Delete(archivo)
        End If
    End Sub

#Region "Datos"
    Public Function ingresarControlArchivo(ByVal usuario As String, ByVal nombre_archivo As String, ByVal tamano_archivo As Long, ByVal aproximacion As Integer) As Boolean
        If SqlHelper.ExecuteNonQuery(_cadenaConexion, "spInsControl", _
                                            usuario, _
                                            nombre_archivo, _
                                            tamano_archivo, _
                                            aproximacion) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function actualizarEstadoArchivo(ByVal usuario As String, ByVal nombre_archivo As String, ByVal archivo_descarga As String, ByVal estado As Integer) As Boolean
        If SqlHelper.ExecuteNonQuery(_cadenaConexion, "spUpdControl", _
                                        usuario, _
                                        nombre_archivo, _
                                        archivo_descarga, _
                                        estado) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function actualizarEstadoArchivo(ByVal usuario As String, ByVal archivo_descarga As String, ByVal estado As Integer) As Boolean
        If SqlHelper.ExecuteNonQuery(_cadenaConexion, "spUpdEstadoControl", _
                                                usuario, _
                                                archivo_descarga, _
                                                estado) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function consultarArchivosProcesados(ByVal usuario As String) As DataSet
        Return SqlHelper.ExecuteDataset(_cadenaConexion, "spConArchivosProcesados", _
                                         usuario)
    End Function

    ''' <summary>
    ''' Retorna un data set con todas las direcciones cargadas en la tabla "TblDirecciones".
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function consultarDireccion(ByVal nombre_usuario As String, ByVal nombre_archivo As String) As Data.DataSet
        Return SqlHelper.ExecuteDataset(_cadenaConexion, "spConDireccion", nombre_usuario, nombre_archivo)
    End Function

    ''' <summary>
    ''' Insertar los datos de llave y código de direccion en la tabla de direcciones.
    ''' </summary>
    ''' <param name="parametros"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function insertarDireccion(ByVal ParamArray parametros() As String) As Boolean
        If SqlHelper.ExecuteNonQuery(_cadenaConexion, "spInsArchivo", parametros) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Actualizar los datos adicionales de la dirección como Localidad, UPZ, Código de Barrio, 
    ''' Código Estado.
    ''' </summary>
    ''' <param name="parametros"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function actualizarDireccion(ByVal ParamArray parametros() As String) As Boolean
        If SqlHelper.ExecuteNonQuery(_cadenaConexion, "spUpdDireccion", parametros) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Borrar la tabla "TblDirecciones" para poder generar cargar las nuevas direcciones.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub borrarDireccion(ByVal nombre_usuario As String)
        SqlHelper.ExecuteNonQuery(_cadenaConexion, "spDelDireccion", nombre_usuario)
    End Sub

    ''' <summary>
    ''' Consultar los archivos en estado 1 - Cargados de la tabla de control.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function consultarArchivosCargados() As Data.DataSet
        Return SqlHelper.ExecuteDataset(_cadenaConexion, "spConControlEstado", 1)
    End Function
#End Region
End Class
