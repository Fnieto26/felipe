Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports codificadorDireccionesLogica

<System.Web.Services.WebService(Namespace:="http://200.75.49.126/direccion")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Direccion
    Inherits System.Web.Services.WebService

    Private cadenaConexion As String = ConfigurationManager.ConnectionStrings.Item("conexion").ConnectionString()
    Private oValida As New ValidaDireccion()
    Private oSeguridad As New seguridad()
    Private oArchivo As New archivos()

    Private tamanoArchivoBytes As String = ConfigurationManager.AppSettings.Item("tamano_archivo_bytes").ToString
    Private rutaRecibido As String = ConfigurationManager.AppSettings.Item("ruta_recibido").ToString
    Private rutaDescarga As String = ConfigurationManager.AppSettings.Item("ruta_descarga").ToString
    Private rutaarchivo As String = ConfigurationManager.AppSettings.Item("ruta_archivo").ToString
    Private rutaGenerar As String = ConfigurationManager.AppSettings.Item("ruta_generar").ToString

    <WebMethod(Description:="Retorna el código de una dirección.")> _
    Public Function obtenerCodDireccion(ByVal Direccion As String, ByVal usuario As String, ByVal clave As String) As String
        Dim valor_direccion As String
        oValida.cadenaConexion = cadenaConexion
        valor_direccion = oValida.limpia(Direccion)
        valor_direccion = oValida.ejes(valor_direccion)
        valor_direccion = oValida.validaEjePrincipal(valor_direccion)

        Dim oCoordenada As New coordenadas()
        oCoordenada.cadenaConexion = cadenaConexion
        Return valor_direccion & ";" & oCoordenada.coordenadas(valor_direccion)
    End Function

    <WebMethod(Description:="Retorna la dirección de un numero de telefono dado.")> _
    Public Function obtenerDireccionPorTelefono(ByVal telefono As String, ByVal usuario As String, ByVal clave As String) As String
        Dim oValidaDireccion As New ValidaDireccion
        oValidaDireccion.cadenaConexion = cadenaConexion
        Return oValidaDireccion.telefono(telefono)
    End Function

    <WebMethod(Description:="Valida un usuario. Retorna True si el usuario es valido.")> _
    Public Function validarUsuario(ByVal usuario As String, ByVal contrasena As String) As Boolean
        oSeguridad.cadenaConexion = cadenaConexion
        Select Case oSeguridad.validarUsuario(usuario, contrasena)
            Case 0
                Return False
            Case 1
                Return True
            Case Else
                Return False
        End Select
    End Function

    <WebMethod(Description:="Consulta los datos de un usuario.")> _
    Public Function consultarUsuario(ByVal usuario As String) As String
        oSeguridad.cadenaConexion = cadenaConexion
        Return oSeguridad.consultarUsuario(usuario)
    End Function

    <WebMethod(Description:="Ingresa el registro de control de un archivo enviado. Retorna un valor booleano si se debe procesar el archivo en línea.")> _
    Public Function ingresarControlArchivo(ByVal usuario As String, ByVal contrasena As String, ByVal nombre_archivo As String, ByVal tamano_archivo As Long, ByVal aproximacion As Integer) As Boolean
        Dim b_generado As Boolean = False
        oArchivo.cadenaConexion = cadenaConexion
        'Ingresa los datos a la tabla.
        oArchivo.ingresarControlArchivo(usuario, nombre_archivo, tamano_archivo, aproximacion)
        If IsNumeric(tamanoArchivoBytes) AndAlso tamano_archivo <= CLng(tamanoArchivoBytes) Then
            'Procesa el archivo automáticamente ya que cumple con el tamaño especifico.
            b_generado = True
            oArchivo.procesoArchivos(rutaRecibido, rutaarchivo, rutaDescarga, rutaGenerar, nombre_archivo, usuario, True, True)
            oArchivo.actualizarEstadoArchivo(usuario, nombre_archivo, nombre_archivo & ".txt", 2)
        End If
        Return b_generado
    End Function

    <WebMethod(Description:="Consulta los archivos cuyo estado sea procesado para un usuario dado. Retorna una cadena con los nombres de los archivos separado por coma.")> _
    Public Function consultarArchivosProcesado(ByVal usuario As String) As String
        Dim archivos As String = String.Empty
        oArchivo.cadenaConexion = cadenaConexion
        Dim dsUsuario As Data.DataSet = oArchivo.consultarArchivosProcesados(usuario)
        If dsUsuario.Tables(0).Rows.Count > 0 Then
            For Each dr As Data.DataRow In dsUsuario.Tables(0).Rows
                archivos &= dr("archivo_descarga").ToString & ";"
            Next
            archivos = archivos.Substring(0, archivos.Length - 1)
        End If
        Return archivos
    End Function

    <WebMethod(Description:="Consulta los parametros de configuración del servicio FTP para descarga de archivos.")> _
    Public Function consultarConfiguracionFTP() As Data.DataSet
        Dim oConfiguracion As New configuracion
        oConfiguracion.cadenaConexion = cadenaConexion
        Return oConfiguracion.consultarConfiguracionFTP()
    End Function

    <WebMethod(Description:="Actualiza el estado de archivo a estado descargado.")> _
    Public Function actualizarEstadoArchivo(ByVal nombre_usuario As String, ByVal archivo_descarga As String) As Boolean
        oArchivo.cadenaConexion = cadenaConexion
        Return oArchivo.actualizarEstadoArchivo(nombre_usuario, archivo_descarga, 3)

    End Function
End Class