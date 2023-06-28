Public Class seguridad
    Private _cadenaConexion As String ' = ConfigurationManager.ConnectionStrings.Item("conexion").ConnectionString()

    Public WriteOnly Property cadenaConexion() As String
        Set(ByVal value As String)
            _cadenaConexion = value
        End Set
    End Property

    ''' <summary>
    ''' Valida el ingreso de usuario contra base de datos. Retorna el número de registros 
    ''' que resulten de la validación.
    ''' </summary>
    ''' <param name="nombre_usuario"></param>
    ''' <param name="contrasena"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function validarUsuario(ByVal nombre_usuario As String, ByVal contrasena As String) As Integer
        Return CInt(SqlHelper.ExecuteScalar(_cadenaConexion, "spConValidarUsuario", nombre_usuario, contrasena))
    End Function

    ''' <summary>
    ''' Consulta los datos de login, nombre de usuario y permiso y los retorna en un string separados por coma.
    ''' </summary>
    ''' <param name="nombre_usuario"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function consultarUsuario(ByVal nombre_usuario As String) As String
        Dim dsUsuario As DataSet = SqlHelper.ExecuteDataset(_cadenaConexion, "conDatosUsuario", nombre_usuario)
        Dim strUsuario As String = ""
        If dsUsuario.Tables(0).Rows.Count = 1 Then
            strUsuario &= nombre_usuario + ","
            strUsuario &= dsUsuario.Tables(0).Rows(0)("usu_nombre").ToString + ","
            strUsuario &= dsUsuario.Tables(0).Rows(0)("usu_permiso").ToString
        End If

        Return strUsuario
    End Function
End Class
