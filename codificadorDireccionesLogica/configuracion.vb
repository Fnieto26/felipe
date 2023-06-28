Public Class configuracion

    Private _cadenaConexion As String ' = ConfigurationManager.ConnectionStrings.Item("conexion").ConnectionString()

    Public WriteOnly Property cadenaConexion() As String
        Set(ByVal value As String)
            _cadenaConexion = value
        End Set
    End Property

    Public Function consultarConfiguracionFTP() As Data.DataSet
        Return SqlHelper.ExecuteDataset(_cadenaConexion, "spConConfiguracion", "FTP")
    End Function

End Class
