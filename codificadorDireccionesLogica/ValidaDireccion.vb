Public Class ValidaDireccion

#Region "Atributos"
    Private _cadenaConexion As String ' = ConfigurationManager.ConnectionStrings.Item("conexion").ConnectionString()
    Private _Sur As Integer = 0
    Private _Este As Integer = 0
#End Region

#Region "Propiedades"
    Public WriteOnly Property cadenaConexion() As String
        Set(ByVal value As String)
            _cadenaConexion = value
        End Set
    End Property

    Public ReadOnly Property Sur() As Integer
        Get
            Return _Sur
        End Get
    End Property

    Public ReadOnly Property Este() As Integer
        Get
            Return _Este
        End Get
    End Property
#End Region

    Public Function obtenerCodDirecion(ByVal Direccion As String) As String
        Direccion = limpia(Direccion)
        Direccion = ejes(Direccion)

        Return validaEjePrincipal(Direccion)
    End Function


    ''' <summary>
    ''' Funcion que permite limpiar la direccion
    ''' </summary>
    ''' <param name="Direccion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function limpia(ByVal Direccion As String) As String
        Dim StrDirec As String
        'recibir parametro de la cadena  PARAMETRO DE ENTRADA
        StrDirec = Direccion.ToUpper
        StrDirec = StrDirec.Replace("-", " ")

        'BUSCA # N_ y varios antes de los remplazos en la tabla
        Dim dsSignos As Data.DataSet = SqlHelper.ExecuteDataset(_cadenaConexion, "conSignos")
        Dim drSignos As Data.DataRow
        For Each drSignos In dsSignos.Tables(0).Rows
            StrDirec = StrDirec.Replace(drSignos("val_errores").ToString, drSignos("cargar").ToString)
        Next

        'Separa numeros de letras
        Dim i As Integer
        Dim num As String = ""
        Dim letra As String = ""
        Dim Caraceje1 As String = ""
        Dim numF As Boolean = False
        Dim letraF As Boolean = False
        Dim cambioF As Boolean = False

        StrDirec = StrDirec.Replace("  ", " ")
        Direccion = ""

        For i = 0 To StrDirec.Length - 1
            Caraceje1 = StrDirec.Substring(i, 1)
            If cambioF Then
                If numF Then
                    Direccion = Direccion & " " & letra
                    letra = ""
                Else
                    Direccion = Direccion & " " & num
                    num = ""
                End If
                cambioF = False
            End If

            If IsNumeric(Caraceje1) Then
                num = num & Caraceje1
                If letraF Then
                    cambioF = True
                End If
                numF = True
                letraF = False
            Else
                letra = letra & Caraceje1
                If numF Then
                    cambioF = True
                End If
                letraF = True
                numF = False
            End If
        Next
        If numF Then
            Direccion = Direccion & " " & letra & " " & num
            letra = ""
        Else
            Direccion = Direccion & " " & num & " " & letra
            num = ""
        End If
        Direccion = Direccion.TrimStart

        Return Direccion
    End Function

    Public Function ejes(ByVal Direccion As String) As String
        Dim posicionE, posicionS As Integer
        _Sur = 0
        _Este = 0
        posicionS = 0
        Direccion = Direccion.ToUpper
        posicionS = Direccion.IndexOf("SUR")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace("SUR", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If
        Direccion = Direccion.Replace("  ", " ")
        posicionS = 0
        posicionS = Direccion.IndexOf("S UR")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace("S UR", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" S U R")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" S U R", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" SU R")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" SU R", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" SU R")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" SU R", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" SUS")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" SUS", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" SUC")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" SUC", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" SIR")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" SIR", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" SU ")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" SU ", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" DUR ")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" DUR ", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" AUR ")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" AUR ", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" SR ")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" SR ", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionS = 0
        posicionS = Direccion.IndexOf(" S ")
        If posicionS > 0 Then
            _Sur = 1
            Direccion = Direccion.Replace(" S ", " ")
            posicionS = 0
            Direccion = Direccion.Replace("  ", " ")
        End If
        Direccion = Direccion.ToUpper
        posicionE = Direccion.IndexOf("ESTE")
        If posicionE > 0 Then
            _Este = 1
            Direccion = Direccion.Replace("ESTE", "")
            posicionE = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionE = Direccion.IndexOf(" EST")
        If posicionE > 0 Then
            _Este = 1
            Direccion = Direccion.Replace(" EST", "")
            posicionE = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionE = Direccion.IndexOf(" ETE")
        If posicionE > 0 Then
            _Este = 1
            Direccion = Direccion.Replace(" ETE", "")
            posicionE = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionE = Direccion.IndexOf(" ETE")
        If posicionE > 0 Then
            _Este = 1
            Direccion = Direccion.Replace(" ETE", " ")
            posicionE = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionE = Direccion.IndexOf(" E STE ")
        If posicionE > 0 Then
            _Este = 1
            Direccion = Direccion.Replace(" E STE ", " ")
            posicionE = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionE = Direccion.IndexOf(" ESTS ")
        If posicionE > 0 Then
            _Este = 1
            Direccion = Direccion.Replace(" ESTS", "  ")
            posicionE = 0
            Direccion = Direccion.Replace("  ", " ")
        End If

        posicionE = Direccion.IndexOf(" ES")
        If posicionE > 0 Then
            _Este = 1
            Direccion = Direccion.Replace(" ES ", " ")
            posicionE = 0
            Direccion = Direccion.Replace("  ", " ")
        End If
        'Nuevo copiar para servicio ============
        Direccion = Direccion.Replace("BIS", " BIS ")
        Direccion = Direccion.Replace(" BI ", " BIS ")
        Direccion = Direccion.Replace(" BI S ", " BIS ")
        Direccion = Direccion.Replace(" B I S ", " BIS ")
        Direccion = Direccion.Replace(" B IS ", " BIS ")
        Direccion = Direccion.Replace(" BS ", " BIS ")
        '=====================
        Direccion = Direccion.Replace("  ", " ")
        Return Direccion
    End Function

    Public Function validaEjePrincipal(ByVal Direccion As String) As String
        '=========== ini
        Dim StrDir2 As String
        Dim i As Integer
        Dim vecCampos() As String
        vecCampos = Direccion.Split(CChar(" "))

        Dim dsEje As Data.DataSet = SqlHelper.ExecuteDataset(_cadenaConexion, "conEje")
        Dim drEje As Data.DataRow
        For Each drEje In dsEje.Tables(0).Rows
            If vecCampos(0).Trim = drEje("Des_EjePrin").ToString Then
                vecCampos(0) = drEje("CodEje").ToString
                Select Case vecCampos(0)
                    Case "CL", "KR", "TV", "AV", "DG", "AC", "AK"
                        Exit For
                End Select
            End If
        Next
        StrDir2 = ""
        If vecCampos(0).Equals("CL") Or vecCampos(0).Equals("KR") Or vecCampos(0).Equals("TV") Or vecCampos(0).Equals("DG") Or vecCampos(0).Equals("AC") Or vecCampos(0).Equals("AK") Then
            For i = 0 To vecCampos.Length - 1
                StrDir2 = StrDir2 & " " & vecCampos(i) & " "
            Next
            StrDir2 = StrDir2.Replace("  ", " ")
            StrDir2 = StrDir2.Trim
            StrDir2 = StrDir2.Replace("  ", " ")
            'MessageBox.Show(StrDir2)
            Return validaTipo1(StrDir2)
        End If
        StrDir2 = ""

        Dim dsEjeAvenida As Data.DataSet = SqlHelper.ExecuteDataset(_cadenaConexion, "conValoresAvenida")
        If vecCampos(0).Equals("AV") Then
            For Each drEje In dsEje.Tables(0).Rows
                If vecCampos(1).Trim = drEje("Des_EjePrin").ToString Then
                    vecCampos(1) = drEje("CodEje").ToString
                End If
                Select Case vecCampos(1)
                    Case "CL"
                        'MessageBox.Show("es avenida calle StrDir2")
                        vecCampos(0) = "AC"
                        vecCampos(1) = ""
                        For i = 0 To vecCampos.Length - 1
                            StrDir2 = StrDir2 & " " & vecCampos(i) & " "
                        Next
                        StrDir2 = StrDir2.Replace("  ", " ")
                        StrDir2 = StrDir2.Trim
                        StrDir2 = StrDir2.Replace("  ", " ")
                        Return validaTipo1(StrDir2)
                        Exit For
                    Case "KR"
                        vecCampos(0) = "AK"
                        vecCampos(1) = ""
                        For i = 0 To vecCampos.Length - 1
                            StrDir2 = StrDir2 & " " & vecCampos(i) & " "
                        Next
                        StrDir2 = StrDir2.Replace("  ", " ")
                        StrDir2 = StrDir2.Trim
                        StrDir2 = StrDir2.Replace("  ", " ")
                        Return validaTipo1(StrDir2)
                        Exit For
                        'evaluar las avenidas numericas
                    Case "100"
                        vecCampos(0) = "AC"
                        vecCampos(1) = "100"
                        For i = 0 To vecCampos.Length - 1
                            StrDir2 = StrDir2 & " " & vecCampos(i) & " "
                        Next
                        StrDir2 = StrDir2.Replace("  ", " ")
                        StrDir2 = StrDir2.Trim
                        StrDir2 = StrDir2.Replace("  ", " ")
                        Return validaTipo1(StrDir2)
                        Exit For

                    Case "68"
                        'evalua si es sur de la 68 
                        If _Sur = 1 Then
                            vecCampos(0) = "AK"
                            vecCampos(1) = "068"
                            For i = 0 To vecCampos.Length - 1
                                StrDir2 = StrDir2 & " " & vecCampos(i) & " "
                            Next
                            StrDir2 = StrDir2.Replace("  ", " ")
                            StrDir2 = StrDir2.Trim
                            StrDir2 = StrDir2.Replace("  ", " ")
                            Return validaTipo1(StrDir2)
                            Exit For
                        Else
                            Return "verifique si es AC o AK"
                            Exit For
                        End If
                    Case "19"
                        'evalua si es la 19 
                        If IsNumeric(vecCampos(2)) Then
                            If CInt(vecCampos(2)) > 50 Then
                                vecCampos(0) = "AK"
                                For i = 0 To vecCampos.Length - 1
                                    StrDir2 = StrDir2 & " " & vecCampos(i) & " "
                                Next
                                StrDir2 = StrDir2.Replace("  ", " ")
                                StrDir2 = StrDir2.Trim
                                StrDir2 = StrDir2.Replace("  ", " ")
                                Return validaTipo1(StrDir2)
                                Exit For
                            Else
                                vecCampos(0) = "AC"
                                For i = 0 To vecCampos.Length - 1
                                    StrDir2 = StrDir2 & " " & vecCampos(i) & " "
                                Next
                                StrDir2 = StrDir2.Replace("  ", " ")
                                StrDir2 = StrDir2.Trim
                                StrDir2 = StrDir2.Replace("  ", " ")
                                Return validaTipo1(StrDir2)
                                Exit For
                            End If
                        End If
                        'Case Else
                        '    If IsNumeric(vecCampos(1)) Then
                        '        Return "es AV pero debe ser AC o AK"
                        '        Exit For
                        '    Else
                        '        Return validaTipo2(StrDir2)
                        '        Exit For
                        '    End If
                        '    'evaluar las otras cadenas av mayo
                End Select
            Next
        End If
        Return validaTipo2(Direccion)
    End Function

    Private Function validaTipo1(ByVal Direccion As String) As String
        Dim i As Integer
        '=======
        Dim vecCampos(), b As String
        Dim c, a1, c1, a As String
        '=======
        Dim vecNumeros(12) As Integer
        Dim veclongitud(12) As Integer
        Dim RecoVecto As Integer = -1
        Dim continum As Integer = -1
        Dim CodDir As String = ""
        Dim bandcambio As Integer = 1
        Dim cadnum0, cadnum1, cadnum2, cadnum3, cadnum4 As Boolean
        Dim cadnum5, cadnum6, cadnum7, cadnum8, cadnum9, bis As Boolean
        Dim eje1 As String = CStr(0)
        Dim recalvec As Integer = 0

        vecCampos = Direccion.Split(CChar(" "))
        'CARGA EL VECTOR vecnumeros con valores
        For i = 0 To vecCampos.Length - 1
            cadnum0 = vecCampos(i).Contains(CStr(0))
            cadnum1 = vecCampos(i).Contains(CStr(1))
            cadnum2 = vecCampos(i).Contains(CStr(2))
            cadnum3 = vecCampos(i).Contains(CStr(3))
            cadnum4 = vecCampos(i).Contains(CStr(4))
            cadnum5 = vecCampos(i).Contains(CStr(5))
            cadnum6 = vecCampos(i).Contains(CStr(6))
            cadnum7 = vecCampos(i).Contains(CStr(7))
            cadnum8 = vecCampos(i).Contains(CStr(8))
            cadnum9 = vecCampos(i).Contains(CStr(9))
            bis = vecCampos(i).Contains("BIS")
            If cadnum0 = True Then
                vecNumeros(i) = 1
                recalvec = i
            End If
            If cadnum1 = True Then
                vecNumeros(i) = 1
                recalvec = i
            End If
            If cadnum2 = True Then
                vecNumeros(i) = 1
                recalvec = i
            End If
            If cadnum3 = True Then
                vecNumeros(i) = 1
                recalvec = i
            End If
            If cadnum4 = True Then
                vecNumeros(i) = 1
                recalvec = i
            End If
            If cadnum5 = True Then
                vecNumeros(i) = 1
                recalvec = i
            End If
            If cadnum6 = True Then
                vecNumeros(i) = 1
                recalvec = i
            End If
            If cadnum7 = True Then
                vecNumeros(i) = 1
                recalvec = i
            End If
            If cadnum8 = True Then
                vecNumeros(i) = 1
                recalvec = i
            End If
            If cadnum9 = True Then
                vecNumeros(i) = 1
                recalvec = i
            End If
            If bis = True Then
                vecNumeros(i) = 3
            End If
            If vecNumeros(0) <> 0 And i = 0 Then
                'TODO: MessageBox.Show("error de cadena en el eje")
            End If
        Next

        'recalculo el vector numeros
        ReDim Preserve vecCampos(recalvec)
        'Genera vector de longitudes
        For i = 0 To recalvec
            veclongitud(i) = vecCampos(i).Length
        Next

        'TODO: MAnejo de error de rutina con valor cortarlo
        If vecNumeros(0) = 1 Then
            CodDir = "61 ERROR DE VIA"
            Return "62 ERROR EJE PRINCIPAL"
            'MessageBox.Show("error de cadena en el eje")
            'termina aplicativo
        End If
        '===========
        If vecNumeros(1) = 1 And vecNumeros(2) = 1 And vecNumeros(3) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(2)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(3)) > 99 Then
                Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(2) = 1 And vecNumeros(4) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(2)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(4)) > 99 Then
                Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(2) = 1 And vecNumeros(5) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(2)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(5)) > 99 Then
                'Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(2) = 1 And vecNumeros(6) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(2)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(6)) > 99 Then
                'Return "63"
            End If
        End If

        If vecNumeros(1) = 1 And vecNumeros(3) = 1 And vecNumeros(4) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(3)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(4)) > 99 Then
                'Return "63"
            End If
        End If

        If vecNumeros(1) = 1 And vecNumeros(4) = 1 And vecNumeros(5) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(4)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(5)) > 99 Then
                'Return "63"
            End If
        End If

        If vecNumeros(1) = 1 And vecNumeros(5) = 1 And vecNumeros(6) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(5)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(6)) > 99 Then
                'Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(3) = 1 And vecNumeros(5) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(3)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(5)) > 99 Then
                'Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(4) = 1 And vecNumeros(6) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(4)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(6)) > 99 Then
                'Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(5) = 1 And vecNumeros(7) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(5)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(7)) > 99 Then
                'Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(3) = 1 And vecNumeros(6) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(3)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(6)) > 99 Then
                'Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(4) = 1 And vecNumeros(7) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(4)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(7)) > 99 Then
                'Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(5) = 1 And vecNumeros(8) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(5)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(8)) > 99 Then
                'Return "63"
            End If
        End If

        If vecNumeros(1) = 1 And vecNumeros(3) = 1 And vecNumeros(7) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(3)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(7)) > 99 Then
                'Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(4) = 1 And vecNumeros(8) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(4)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(8)) > 99 Then
                'Return "63"
            End If
        End If
        If vecNumeros(1) = 1 And vecNumeros(5) = 1 And vecNumeros(9) = 1 Then
            If CInt(vecCampos(1)) > 250 Then
                Return "61"
            End If
            If CInt(vecCampos(5)) > 199 Then
                Return "62"
            End If
            If CInt(vecCampos(9)) > 99 Then
                'Return "63"
            End If
        End If

        'remplazar cadena 0 por valor de la tabla si lo encuentra trae campo en cl
        If vecCampos(0) = "CL" Then
            If _Sur = 1 Then
                If _Este = 1 Then
                    eje1 = "32"
                Else
                    eje1 = "22"
                End If
            Else
                If _Este = 1 Then
                    eje1 = "12"
                Else
                    eje1 = "02"
                End If
            End If
        End If
        If vecCampos(0) = "DG" Then
            If _Sur = 1 Then
                If _Este = 1 Then
                    eje1 = "33"
                Else
                    eje1 = "23"
                End If
            Else
                If _Este = 1 Then
                    eje1 = "13"
                Else
                    eje1 = "03"
                End If
            End If
        End If
        If vecCampos(0) = "AC" Then
            If _Sur = 1 Then
                If _Este = 1 Then
                    eje1 = "34"
                Else
                    eje1 = "24"
                End If
            Else
                If _Este = 1 Then
                    eje1 = "14"
                Else
                    eje1 = "04"
                End If
            End If
        End If

        If vecCampos(0) = "KR" Then
            If _Sur = 1 Then
                If _Este = 1 Then
                    eje1 = "35"
                Else
                    eje1 = "25"
                End If
            Else
                If _Este = 1 Then
                    eje1 = "15"
                Else
                    eje1 = "05"
                End If
            End If
        End If

        If vecCampos(0) = "TV" Then
            If _Sur = 1 Then
                If _Este = 1 Then
                    eje1 = "36"
                Else
                    eje1 = "26"
                End If
            Else
                If _Este = 1 Then
                    eje1 = "16"
                Else
                    eje1 = "06"
                End If
            End If
        End If
        If vecCampos(0) = "AK" Then
            If _Sur = 1 Then
                If _Este = 1 Then
                    eje1 = "37"
                Else
                    eje1 = "27"
                End If
            Else
                If _Este = 1 Then
                    eje1 = "17"
                Else
                    eje1 = "07"
                End If
            End If
        End If

        If vecNumeros(1) = 1 And vecNumeros(2) = 1 And vecNumeros(3) = 1 Then
            vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
            vecCampos(1) = vecCampos(1).Trim

            vecCampos(2) = vecCampos(2).PadLeft(3, CChar("0"))
            vecCampos(2) = vecCampos(2).Trim

            vecCampos(3) = vecCampos(3).PadLeft(3, CChar("0"))
            vecCampos(3) = vecCampos(3).Trim
            b = eje1.Trim & vecCampos(1).Trim & "-" & "0" & "-" & vecCampos(2).Trim & "-" & "0" & "-" & vecCampos(3).Trim
            Return b
            eje1 = ""
        End If

        If vecNumeros(1) = 1 And vecNumeros(3) = 1 And vecNumeros(4) = 1 Then
            If veclongitud(2) = 1 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim & vecCampos(2).TrimEnd

                vecCampos(3) = vecCampos(3).PadLeft(3, CChar("0"))
                vecCampos(3) = vecCampos(3).Trim

                vecCampos(4) = vecCampos(4).PadLeft(3, CChar("0"))
                vecCampos(4) = vecCampos(4).Trim
                b = eje1.Trim & vecCampos(1).Trim & "0" & "-" & vecCampos(3).Trim & "-" & "0" & "-" & vecCampos(4).Trim
                Return b
                eje1 = ""
            End If
            '=====
            If veclongitud(2) = 3 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim

                vecCampos(3) = vecCampos(3).PadLeft(3, CChar("0"))
                vecCampos(3) = vecCampos(3).Trim

                vecCampos(4) = vecCampos(4).PadLeft(3, CChar("0"))
                vecCampos(4) = vecCampos(4).Trim
                b = eje1.Trim & vecCampos(1).Trim & "-" & "1" & "-" & vecCampos(3).Trim & "-" & "0" & "-" & vecCampos(4).Trim
                Return b
                eje1 = ""

            End If
        End If

        If vecNumeros(1) = 1 And vecNumeros(2) = 1 And vecNumeros(4) = 1 Then
            If veclongitud(3) = 1 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim

                vecCampos(2) = vecCampos(2).PadLeft(3, CChar("0"))
                vecCampos(2) = vecCampos(2).Trim & vecCampos(3).TrimEnd

                vecCampos(4) = vecCampos(4).PadLeft(3, CChar("0"))
                vecCampos(4) = vecCampos(4).Trim
                b = eje1.Trim & vecCampos(1).Trim & "-" & "0" & "-" & vecCampos(2).Trim & "0" & "-" & vecCampos(4).Trim
                Return b
                eje1 = ""
            End If

            If veclongitud(3) = 3 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim

                vecCampos(2) = vecCampos(2).PadLeft(3, CChar("0"))
                vecCampos(2) = vecCampos(2).Trim

                vecCampos(4) = vecCampos(4).PadLeft(3, CChar("0"))
                vecCampos(4) = vecCampos(4).Trim
                b = eje1.Trim & vecCampos(1).Trim & "-" & "0" & "-" & vecCampos(2).Trim & "-" & "1" & "-" & vecCampos(4).Trim
                Return b
                eje1 = ""
            End If
        End If

        If vecNumeros(1) = 1 And vecNumeros(4) = 1 And vecNumeros(5) = 1 Then
            If veclongitud(2) = 1 And veclongitud(3) = 3 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim & vecCampos(2).TrimEnd

                vecCampos(4) = vecCampos(4).PadLeft(3, CChar("0"))
                vecCampos(4) = vecCampos(4).Trim

                vecCampos(5) = vecCampos(5).PadLeft(3, CChar("0"))
                vecCampos(5) = vecCampos(5).Trim
                b = eje1.Trim & vecCampos(1).Trim & "1" & "-" & vecCampos(4).Trim & "-" & "0" & "-" & vecCampos(5).Trim
                Return b
                eje1 = ""
            End If

            If veclongitud(2) = 3 And veclongitud(3) = 1 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim

                vecCampos(4) = vecCampos(4).PadLeft(3, CChar("0"))
                vecCampos(4) = vecCampos(3).Trim & vecCampos(4).Trim

                vecCampos(5) = vecCampos(5).PadLeft(3, CChar("0"))
                vecCampos(5) = vecCampos(5).Trim
                b = eje1.Trim & vecCampos(1).Trim & "-" & "1" & vecCampos(4).Trim & "-" & "0" & "-" & vecCampos(5).Trim
                Return b
                eje1 = ""
            End If
        End If

        If vecNumeros(1) = 1 And vecNumeros(5) = 1 And vecNumeros(6) = 1 Then
            If veclongitud(2) = 1 And veclongitud(3) = 3 And veclongitud(4) = 1 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim & vecCampos(2).TrimEnd

                vecCampos(5) = vecCampos(5).PadLeft(3, CChar("0"))
                vecCampos(5) = vecCampos(4).Trim & vecCampos(5).TrimEnd

                vecCampos(6) = vecCampos(6).PadLeft(3, CChar("0"))
                vecCampos(6) = vecCampos(6).Trim
                b = eje1.Trim & vecCampos(1).Trim & "1" & vecCampos(5).Trim & "-" & "0" & "-" & vecCampos(6).Trim
                Return b
                eje1 = ""
            End If

        End If

        If vecNumeros(1) = 1 And vecNumeros(2) = 1 And vecNumeros(5) = 1 Then
            If veclongitud(3) = 1 And veclongitud(4) = 3 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim

                vecCampos(2) = vecCampos(2).PadLeft(3, CChar("0"))
                vecCampos(2) = vecCampos(2).Trim & vecCampos(3).TrimEnd

                vecCampos(5) = vecCampos(5).PadLeft(3, CChar("0"))
                vecCampos(5) = vecCampos(5).Trim
                b = eje1.Trim & vecCampos(1).Trim & "-" & "0" & "-" & vecCampos(2).Trim & "1" & "-" & vecCampos(5).Trim
                Return b
                eje1 = ""
            End If

            If veclongitud(3) = 3 And veclongitud(4) = 1 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim

                vecCampos(2) = vecCampos(2).PadLeft(3, CChar("0"))
                vecCampos(2) = vecCampos(2).Trim

                vecCampos(5) = vecCampos(5).PadLeft(3, CChar("0"))
                vecCampos(5) = vecCampos(4).Trim & vecCampos(5).Trim
                b = eje1.Trim & vecCampos(1).Trim & "-" & "0" & "-" & vecCampos(2).Trim & "-" & "1" & vecCampos(5).Trim
                Return b
                eje1 = ""
            End If
        End If

        If vecNumeros(1) = 1 And vecNumeros(2) = 1 And vecNumeros(6) = 1 Then
            If veclongitud(3) = 1 And veclongitud(4) = 3 And veclongitud(5) = 1 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim

                vecCampos(2) = vecCampos(2).PadLeft(3, CChar("0"))
                vecCampos(2) = vecCampos(2).Trim & vecCampos(3).TrimEnd

                vecCampos(6) = vecCampos(6).PadLeft(3, CChar("0"))
                vecCampos(6) = vecCampos(5).Trim & vecCampos(6).TrimEnd
                b = eje1.Trim & vecCampos(1).Trim & "-" & "0" & "-" & vecCampos(2).Trim & "1" & vecCampos(6).Trim
                Return b
                eje1 = ""
            End If
        End If

        If vecNumeros(1) = 1 And vecNumeros(3) = 1 And vecNumeros(5) = 1 Then
            If veclongitud(2) = 1 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim & vecCampos(2).TrimEnd
                a = eje1.Trim & vecCampos(1).Trim & "0" & "-"
            End If
            If veclongitud(2) = 3 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim
                a1 = eje1.Trim & vecCampos(1).Trim & "-" & "1" & "-"
            End If
            If veclongitud(4) = 1 Then
                vecCampos(3) = vecCampos(3).PadLeft(3, CChar("0"))
                vecCampos(3) = vecCampos(3).Trim & vecCampos(4).TrimEnd

                vecCampos(5) = vecCampos(5).PadLeft(3, CChar("0"))
                vecCampos(5) = vecCampos(5).Trim
                c = vecCampos(3).Trim & "0" & "-" & vecCampos(5).Trim
            End If
            If veclongitud(4) = 3 Then
                vecCampos(3) = vecCampos(3).PadLeft(3, CChar("0"))
                vecCampos(3) = vecCampos(3).Trim

                vecCampos(5) = vecCampos(5).PadLeft(3, CChar("0"))
                vecCampos(5) = vecCampos(5).Trim
                c1 = vecCampos(3).Trim & "-" & "1" & "-" & vecCampos(5).Trim
            End If
            If veclongitud(2) = 1 And veclongitud(4) = 1 Then
                b = a.Trim & c.Trim
                Return b
                eje1 = ""
            End If
            If veclongitud(2) = 1 And veclongitud(4) = 3 Then
                b = a.Trim & c1.Trim
                Return b
                eje1 = ""
            End If
            If veclongitud(2) = 3 And veclongitud(4) = 1 Then
                b = a1.Trim & c.Trim
                Return b
                eje1 = ""
            End If
            If veclongitud(2) = 3 And veclongitud(4) = 3 Then
                b = a1.Trim & c1.Trim
                Return b
                eje1 = ""
            End If
        End If
        'VERIFICAR LOS 8 CALCULOS QUE FALTAN

        If vecNumeros(1) = 1 And vecNumeros(4) = 1 And vecNumeros(6) = 1 Then
            If veclongitud(2) = 1 And veclongitud(3) = 3 And veclongitud(5) = 1 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim & vecCampos(2).TrimEnd

                vecCampos(4) = vecCampos(4).PadLeft(3, CChar("0"))
                vecCampos(4) = vecCampos(4).Trim & vecCampos(5).TrimEnd

                vecCampos(6) = vecCampos(6).PadLeft(3, CChar("0"))
                vecCampos(6) = vecCampos(6).Trim
                b = eje1.Trim & vecCampos(1).Trim & "1" & "-" & vecCampos(4).Trim & "0" & "-" & vecCampos(6).Trim
                Return b
                eje1 = ""
            End If

            If veclongitud(2) = 1 And veclongitud(3) = 3 And veclongitud(5) = 3 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim & vecCampos(2).TrimEnd

                vecCampos(4) = vecCampos(4).PadLeft(3, CChar("0"))
                vecCampos(4) = vecCampos(4).Trim

                vecCampos(6) = vecCampos(6).PadLeft(3, CChar("0"))
                vecCampos(6) = vecCampos(6).Trim
                b = eje1.Trim & vecCampos(1).Trim & "1" & "-" & vecCampos(4).Trim & "-" & "1" & "-" & vecCampos(6).Trim
                Return b
                eje1 = ""
            End If

            If veclongitud(2) = 3 And veclongitud(3) = 1 And veclongitud(5) = 1 Then
                vecCampos(1) = vecCampos(1).PadLeft(3, CChar("0"))
                vecCampos(1) = vecCampos(1).Trim

                vecCampos(4) = vecCampos(4).PadLeft(3, CChar("0"))
                vecCampos(4) = vecCampos(3).Trim & vecCampos(4).Trim & vecCampos(5).Trim

                vecCampos(6) = vecCampos(6).PadLeft(3, CChar("0"))
                vecCampos(6) = vecCampos(6).Trim
                b = eje1.Trim & vecCampos(1).Trim & "-" & "1" & vecCampos(4).Trim & "0" & "-" & vecCampos(6).Trim
                Return b
                eje1 = ""
            End If
        End If

        Return "64"
        '===========
    End Function

    Private Function validaTipo2(ByVal Direccion As String) As String
        Dim Contiene As Boolean
        Dim Direc, ViaEsp, CadNueva, Letraeje, DirecNu As String
        Dim UbicaCad, LongNom, Extrae, EjeSecun, i, Numpla, a, LogVec As Integer
        Dim vecVias() As String
        a = 0
        Direc = Direccion
        Letraeje = ""
        Direc.ToUpper()
        Contiene = False
        Dim dsEje As Data.DataSet = SqlHelper.ExecuteDataset(_cadenaConexion, "ConViasEspecial")
        Dim drEje As Data.DataRow
        For Each drEje In dsEje.Tables(0).Rows
            Contiene = Direc.Contains(drEje("NOM_ID").ToString)
            If Contiene = True Then
                UbicaCad = Direc.IndexOf(drEje("NOM_ID").ToString)
                LongNom = drEje("NOM_ID").ToString.Length
                Extrae = UbicaCad + LongNom + 1
                CadNueva = Direc.Substring(Extrae)
                vecVias = CadNueva.Split(CChar(" "))
                ''Divide en un vector
                LogVec = vecVias.Length
                If Not vecVias.Length = 1 Then
                    For i = 0 To vecVias.Length - 1
                        If IsNumeric(vecVias(i)) Then
                            EjeSecun = CInt(vecVias(i))

                            If IsNumeric(vecVias(i + 1)) Then
                                Numpla = CInt(vecVias(i + 1))
                                Letraeje = ""
                            Else
                                If vecVias(i + 1).Length = 1 Then
                                    Letraeje = vecVias(i + 1)
                                    If i + 2 < vecVias.Length AndAlso IsNumeric(vecVias(i + 2)) Then
                                        Numpla = CInt(vecVias(i + 2))
                                    Else
                                        If i + 3 < vecVias.Length AndAlso IsNumeric(vecVias(i + 3)) Then
                                            Numpla = CInt(vecVias(i + 3))
                                        End If
                                    End If
                                Else
                                    Return "Error en placa ViaEspecial"
                                End If
                            End If
                            Exit For
                        End If
                    Next
                Else
                    Return "Error en placa ViaEspecial"
                End If

                ViaEsp = drEje("NOM_BUS").ToString
                Dim dsGen As DataSet = SqlHelper.ExecuteDataset(_cadenaConexion, "ConMallaAveEspe")
                Dim drGen As Data.DataRow
                For Each drGen In dsGen.Tables(0).Rows
                    If drGen.Item(0).ToString = ViaEsp Then
                        If CInt(drGen.Item(5).ToString) = EjeSecun Then
                            a = a + 1
                            If a < 1 Then
                                If drGen.Item(6).ToString = Letraeje.ToString Then
                                    DirecNu = (drGen.Item(1).ToString & " " & drGen.Item(2).ToString & " " & drGen.Item(3).ToString & " " & drGen.Item(4).ToString & " " & drGen.Item(5).ToString & " " & drGen.Item(6).ToString & " " & drGen.Item(7).ToString & " " & Numpla)
                                    Return obtenerCodDirecion(DirecNu)
                                    Exit For
                                End If
                            Else
                                DirecNu = (drGen.Item(1).ToString & " " & drGen.Item(2).ToString & " " & drGen.Item(3).ToString & " " & drGen.Item(4).ToString & " " & drGen.Item(5).ToString & " " & drGen.Item(6).ToString & " " & drGen.Item(7).ToString & " " & Numpla)
                                Return obtenerCodDirecion(DirecNu)
                                Exit For
                            End If
                        End If
                    End If
                Next
                Exit For
            End If

            Contiene = False
        Next
        Return ("ERROR EN DIRECCION")
    End Function

    Public Function telefono(ByVal Numtelefono As Integer) As String
        'Busqueda en el predial
        Dim ds As DataSet = SqlHelper.ExecuteDataset(_cadenaConexion, "SpConsultaTelefono", Numtelefono)

        If ds.Tables(0).Rows.Count > 0 Then
            'si existe retorna los valores del telefono
            Dim coorx As Double
            Dim coory As Double
            Dim Barrio As String = String.Empty
            Dim Localidad As Integer
            Dim Upz As Integer
            Dim codigodireccion As String = String.Empty
            Dim estrato As Integer

            If Not ds.Tables(0).Rows(0)("coorx").GetType.ToString.Equals("System.DBNull") Then
                coorx = CDbl(ds.Tables(0).Rows(0)("coorx").ToString())
            End If
            If Not ds.Tables(0).Rows(0)("coory").GetType.ToString.Equals("System.DBNull") Then
                coory = CDbl(ds.Tables(0).Rows(0)("coory").ToString())
            End If
            If Not ds.Tables(0).Rows(0)("barrio").GetType.ToString.Equals("System.DBNull") Then
                Barrio = ds.Tables(0).Rows(0)("barrio").ToString()
            End If
            If Not ds.Tables(0).Rows(0)("localidad").GetType.ToString.Equals("System.DBNull") Then
                Localidad = CInt(ds.Tables(0).Rows(0)("localidad").ToString())
            End If
            If Not ds.Tables(0).Rows(0)("upz").GetType.ToString.Equals("System.DBNull") Then
                Upz = CInt(ds.Tables(0).Rows(0)("upz").ToString())
            End If
            If Not ds.Tables(0).Rows(0)("codigodir").GetType.ToString.Equals("System.DBNull") Then
                codigodireccion = ds.Tables(0).Rows(0)("codigodir").ToString()
            End If
            If Not ds.Tables(0).Rows(0)("estrato").GetType.ToString.Equals("System.DBNull") Then
                estrato = CInt(ds.Tables(0).Rows(0)("estrato").ToString())
            End If

            Return codigodireccion & ";" & Localidad & ";" & Upz & ";" & Barrio & ";" & coorx & ";" & coory & ";" & estrato & ";10"
        Else
            Return "11 No existe en BD"
        End If

    End Function

End Class
