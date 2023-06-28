Imports System.Math

Public Class coordenadas

    Public _cadenaConexion As String '= ConfigurationManager.ConnectionStrings.Item(1).ConnectionString()

    Public WriteOnly Property cadenaConexion() As String
        Set(ByVal value As String)
            _cadenaConexion = value
        End Set
    End Property

    Public Function coordenadas(ByVal codigo As String) As String
        Dim codigoDir As String
        Dim Vplaca As Integer  'placa del predio
        Dim Residuo As Double 'el residuo de dividir la placa por 2, para identificar si la placa es par o impar
        Dim Tplaca As Integer 'Indica si la placa es par o impar

        If codigo.Length = 17 Then
            codigoDir = codigo.Substring(0, 14)
            Vplaca = CInt(codigo.Substring(14, 3))
            Residuo = Vplaca Mod 2
            Tplaca = 0
            If Residuo = 0 Then
                Tplaca = 1  ' par
            Else
                Tplaca = 0 ' impar
            End If

            '1. Busqueda en el predial  1: Búsqueda sobre predial actual

            Dim ds As DataSet = SqlHelper.ExecuteDataset(_cadenaConexion, "SpConsultarPredio", codigo)

            If ds.Tables(0).Rows.Count > 0 Then
                'si existe retorna los valores del predio
                Dim coorx As Double
                Dim coory As Double
                Dim Barrio As String
                Dim Localidad As Integer
                Dim Upz As Integer
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
                If Not ds.Tables(0).Rows(0)("estrato").GetType.ToString.Equals("System.DBNull") Then
                    estrato = CInt(ds.Tables(0).Rows(0)("estrato").ToString())
                Else
                    estrato = 0
                End If
                Return Localidad & ";" & Upz & ";" & Barrio & ";" & coorx & ";" & coory & ";" & estrato & ";20 Encontrado en predial exacto."

            Else
                '2. Buscar por malla vial exacta
                Dim DR As DataSet
                DR = SqlHelper.ExecuteDataset(_cadenaConexion, "SpConsultarMalla", codigoDir)
                If DR.Tables(0).Rows.Count > 0 Then
                    'llamar la funcion buscar por malla, para retornar las coordenadas
                    Dim cadenaretorno As String
                    Dim CoordX1 As Double
                    Dim CoordX2 As Double
                    Dim CoordY1 As Double
                    Dim CoordY2 As Double
                    Dim barrio_I, barrio_D As String
                    Dim localidad_I, localidad_D As String
                    Dim Estrato_I, Estrato_D As String
                    Dim Upz_I, Upz_D As String
                    Dim cuadrante1 As String  'Dos primeros digitos del codigo de la direccion

                    CoordX1 = CDbl(DR.Tables(0).Rows(0)("CoorX1").ToString())
                    CoordX2 = CDbl(DR.Tables(0).Rows(0)("CoorX2").ToString())
                    CoordY1 = CDbl(DR.Tables(0).Rows(0)("CoorY1").ToString())
                    CoordY2 = CDbl(DR.Tables(0).Rows(0)("CoorY2").ToString())
                    barrio_I = DR.Tables(0).Rows(0)("barrio_izq").ToString()
                    barrio_D = DR.Tables(0).Rows(0)("barrio_der").ToString()
                    localidad_I = DR.Tables(0).Rows(0)("loc_der").ToString()
                    localidad_D = DR.Tables(0).Rows(0)("loc_izq").ToString()
                    Upz_I = DR.Tables(0).Rows(0)("upz_izq").ToString()
                    Upz_D = DR.Tables(0).Rows(0)("upz_der").ToString()
                    Estrato_I = DR.Tables(0).Rows(0)("estrato_izq").ToString()
                    Estrato_D = DR.Tables(0).Rows(0)("estrato_der").ToString()

                    ' extraer los dos primeros digitos del codigo de la direccion
                    cuadrante1 = codigoDir.Substring(0, 2)

                    cadenaretorno = Malla(CoordX1, CoordX2, CoordY1, CoordY2, cuadrante1, Tplaca, Vplaca, 10, barrio_D, barrio_I, localidad_I, localidad_D, Upz_I, Upz_D, Estrato_D, Estrato_I)
                    Return cadenaretorno & ";31 Encontrado en Malla Vial Exacta"
                Else
                    '3. Buscar aproximada por predial en la misma manzana
                    Dim DRPredial As DataSet
                    Dim codigovector As String
                    codigovector = codigoDir & "0"
                    DRPredial = SqlHelper.ExecuteDataset(_cadenaConexion, "SpConsultarVectorPredio", codigovector, Tplaca)
                    If DRPredial.Tables(0).Rows.Count > 0 Then
                        If DRPredial.Tables(0).Rows.Count = 1 Then
                            'aproxima con el mismo predio encontrado
                            Dim coorxmmp As Double
                            Dim coorymmp As Double
                            Dim Barriommp As String
                            Dim Localidadmmp As Integer
                            Dim Upzmmp As Integer
                            Dim estratommp As Integer

                            If Not DRPredial.Tables(0).Rows(0)("coorx").GetType.ToString.Equals("System.DBNull") Then
                                coorxmmp = CDbl(DRPredial.Tables(0).Rows(0)("coorx").ToString())
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("coory").GetType.ToString.Equals("System.DBNull") Then
                                coorymmp = CDbl(DRPredial.Tables(0).Rows(0)("coory").ToString())
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("barrio").GetType.ToString.Equals("System.DBNull") Then
                                Barriommp = DRPredial.Tables(0).Rows(0)("barrio").ToString()
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("localidad").GetType.ToString.Equals("System.DBNull") Then
                                Localidadmmp = CInt(DRPredial.Tables(0).Rows(0)("localidad").ToString())
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("upz").GetType.ToString.Equals("System.DBNull") Then
                                Upzmmp = CInt(DRPredial.Tables(0).Rows(0)("upz").ToString())
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("estrato").GetType.ToString.Equals("System.DBNull") Then
                                estratommp = CInt(DRPredial.Tables(0).Rows(0)("estrato").ToString())
                            Else
                                estratommp = 0
                            End If
                            Return Localidadmmp & ";" & Upzmmp & ";" & Barriommp & ";" & coorxmmp & ";" & coorymmp & ";" & estratommp & ";" & estratommp & ";22 Encontrado por aproximacion predial en la misma manzana, un solo predio encontrado"
                        Else
                            'aproxima por interpolacion, entre el 1 y ultimo predio encontrado en la misma manzana
                            Dim corx1, corx2, cory1, cory2 As Double
                            Dim placa1, placa2 As Integer
                            Dim Barriommi As String
                            Dim Localidadmmi As Integer
                            Dim Upzmmi As Integer
                            Dim estratommi As Integer

                            Dim retornoAproxPredial As String

                            'asignar los valores del primer predio encontrado
                            If Not DRPredial.Tables(0).Rows(0)("coorx").GetType.ToString.Equals("System.DBNull") Then
                                corx1 = CDbl(DRPredial.Tables(0).Rows(0)("coorx").ToString())
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("coory").GetType.ToString.Equals("System.DBNull") Then
                                cory1 = CDbl(DRPredial.Tables(0).Rows(0)("coory").ToString())
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("barrio").GetType.ToString.Equals("System.DBNull") Then
                                Barriommi = DRPredial.Tables(0).Rows(0)("barrio").ToString()
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("localidad").GetType.ToString.Equals("System.DBNull") Then
                                Localidadmmi = CInt(DRPredial.Tables(0).Rows(0)("localidad").ToString())
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("upz").GetType.ToString.Equals("System.DBNull") Then
                                Upzmmi = CInt(DRPredial.Tables(0).Rows(0)("upz").ToString())
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("estrato").GetType.ToString.Equals("System.DBNull") Then
                                estratommi = CInt(DRPredial.Tables(0).Rows(0)("estrato").ToString())
                            Else
                                estratommi = 0
                            End If
                            If Not DRPredial.Tables(0).Rows(0)("codplaca").GetType.ToString.Equals("System.DBNull") Then
                                placa1 = CInt(DRPredial.Tables(0).Rows(0)("codplaca").ToString())
                            End If

                            For Each predio As Data.DataRow In DRPredial.Tables(0).Rows
                                corx2 = CDbl(predio("coorx").ToString)
                                cory2 = CDbl(predio("coory").ToString)
                                placa2 = CInt(predio("codplaca").ToString)
                            Next
                            retornoAproxPredial = AproxPredialInterpola(corx1, corx2, cory1, cory2, placa1, placa2, Vplaca)
                            Return Localidadmmi & ";" & Upzmmi & ";" & Barriommi & ";" & retornoAproxPredial & ";" & estratommi & ";" & estratommi & ";21 Encontrado por aproximacion predial en la misma manzana, interpolado entre dos predios"

                        End If
                    Else
                        '4. buscar aproximada por predial en la manzana del frente
                        Dim DRPredialFrente As DataSet
                        Dim NTPlaca As Integer
                        If Tplaca = 0 Then
                            NTPlaca = 1
                        Else
                            NTPlaca = 0
                        End If
                        DRPredialFrente = SqlHelper.ExecuteDataset(_cadenaConexion, "SpConsultarVectorPredio", codigoDir, NTPlaca)
                        If DRPredialFrente.Tables(0).Rows.Count > 0 Then
                            'aproxima por predial en la manzana del frente
                            If DRPredialFrente.Tables(0).Rows.Count = 1 Then
                                'aproxima sobre el mismo predio encontrado de la manzana del frente
                                Dim coorxmfp As Double
                                Dim coorymfp As Double
                                Dim Barriomfp As String
                                Dim Localidadmfp As Integer
                                Dim Upzmfp As Integer
                                Dim estratomfp As Integer

                                If Not DRPredialFrente.Tables(0).Rows(0)("coorx").GetType.ToString.Equals("System.DBNull") Then
                                    coorxmfp = CDbl(DRPredialFrente.Tables(0).Rows(0)("coorx").ToString())
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("coory").GetType.ToString.Equals("System.DBNull") Then
                                    coorymfp = CDbl(DRPredialFrente.Tables(0).Rows(0)("coory").ToString())
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("barrio").GetType.ToString.Equals("System.DBNull") Then
                                    Barriomfp = DRPredialFrente.Tables(0).Rows(0)("barrio").ToString()
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("localidad").GetType.ToString.Equals("System.DBNull") Then
                                    Localidadmfp = CInt(DRPredialFrente.Tables(0).Rows(0)("localidad").ToString())
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("upz").GetType.ToString.Equals("System.DBNull") Then
                                    Upzmfp = CInt(DRPredialFrente.Tables(0).Rows(0)("upz").ToString())
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("estrato").GetType.ToString.Equals("System.DBNull") Then
                                    estratomfp = CInt(DRPredialFrente.Tables(0).Rows(0)("estrato").ToString())
                                Else
                                    estratomfp = 0
                                End If
                                Return Localidadmfp & ";" & Upzmfp & ";" & Barriomfp & ";" & coorxmfp & ";" & coorymfp & ";" & estratomfp & ";" & estratomfp & ";24 Encontrado por aproximacion predial en la manzana del frente, un solo predio encontrado"
                            Else
                                'aproxima por interpolacion entre el 1 y el ultimo predio encontrado en la manzana del frente
                                Dim corx1f, corx2f, cory1f, cory2f As Double
                                Dim placa1f, placa2f As Integer
                                Dim Barriomfi As String
                                Dim Localidadmfi As Integer
                                Dim Upzmfi As Integer
                                Dim estratomfi As Integer
                                Dim retornoAproxPredialf As String

                                'asignar los valores del primer predio encontrado
                                If Not DRPredialFrente.Tables(0).Rows(0)("coorx").GetType.ToString.Equals("System.DBNull") Then
                                    corx1f = CDbl(DRPredialFrente.Tables(0).Rows(0)("coorx").ToString())
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("coory").GetType.ToString.Equals("System.DBNull") Then
                                    cory1f = CDbl(DRPredialFrente.Tables(0).Rows(0)("coory").ToString())
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("barrio").GetType.ToString.Equals("System.DBNull") Then
                                    Barriomfi = DRPredialFrente.Tables(0).Rows(0)("barrio").ToString()
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("localidad").GetType.ToString.Equals("System.DBNull") Then
                                    Localidadmfi = CInt(DRPredialFrente.Tables(0).Rows(0)("localidad").ToString())
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("upz").GetType.ToString.Equals("System.DBNull") Then
                                    Upzmfi = CInt(DRPredialFrente.Tables(0).Rows(0)("upz").ToString())
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("estrato").GetType.ToString.Equals("System.DBNull") Then
                                    estratomfi = CInt(DRPredialFrente.Tables(0).Rows(0)("estrato").ToString())
                                Else
                                    estratomfi = 0
                                End If
                                If Not DRPredialFrente.Tables(0).Rows(0)("codplaca").GetType.ToString.Equals("System.DBNull") Then
                                    placa1f = CInt(DRPredialFrente.Tables(0).Rows(0)("codplaca").ToString())
                                End If

                                For Each prediof As Data.DataRow In DRPredialFrente.Tables(0).Rows
                                    corx2f = CDbl(prediof("coorx").ToString)
                                    cory2f = CDbl(prediof("coory").ToString)
                                    placa2f = CInt(prediof("codplaca").ToString)
                                Next
                                retornoAproxPredialf = AproxPredialInterpola(corx1f, corx2f, cory1f, cory2f, placa1f, placa2f, Vplaca)
                                Return Localidadmfi & ";" & Upzmfi & ";" & Barriomfi & ";" & retornoAproxPredialf & ";" & estratomfi & ";" & estratomfi & ";23 Encontrado por aproximacion predial en la manzana del frente, interpolado entre dos predios"
                            End If
                        Else
                            '5. buscar por malla vial aproximada
                            'Dim CadenaAproximada As String
                            'CadenaAproximada = AproximaMalla(codigo)
                            'Return CadenaAproximada
                            Return ";;;;;0; siguiente busqueda por malla vial aproximada"
                        End If

                    End If
                End If
            End If
        Else
            'Error en la codificacion de la direcciòn
            Dim ErrorRetorno As String
            Dim DigInicial As String
            DigInicial = (codigo.Substring(0, 2))
            Select Case DigInicial
                Case "61"
                    ErrorRetorno = ";;;;;0;" & codigo & " Error en número de placa principal"
                Case "62"
                    ErrorRetorno = ";;;;;0;" & codigo & " Error en número de placa secundaria"
                Case "63"
                    ErrorRetorno = ";;;;;0;" & codigo & " Error en número de placa"
                Case "64"
                    ErrorRetorno = ";;;;;0;" & codigo & " Error en la estructura de la dirección"
                Case "65"
                    ErrorRetorno = ";;;;;0;" & codigo & " Error por dirección nula"
                Case Else
                    ErrorRetorno = ";;;;;0;" & " Error no identificado"
            End Select
            Return ErrorRetorno
        End If
    End Function


    Public Function AproxPredialInterpola(ByVal X1 As Double, ByVal X2 As Double, ByVal Y1 As Double, ByVal Y2 As Double, ByVal placa1 As Integer, ByVal placa2 As Integer, ByVal placab As Integer) As String
        Dim X, Y As Double
        Dim angulo As Double
        Dim distancia As Double
        Dim diferencia As Double
        Dim distanciatotal As Double
        Dim DpuntoN As Double
        Dim temporal As Double
        Dim temporal2 As Integer
        Dim Xn, Yn As Double
        Dim Hn As Double  'Hipotenusa para el punto n

        'identificar el menor vértice y convertirlo para que quede en el indice 1
        If X2 < X1 Then
            temporal = X1
            X1 = X2
            X2 = temporal
            temporal = Y1
            Y1 = Y2
            Y2 = temporal
            temporal2 = placa1
            placa1 = placa2
            placa2 = temporal2
        End If

        If Y2 > Y1 Then
            X = Abs(X2 - X1)
            Y = Abs(Y2 - Y1)
            distancia = Sqrt(X * X + Y * Y)
            angulo = Asin(Y / distancia)
            diferencia = placa2 - placa1
            distanciatotal = (100 * distancia) / diferencia
            DpuntoN = (distanciatotal * placab) / 100
            If diferencia > 0 Then
                Hn = DpuntoN - (placa1 * distanciatotal / 100)
                Xn = X1 + Hn * Cos(angulo)
                Yn = Y1 + Hn * Sin(angulo)
            Else
                Hn = DpuntoN - (placa2 * distanciatotal / 100)
                Xn = X2 - Hn * Cos(angulo)
                Yn = Y2 - Hn * Sin(angulo)
            End If

        Else
            X = Abs(X2 - X1)
            Y = Abs(Y1 - Y2)
            distancia = Sqrt(X * X + Y * Y)
            angulo = Asin(Y / distancia)
            diferencia = placa2 - placa1
            distanciatotal = (100 * distancia) / diferencia
            DpuntoN = (distanciatotal * placab) / 100
            If diferencia > 0 Then
                Hn = DpuntoN - (placa1 * distanciatotal / 100)
                Xn = X1 + Hn * Cos(angulo)
                Yn = Y1 - Hn * Sin(angulo)
            Else
                Hn = DpuntoN - (placa2 * distanciatotal / 100)
                Xn = X2 - Hn * Cos(angulo)
                Yn = Y2 + Hn * Sin(angulo)
            End If

        End If
        Return Xn & ";" & Yn
    End Function

    

    Public Function Malla(ByVal coorX1 As Double, ByVal coorX2 As Double, ByVal coorY1 As Double, ByVal coorY2 As Double, ByVal cuadrante As String, ByVal Oplaca As Integer, ByVal placa As Integer, ByVal De As Double, ByVal Barrio_Der As String, ByVal Barrio_Izq As String, ByVal Localidad_izq As String, ByVal Localidad_Der As String, ByVal Upz_Izq As String, ByVal Upz_Der As String, ByVal Estrato_Der As String, ByVal Estrato_izq As String) As String

        Dim Adyacente As Double
        Dim Opuesto As Double
        Dim Hipotenusa As Double
        Dim d As Double   'distancia de la placa, en mts.
        Dim b, a As Double
        Dim b1, a1 As Double
        Dim Temporal As Double
        Dim angulo, angulo1 As Double
        Dim X, Y As Double
        Dim Barrio, Localidad, Upz, Estrato As String

        Barrio = ""
        Localidad = ""
        Estrato = ""
        Upz = ""

        'identificar el menor vértice y convertirlo para que quede en el indice 1
        If coorX2 < coorX1 Then
            Temporal = coorX1
            coorX1 = coorX2
            coorX2 = Temporal
            Temporal = coorY1
            coorY1 = coorY2
            coorY2 = Temporal
        End If

        Adyacente = Abs(coorX2 - coorX1)
        Opuesto = Abs(coorY2 - coorY1)

        If cuadrante = "12" Or cuadrante = "13" Or cuadrante = "14" Then
            'Noroccidente -  calle, diagonal y avenida calle
            If Opuesto < 0.0000001 Then
                'es una calle totalmente horizontal
                d = Abs(coorX2 - coorX1) * placa / 100
                If De > 0 Then
                    'se calcula a una distancia de de la via
                    If Oplaca = 1 Then
                        'placa par
                        X = coorX1 + d
                        Y = coorY1 + De
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        X = coorX1 + d
                        Y = coorY1 - De
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                Else
                    'se calcula sobre el eje de la via
                    X = coorX1 + d
                    Y = coorY1
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                End If
            Else
                ' es una calle no vertical
                Hipotenusa = Sqrt(Adyacente * Adyacente + Opuesto * Opuesto)
                angulo = Asin(Opuesto / Hipotenusa)
                d = Hipotenusa * placa / 100
                a = d * Cos(angulo)
                b = d * Sin(angulo)
                If De > 0 Then
                    'se calcula a una distancia de la via
                    angulo1 = 90 - angulo
                    b1 = De * Sin(angulo1)
                    a1 = De * Cos(angulo1)
                    If Oplaca = 1 Then
                        'placa es par
                        If coorY1 > coorY2 Then
                            X = (coorX1 + a) + a1
                            Y = coorY1 - b + b1
                        Else
                            X = coorX1 + a - a1
                            Y = coorY1 + b + b1
                        End If
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Estrato = Estrato_izq
                        Upz = Upz_Izq
                    Else
                        'placa es impar
                        If coorY1 > coorY2 Then
                            X = coorX1 + a - a1
                            Y = coorY1 - b - b1
                        Else
                            X = coorX1 + a + a1
                            Y = coorY1 + b - b1
                        End If
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Estrato = Estrato_izq
                        Upz = Upz_Izq
                    End If
                Else
                    'se calcula las coordenadas sobre la via
                    If coorY1 > coorY2 Then
                        X = coorX1 + a
                        Y = coorY1 - b
                    Else
                        X = coorX1 + a
                        Y = coorY1 + b
                    End If
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                End If

            End If

        ElseIf cuadrante = "15" Or cuadrante = "16" Or cuadrante = "17" Then
            'norte -este - Carrera, transversal, avenida carrera

            If Adyacente < 0.0000001 Then
                'es una carrera totalmente vertical
                d = Abs(coorY2 - coorY1) * placa / 100
                If De > 0 Then
                    'se calcula a una distancia de de la via
                    If Oplaca = 1 Then
                        'placa par
                        X = coorX1 + De
                        Y = coorY1 + d
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        X = coorX1 - De
                        Y = coorY1 + d
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                Else
                    'se calcula sobre el eje de la via
                    X = coorX1
                    Y = coorY1 + d
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                End If
            Else
                ' es una carrera no vertical
                Hipotenusa = Sqrt(Adyacente * Adyacente + Opuesto * Opuesto)
                angulo = Asin(Opuesto / Hipotenusa)
                d = Hipotenusa * placa / 100
                a = d * Cos(angulo)
                b = d * Sin(angulo)
                If De > 0 Then
                    'se calcula a una distancia de la via
                    angulo1 = 90 - angulo
                    b1 = De * Sin(angulo1)
                    a1 = De * Cos(angulo1)
                    If Oplaca = 1 Then
                        'placa es par
                        If coorY1 > coorY2 Then
                            X = coorX2 - a + a1
                            Y = coorY2 + b + b1
                        Else
                            X = coorX2 - a + a1
                            Y = coorY2 - b - b1
                        End If
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        If coorY1 > coorY2 Then
                            X = (coorX2 - a) - a1
                            Y = coorY2 + b - b1
                        Else
                            X = coorX2 - a - a1
                            Y = coorY2 - b + b1
                        End If
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                Else
                    'se calcula las coordenadas sobre la via
                    If coorY1 > coorY2 Then
                        X = coorX2 - a
                        Y = coorY2 + b
                    Else
                        X = coorX1 + a
                        Y = coorY1 + b
                    End If
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                End If

            End If


        ElseIf cuadrante = "32" Or cuadrante = "33" Or cuadrante = "34" Then
            'sur -este - calle, diagonal, avenida calle

            If Opuesto < 0.0000001 Then
                'es una calle totalmente horizontal
                d = Abs(coorX2 - coorX1) * placa / 100
                If De > 0 Then
                    'se calcula a una distancia de de la via
                    If Oplaca = 1 Then
                        'placa par
                        X = coorX1 + d
                        Y = coorY1 + De
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        X = coorX1 + d
                        Y = coorY1 - De
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                Else
                    'se calcula sobre el eje de la via
                    X = coorX1 + d
                    Y = coorY1
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                End If
            Else
                ' es una calle no vertical
                Hipotenusa = Sqrt(Adyacente * Adyacente + Opuesto * Opuesto)
                angulo = Asin(Opuesto / Hipotenusa)
                d = Hipotenusa * placa / 100
                a = d * Cos(angulo)
                b = d * Sin(angulo)
                If De > 0 Then
                    'se calcula a una distancia de la via
                    angulo1 = 90 - angulo
                    b1 = De * Sin(angulo1)
                    a1 = De * Cos(angulo1)
                    If Oplaca = 1 Then
                        'placa es par
                        If coorY1 > coorY2 Then
                            X = (coorX1 + a) + a1
                            Y = coorY1 - b + b1
                        Else
                            X = coorX1 + a - a1
                            Y = coorY1 + b + b1
                        End If
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        If coorY1 > coorY2 Then
                            X = coorX1 + a - a1
                            Y = coorY1 - b - b1
                        Else
                            X = coorX1 + a + a1
                            Y = coorY1 + b - b1
                        End If
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                Else
                    'se calcula las coordenadas sobre la via
                    If coorY1 > coorY2 Then
                        X = coorX1 + a
                        Y = coorY1 - b
                    Else
                        X = coorX1 + a
                        Y = coorY1 + b
                    End If
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                End If

            End If


        ElseIf cuadrante = "35" Or cuadrante = "36" Or cuadrante = "37" Then
            'sur -este - carrera, tranversal, avenida carrera

            If Adyacente < 0.0000001 Then
                'es una carrera totalmente vertical
                d = Abs(coorY2 - coorY1) * placa / 100
                If De > 0 Then
                    'se calcula a una distancia de de la via
                    If Oplaca = 1 Then
                        'placa par
                        X = coorX1 + De
                        Y = coorY2 - d
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        X = coorX1 - De
                        Y = coorY2 - d
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                Else
                    'se calcula sobre el eje de la via
                    X = coorX1
                    Y = coorY2 - d
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                End If
            Else
                ' es una carrera no vertical
                Hipotenusa = Sqrt(Adyacente * Adyacente + Opuesto * Opuesto)
                angulo = Asin(Opuesto / Hipotenusa)
                d = Hipotenusa * placa / 100
                a = d * Cos(angulo)
                b = d * Sin(angulo)
                If De > 0 Then
                    'se calcula a una distancia de la via
                    angulo1 = 90 - angulo
                    b1 = De * Sin(angulo1)
                    a1 = De * Cos(angulo1)
                    If Oplaca = 1 Then
                        'placa es par
                        If coorY1 > coorY2 Then
                            X = coorX1 + a + a1
                            Y = coorY1 - b + b1
                        Else
                            X = coorX2 - a + a1
                            Y = coorY2 - b - b1
                        End If
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        If coorY1 > coorY2 Then
                            X = (coorX1 + a) - a1
                            Y = coorY1 - b - b1
                        Else
                            X = coorX2 - a - a1
                            Y = coorY2 - b + b1
                        End If
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                Else
                    'se calcula las coordenadas sobre la via
                    If coorY1 > coorY2 Then
                        X = coorX1 + a
                        Y = coorY1 - b
                    Else
                        X = coorX2 - a
                        Y = coorY2 - b
                    End If
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                End If

            End If

        ElseIf cuadrante = "02" Or cuadrante = "03" Or cuadrante = "04" Then
            'norte -este - calle, diagonal, avenida calle

            If Opuesto < 0.0000001 Then
                'es una calle totalmente horizontal
                d = Abs(coorX2 - coorX1) * placa / 100
                If De > 0 Then
                    'se calcula a una distancia de de la via
                    If Oplaca = 1 Then
                        'placa par
                        X = coorX2 - d
                        Y = coorY1 + De
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        X = coorX2 - d
                        Y = coorY1 - De
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                Else
                    'se calcula sobre el eje de la via
                    X = coorX2 - d
                    Y = coorY1
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                End If
            Else
                ' es una calle no vertical
                Hipotenusa = Sqrt(Adyacente * Adyacente + Opuesto * Opuesto)
                angulo = Asin(Opuesto / Hipotenusa)
                d = Hipotenusa * placa / 100
                a = d * Cos(angulo)
                b = d * Sin(angulo)
                If De > 0 Then
                    'se calcula a una distancia de la via
                    angulo1 = 90 - angulo
                    b1 = De * Sin(angulo1)
                    a1 = De * Cos(angulo1)
                    If Oplaca = 1 Then
                        'placa es par
                        If coorY1 > coorY2 Then
                            X = (coorX2 - a) + a1
                            Y = coorY2 + b + b1
                        Else
                            X = coorX2 - a - a1
                            Y = coorY2 - b + b1
                        End If
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        If coorY1 > coorY2 Then
                            X = coorX2 - a - a1
                            Y = coorY2 + b - b1
                        Else
                            X = coorX2 - a + a1
                            Y = coorY2 - b - b1
                        End If
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                Else
                    'se calcula las coordenadas sobre la via
                    If coorY1 > coorY2 Then
                        X = coorX2 - a
                        Y = coorY2 + b
                    Else
                        X = coorX2 - a
                        Y = coorY2 - b
                    End If
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                End If

            End If

        ElseIf cuadrante = "05" Or cuadrante = "06" Or cuadrante = "07" Then
            'norte -occidente - carrera, transversal, avenida carrera
            If Adyacente < 0.0000001 Then
                'es una carrera totalmente vertical
                d = Abs(coorY2 - coorY1) * placa / 100
                If De > 0 Then
                    If Oplaca = 1 Then
                        'placa es par
                        X = coorX1 + De
                        Y = coorY1 + d
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        X = coorX1 - De
                        Y = coorY1 + d
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                Else
                    X = coorX1
                    Y = coorY1 + d
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If

                End If
            Else
                ' es una carrera no vertical
                Hipotenusa = Sqrt(Adyacente * Adyacente + Opuesto * Opuesto)
                angulo = Asin(Opuesto / Hipotenusa)
                d = Hipotenusa * placa / 100
                a = d * Cos(angulo)
                b = d * Sin(angulo)
                If De > 0 Then
                    'se calcula a una distancia de la via
                    angulo1 = 90 - angulo
                    b1 = De * Sin(angulo1)
                    a1 = De * Cos(angulo1)
                    If Oplaca = 1 Then
                        'placa es par
                        If coorY1 > coorY2 Then
                            X = coorX2 - a + a1
                            Y = coorY2 + b + b1
                        Else
                            X = coorX1 + a + a1
                            Y = coorY1 + b - b1
                        End If
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else

                        'placa es impar
                        If coorY1 > coorY2 Then
                            X = (coorX2 - a) - a1
                            Y = coorY2 + b - b1
                        Else
                            X = coorX1 + a - a1
                            Y = coorY1 + b + b1
                        End If
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq

                    End If
                Else
                    'se calcula las coordenadas sobre la via
                    If coorY1 > coorY2 Then
                        X = coorX2 - a
                        Y = coorY2 + b
                    Else
                        X = coorX1 + a
                        Y = coorY1 + b
                    End If
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                End If
            End If


        ElseIf cuadrante = "22" Or cuadrante = "23" Or cuadrante = "24" Then
            'sur -occidente - calle, diagonal, avenida calle

            If Opuesto < 0.0000001 Then
                'es una calle totalmente horizontal
                d = Abs(coorX2 - coorX1) * placa / 100
                If De > 0 Then
                    'se calcula a una distancia de de la via
                    If Oplaca = 1 Then
                        'placa par
                        X = coorX2 - d
                        Y = coorY1 + De
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        X = coorX2 - d
                        Y = coorY1 - De
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                Else
                    'se calcula sobre el eje de la via
                    X = coorX2 - d
                    Y = coorY1
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                End If
            Else
                ' es una calle no vertical
                Hipotenusa = Sqrt(Adyacente * Adyacente + Opuesto * Opuesto)
                angulo = Asin(Opuesto / Hipotenusa)
                d = Hipotenusa * placa / 100
                a = d * Cos(angulo)
                b = d * Sin(angulo)
                If De > 0 Then
                    'se calcula a una distancia de la via
                    angulo1 = 90 - angulo
                    b1 = De * Sin(angulo1)
                    a1 = De * Cos(angulo1)
                    If Oplaca = 1 Then
                        'placa es par
                        If coorY1 > coorY2 Then
                            X = (coorX2 - a) + a1
                            Y = coorY2 + b + b1
                        Else
                            X = coorX2 - a - a1
                            Y = coorY2 - b + b1
                        End If
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        If coorY1 > coorY2 Then
                            X = coorX2 - a - a1
                            Y = coorY2 + b - b1
                        Else
                            X = coorX2 - a + a1
                            Y = coorY2 - b - b1
                        End If
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                Else
                    'se calcula las coordenadas sobre la via
                    If coorY1 > coorY2 Then
                        X = coorX2 - a
                        Y = coorY2 + b
                    Else
                        X = coorX2 - a
                        Y = coorY2 - b
                    End If
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    Else
                        'placa es impar
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    End If
                End If

            End If


        ElseIf cuadrante = "25" Or cuadrante = "26" Or cuadrante = "27" Then
            'sur -oeste - carrera, transversal, avenida carrera

            If Adyacente < 0.0000001 Then
                'es una carrera totalmente vertical
                d = Abs(coorY2 - coorY1) * placa / 100
                If De > 0 Then
                    'se calcula a una distancia de de la via
                    If Oplaca = 1 Then
                        'placa par
                        X = coorX1 + De
                        Y = coorY2 - d
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        X = coorX1 - De
                        Y = coorY2 - d
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                Else
                    'se calcula sobre el eje de la via
                    X = coorX1
                    Y = coorY2 - d
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                End If
            Else
                ' es una carrera no vertical
                Hipotenusa = Sqrt(Adyacente * Adyacente + Opuesto * Opuesto)
                angulo = Asin(Opuesto / Hipotenusa)
                d = Hipotenusa * placa / 100
                a = d * Cos(angulo)
                b = d * Sin(angulo)
                If De > 0 Then
                    'se calcula a una distancia de la via
                    angulo1 = 90 - angulo
                    b1 = De * Sin(angulo1)
                    a1 = De * Cos(angulo1)
                    If Oplaca = 1 Then
                        'placa es par
                        If coorY1 > coorY2 Then
                            X = coorX1 + a + a1
                            Y = coorY1 - b + b1
                        Else
                            X = coorX2 - a + a1
                            Y = coorY2 - b - b1
                        End If
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        If coorY1 > coorY2 Then
                            X = (coorX1 + a) - a1
                            Y = coorY1 - b - b1
                        Else
                            X = coorX2 - a - a1
                            Y = coorY2 - b + b1
                        End If
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                Else
                    'se calcula las coordenadas sobre la via
                    If coorY1 > coorY2 Then
                        X = coorX1 + a
                        Y = coorY1 - b
                    Else
                        X = coorX2 - a
                        Y = coorY2 - b
                    End If
                    If Oplaca = 1 Then
                        'placa es par
                        Barrio = Barrio_Izq
                        Localidad = Localidad_izq
                        Upz = Upz_Izq
                        Estrato = Estrato_izq
                    Else
                        'placa es impar
                        Barrio = Barrio_Der
                        Localidad = Localidad_Der
                        Upz = Upz_Der
                        Estrato = Estrato_Der
                    End If
                End If

            End If

        End If

        Return Localidad & ";" & Upz & ";" & Barrio & ";" & X & ";" & Y & ";" & Estrato

    End Function

    Public Function AproximaMalla(ByVal DireccionCod As String) As String

        ' *******  Se debe revisar el envio del numero de iteraciones, cuando se envia por archivo.....****

        Dim RegistroMalla As DataSet
        Dim Iteraciones As Integer  'Numero de iteraciones a realizar en la malla, para aproximar.
        Iteraciones = 3
        Dim CodigoEje As String
        Dim CodigoAproximado As String
        Dim ValPlaca As Integer
        Dim Residuo As Double
        Dim Tplaca As Integer
        Dim encontrado As Boolean
        Dim i As Integer
        CodigoEje = DireccionCod.Substring(0, 14)
        ValPlaca = CInt(DireccionCod.Substring(14, 3))
        Residuo = ValPlaca Mod 2
        Tplaca = 0
        If Residuo = 0 Then
            Tplaca = 1  ' par
        Else
            Tplaca = 0 ' impar
        End If

        encontrado = False
        i = 0
        While encontrado = False And i <= Iteraciones
            i = i + 1

        End While
        'realizar la búsqueda en la BD con el código aproximado
        RegistroMalla = SqlHelper.ExecuteDataset(_cadenaConexion, "SpConsultarMalla", CodigoAproximado)
        If RegistroMalla.Tables(0).Rows.Count > 0 Then
            'llamar la funcion buscar por malla, para retornar las coordenadas
            Dim cadenaretorno As String
            Dim CoordX1 As Double
            Dim CoordX2 As Double
            Dim CoordY1 As Double
            Dim CoordY2 As Double
            Dim barrio_I, barrio_D As String
            Dim localidad_I, localidad_D As String
            Dim Estrato_I, Estrato_D As String
            Dim Upz_I, Upz_D As String
            Dim cuadrante1 As String  'Dos primeros digitos del codigo de la direccion

            CoordX1 = CDbl(RegistroMalla.Tables(0).Rows(0)("X1").ToString())
            CoordX2 = CDbl(RegistroMalla.Tables(0).Rows(0)("X2").ToString())
            CoordY1 = CDbl(RegistroMalla.Tables(0).Rows(0)("Y1").ToString())
            CoordY2 = CDbl(RegistroMalla.Tables(0).Rows(0)("Y2").ToString())
            barrio_I = RegistroMalla.Tables(0).Rows(0)("barrio_der").ToString()
            barrio_D = RegistroMalla.Tables(0).Rows(0)("barrio_izq").ToString()
            localidad_I = RegistroMalla.Tables(0).Rows(0)("loc_der").ToString()
            localidad_D = RegistroMalla.Tables(0).Rows(0)("loc_izq").ToString()
            Upz_I = RegistroMalla.Tables(0).Rows(0)("loc_der").ToString()
            Upz_D = RegistroMalla.Tables(0).Rows(0)("loc_der").ToString()
            Estrato_I = RegistroMalla.Tables(0).Rows(0)("loc_der").ToString()
            Estrato_D = RegistroMalla.Tables(0).Rows(0)("loc_der").ToString()

            ' extraer los dos primeros digitos del codigo de la direccion
            cuadrante1 = CodigoAproximado.Substring(0, 2)

            cadenaretorno = Malla(CoordX1, CoordX2, CoordY1, CoordY2, cuadrante1, Tplaca, ValPlaca, 10, barrio_D, barrio_I, localidad_I, localidad_D, Upz_I, Upz_D, Estrato_D, Estrato_I)
            Return cadenaretorno & ";31 Encontrado en Malla Vial Aproximada en " & i & "iteraciones"
        End If


        If encontrado = False Then

        Else

        End If
    End Function
    
End Class


