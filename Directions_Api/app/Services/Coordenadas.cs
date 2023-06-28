
using Directions_Api.Helpers;
using Microsoft.VisualBasic;
using System.Data;

namespace Directions_Api.Bussines_Logic
{

    public class Coordenadas
    {
        private readonly string cadenaConexion; // = ConfigurationManager.ConnectionStrings.Item(1).ConnectionString()

        public Coordenadas(string cadenaConexion)
        { 
            this.cadenaConexion = cadenaConexion;
        }
        public string coordenadas(string codigo)
        {
            string codigoDir;
            int Vplaca;  // placa del predio
            double Residuo; // el residuo de dividir la placa por 2, para identificar si la placa es par o impar
            int Tplaca; // Indica si la placa es par o impar

            if (codigo.Length == 17)
            {
                codigoDir = codigo.Substring(0, 14);
                Vplaca = System.Convert.ToInt32(codigo.Substring(14, 3));
                Residuo = Vplaca % 2;
                Tplaca = 0;
                if (Residuo == 0)
                    Tplaca = 1;  // par
                else
                    Tplaca = 0;// impar

                // 1. Busqueda en el predial  1: Búsqueda sobre predial actual

                DataSet ds = SqlHelper.ExecuteDataset(cadenaConexion, "SpConsultarPredio", codigo);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    // si existe retorna los valores del predio
                    double coorx = 0.0;
                    double coory = 0.0;
                    string Barrio = null;
                    int Localidad = 0;
                    int Upz = 0;
                    int estrato = 0;

                    if (!ds.Tables[0].Rows[0]["coorx"].GetType().ToString().Equals("System.DBNull"))
                        coorx = System.Convert.ToDouble(ds.Tables[0].Rows[0]["coorx"].ToString());
                    if (!ds.Tables[0].Rows[0]["coory"].GetType().ToString().Equals("System.DBNull"))
                        coory = System.Convert.ToDouble(ds.Tables[0].Rows[0]["coory"].ToString());
                    if (!ds.Tables[0].Rows[0]["barrio"].GetType().ToString().Equals("System.DBNull"))
                        Barrio = ds.Tables[0].Rows[0]["barrio"].ToString();
                    if (!ds.Tables[0].Rows[0]["localidad"].GetType().ToString().Equals("System.DBNull"))
                        Localidad = System.Convert.ToInt32(ds.Tables[0].Rows[0]["localidad"].ToString());
                    if (!ds.Tables[0].Rows[0]["upz"].GetType().ToString().Equals("System.DBNull"))
                        Upz = System.Convert.ToInt32(ds.Tables[0].Rows[0]["upz"].ToString());
                    if (!ds.Tables[0].Rows[0]["estrato"].GetType().ToString().Equals("System.DBNull"))
                        estrato = System.Convert.ToInt32(ds.Tables[0].Rows[0]["estrato"].ToString());
                    else
                        estrato = 0;
                    return Localidad + ";" + Upz + ";" + Barrio + ";" + coorx + ";" + coory + ";" + estrato + ";20 Encontrado en predial exacto.";
                }
                else
                {
                    // 2. Buscar por malla vial exacta
                    DataSet DR;
                    DR = SqlHelper.ExecuteDataset(cadenaConexion, "SpConsultarMalla", codigoDir);
                    if (DR.Tables[0].Rows.Count > 0)
                    {
                        // llamar la funcion buscar por malla, para retornar las coordenadas
                        string cadenaretorno;
                        double CoordX1;
                        double CoordX2;
                        double CoordY1;
                        double CoordY2;
                        string barrio_I, barrio_D;
                        string localidad_I, localidad_D;
                        string Estrato_I, Estrato_D;
                        string Upz_I, Upz_D;
                        string cuadrante1;  // Dos primeros digitos del codigo de la direccion

                        CoordX1 = System.Convert.ToDouble(DR.Tables[0].Rows[0]["CoorX1"].ToString());
                        CoordX2 = System.Convert.ToDouble(DR.Tables[0].Rows[0]["CoorX2"].ToString());
                        CoordY1 = System.Convert.ToDouble(DR.Tables[0].Rows[0]["CoorY1"].ToString());
                        CoordY2 = System.Convert.ToDouble(DR.Tables[0].Rows[0]["CoorY2"].ToString());
                        barrio_I = DR.Tables[0].Rows[0]["barrio_izq"].ToString();
                        barrio_D = DR.Tables[0].Rows[0]["barrio_der"].ToString();
                        localidad_I = DR.Tables[0].Rows[0]["loc_der"].ToString();
                        localidad_D = DR.Tables[0].Rows[0]["loc_izq"].ToString();
                        Upz_I = DR.Tables[0].Rows[0]["upz_izq"].ToString();
                        Upz_D = DR.Tables[0].Rows[0]["upz_der"].ToString();
                        Estrato_I = DR.Tables[0].Rows[0]["estrato_izq"].ToString();
                        Estrato_D = DR.Tables[0].Rows[0]["estrato_der"].ToString();

                        // extraer los dos primeros digitos del codigo de la direccion
                        cuadrante1 = codigoDir.Substring(0, 2);

                        cadenaretorno = Malla(CoordX1, CoordX2, CoordY1, CoordY2, cuadrante1, Tplaca, Vplaca, 10, barrio_D, barrio_I, localidad_I, localidad_D, Upz_I, Upz_D, Estrato_D, Estrato_I);
                        return cadenaretorno + ";31 Encontrado en Malla Vial Exacta";
                    }
                    else
                    {
                        // 3. Buscar aproximada por predial en la misma manzana
                        DataSet DRPredial;
                        string codigovector;
                        codigovector = codigoDir + "0";
                        DRPredial = SqlHelper.ExecuteDataset(cadenaConexion, "SpConsultarVectorPredio", codigovector, Tplaca);
                        if (DRPredial.Tables[0].Rows.Count > 0)
                        {
                            if (DRPredial.Tables[0].Rows.Count == 1)
                            {
                                // aproxima con el mismo predio encontrado
                                double coorxmmp = default;
                                double coorymmp = default;
                                string Barriommp = default;
                                int Localidadmmp = default;
                                int Upzmmp = default; 
                                int estratommp = default;

                                if (!DRPredial.Tables[0].Rows[0]["coorx"].GetType().ToString().Equals("System.DBNull"))
                                    coorxmmp = System.Convert.ToDouble(DRPredial.Tables[0].Rows[0]["coorx"].ToString());
                                if (!DRPredial.Tables[0].Rows[0]["coory"].GetType().ToString().Equals("System.DBNull"))
                                    coorymmp = System.Convert.ToDouble(DRPredial.Tables[0].Rows[0]["coory"].ToString());
                                if (!DRPredial.Tables[0].Rows[0]["barrio"].GetType().ToString().Equals("System.DBNull"))
                                    Barriommp = DRPredial.Tables[0].Rows[0]["barrio"].ToString();
                                if (!DRPredial.Tables[0].Rows[0]["localidad"].GetType().ToString().Equals("System.DBNull"))
                                    Localidadmmp = System.Convert.ToInt32(DRPredial.Tables[0].Rows[0]["localidad"].ToString());
                                if (!DRPredial.Tables[0].Rows[0]["upz"].GetType().ToString().Equals("System.DBNull"))
                                    Upzmmp = System.Convert.ToInt32(DRPredial.Tables[0].Rows[0]["upz"].ToString());
                                if (!DRPredial.Tables[0].Rows[0]["estrato"].GetType().ToString().Equals("System.DBNull"))
                                    estratommp = System.Convert.ToInt32(DRPredial.Tables[0].Rows[0]["estrato"].ToString());
                                else
                                    estratommp = 0;
                                return Localidadmmp + ";" + Upzmmp + ";" + Barriommp + ";" + coorxmmp + ";" + coorymmp + ";" + estratommp + ";" + estratommp + ";22 Encontrado por aproximacion predial en la misma manzana, un solo predio encontrado";
                            }
                            else
                            {
                                // aproxima por interpolacion, entre el 1 y ultimo predio encontrado en la misma manzana
                                double corx1 = default;
                                double corx2 = default; 
                                double cory1 = default;
                                double cory2 = default;
                                int placa1 = default;
                                int placa2 = default;
                                string Barriommi = default;
                                int Localidadmmi = default;
                                int Upzmmi = default;
                                int estratommi = default;

                                string retornoAproxPredial = default;

                                // asignar los valores del primer predio encontrado
                                if (!DRPredial.Tables[0].Rows[0]["coorx"].GetType().ToString().Equals("System.DBNull"))
                                    corx1 = System.Convert.ToDouble(DRPredial.Tables[0].Rows[0]["coorx"].ToString());
                                if (!DRPredial.Tables[0].Rows[0]["coory"].GetType().ToString().Equals("System.DBNull"))
                                    cory1 = System.Convert.ToDouble(DRPredial.Tables[0].Rows[0]["coory"].ToString());
                                if (!DRPredial.Tables[0].Rows[0]["barrio"].GetType().ToString().Equals("System.DBNull"))
                                    Barriommi = DRPredial.Tables[0].Rows[0]["barrio"].ToString();
                                if (!DRPredial.Tables[0].Rows[0]["localidad"].GetType().ToString().Equals("System.DBNull"))
                                    Localidadmmi = System.Convert.ToInt32(DRPredial.Tables[0].Rows[0]["localidad"].ToString());
                                if (!DRPredial.Tables[0].Rows[0]["upz"].GetType().ToString().Equals("System.DBNull"))
                                    Upzmmi = System.Convert.ToInt32(DRPredial.Tables[0].Rows[0]["upz"].ToString());
                                if (!DRPredial.Tables[0].Rows[0]["estrato"].GetType().ToString().Equals("System.DBNull"))
                                    estratommi = System.Convert.ToInt32(DRPredial.Tables[0].Rows[0]["estrato"].ToString());
                                else
                                    estratommi = 0;
                                if (!DRPredial.Tables[0].Rows[0]["codplaca"].GetType().ToString().Equals("System.DBNull"))
                                    placa1 = System.Convert.ToInt32(DRPredial.Tables[0].Rows[0]["codplaca"].ToString());

                                foreach (System.Data.DataRow predio in DRPredial.Tables[0].Rows)
                                {
                                    corx2 = System.Convert.ToDouble(predio["coorx"].ToString());
                                    cory2 = System.Convert.ToDouble(predio["coory"].ToString());
                                    placa2 = System.Convert.ToInt32(predio["codplaca"].ToString());
                                }
                                retornoAproxPredial = AproxPredialInterpola(corx1, corx2, cory1, cory2, placa1, placa2, Vplaca);
                                return Localidadmmi + ";" + Upzmmi + ";" + Barriommi + ";" + retornoAproxPredial + ";" + estratommi + ";" + estratommi + ";21 Encontrado por aproximacion predial en la misma manzana, interpolado entre dos predios";
                            }
                        }
                        else
                        {
                            // 4. buscar aproximada por predial en la manzana del frente
                            DataSet DRPredialFrente;
                            int NTPlaca;
                            if (Tplaca == 0)
                                NTPlaca = 1;
                            else
                                NTPlaca = 0;
                            DRPredialFrente = SqlHelper.ExecuteDataset(cadenaConexion, "SpConsultarVectorPredio", codigoDir, NTPlaca);
                            if (DRPredialFrente.Tables[0].Rows.Count > 0)
                            {
                                // aproxima por predial en la manzana del frente
                                if (DRPredialFrente.Tables[0].Rows.Count == 1)
                                {
                                    // aproxima sobre el mismo predio encontrado de la manzana del frente
                                    double coorxmfp = default;
                                    double coorymfp = default;
                                    string Barriomfp = default;
                                    int Localidadmfp = default;
                                    int Upzmfp = default;
                                    int estratomfp = default;

                                    if (!DRPredialFrente.Tables[0].Rows[0]["coorx"].GetType().ToString().Equals("System.DBNull"))
                                        coorxmfp = System.Convert.ToDouble(DRPredialFrente.Tables[0].Rows[0]["coorx"].ToString());
                                    if (!DRPredialFrente.Tables[0].Rows[0]["coory"].GetType().ToString().Equals("System.DBNull"))
                                        coorymfp = System.Convert.ToDouble(DRPredialFrente.Tables[0].Rows[0]["coory"].ToString());
                                    if (!DRPredialFrente.Tables[0].Rows[0]["barrio"].GetType().ToString().Equals("System.DBNull"))
                                        Barriomfp = DRPredialFrente.Tables[0].Rows[0]["barrio"].ToString();
                                    if (!DRPredialFrente.Tables[0].Rows[0]["localidad"].GetType().ToString().Equals("System.DBNull"))
                                        Localidadmfp = System.Convert.ToInt32(DRPredialFrente.Tables[0].Rows[0]["localidad"].ToString());
                                    if (!DRPredialFrente.Tables[0].Rows[0]["upz"].GetType().ToString().Equals("System.DBNull"))
                                        Upzmfp = System.Convert.ToInt32(DRPredialFrente.Tables[0].Rows[0]["upz"].ToString());
                                    if (!DRPredialFrente.Tables[0].Rows[0]["estrato"].GetType().ToString().Equals("System.DBNull"))
                                        estratomfp = System.Convert.ToInt32(DRPredialFrente.Tables[0].Rows[0]["estrato"].ToString());
                                    else
                                        estratomfp = 0;
                                    return Localidadmfp + ";" + Upzmfp + ";" + Barriomfp + ";" + coorxmfp + ";" + coorymfp + ";" + estratomfp + ";" + estratomfp + ";24 Encontrado por aproximacion predial en la manzana del frente, un solo predio encontrado";
                                }
                                else
                                {
                                    // aproxima por interpolacion entre el 1 y el ultimo predio encontrado en la manzana del frente
                                    double corx1f = default;
                                    double corx2f = default;
                                    double cory1f = default;
                                    double cory2f = default;
                                    int placa1f = default;
                                    int placa2f = default;
                                    string Barriomfi = default;
                                    int Localidadmfi = default;
                                    int Upzmfi = default;
                                    int estratomfi = default;
                                    string retornoAproxPredialf = default;

                                    // asignar los valores del primer predio encontrado
                                    if (!DRPredialFrente.Tables[0].Rows[0]["coorx"].GetType().ToString().Equals("System.DBNull"))
                                        corx1f = System.Convert.ToDouble(DRPredialFrente.Tables[0].Rows[0]["coorx"].ToString());
                                    if (!DRPredialFrente.Tables[0].Rows[0]["coory"].GetType().ToString().Equals("System.DBNull"))
                                        cory1f = System.Convert.ToDouble(DRPredialFrente.Tables[0].Rows[0]["coory"].ToString());
                                    if (!DRPredialFrente.Tables[0].Rows[0]["barrio"].GetType().ToString().Equals("System.DBNull"))
                                        Barriomfi = DRPredialFrente.Tables[0].Rows[0]["barrio"].ToString();
                                    if (!DRPredialFrente.Tables[0].Rows[0]["localidad"].GetType().ToString().Equals("System.DBNull"))
                                        Localidadmfi = System.Convert.ToInt32(DRPredialFrente.Tables[0].Rows[0]["localidad"].ToString());
                                    if (!DRPredialFrente.Tables[0].Rows[0]["upz"].GetType().ToString().Equals("System.DBNull"))
                                        Upzmfi = System.Convert.ToInt32(DRPredialFrente.Tables[0].Rows[0]["upz"].ToString());
                                    if (!DRPredialFrente.Tables[0].Rows[0]["estrato"].GetType().ToString().Equals("System.DBNull"))
                                        estratomfi = System.Convert.ToInt32(DRPredialFrente.Tables[0].Rows[0]["estrato"].ToString());
                                    else
                                        estratomfi = 0;
                                    if (!DRPredialFrente.Tables[0].Rows[0]["codplaca"].GetType().ToString().Equals("System.DBNull"))
                                        placa1f = System.Convert.ToInt32(DRPredialFrente.Tables[0].Rows[0]["codplaca"].ToString());

                                    foreach (System.Data.DataRow prediof in DRPredialFrente.Tables[0].Rows)
                                    {
                                        corx2f = System.Convert.ToDouble(prediof["coorx"].ToString());
                                        cory2f = System.Convert.ToDouble(prediof["coory"].ToString());
                                        placa2f = System.Convert.ToInt32(prediof["codplaca"].ToString());
                                    }
                                    retornoAproxPredialf = AproxPredialInterpola(corx1f, corx2f, cory1f, cory2f, placa1f, placa2f, Vplaca);
                                    return Localidadmfi + ";" + Upzmfi + ";" + Barriomfi + ";" + retornoAproxPredialf + ";" + estratomfi + ";" + estratomfi + ";23 Encontrado por aproximacion predial en la manzana del frente, interpolado entre dos predios";
                                }
                            }
                            else
                                // 5. buscar por malla vial aproximada
                                // Dim CadenaAproximada As String
                                // CadenaAproximada = AproximaMalla(codigo)
                                // Return CadenaAproximada
                                return ";;;;;0; siguiente busqueda por malla vial aproximada";
                        }
                    }
                }
            }
            else
            {
                // Error en la codificacion de la direcciòn
                string ErrorRetorno;
                string DigInicial;
                DigInicial = (codigo.Substring(0, 2));
                switch (DigInicial)
                {
                    case "61":
                        {
                            ErrorRetorno = ";;;;;0;" + codigo + " Error en número de placa principal";
                            break;
                        }

                    case "62":
                        {
                            ErrorRetorno = ";;;;;0;" + codigo + " Error en número de placa secundaria";
                            break;
                        }

                    case "63":
                        {
                            ErrorRetorno = ";;;;;0;" + codigo + " Error en número de placa";
                            break;
                        }

                    case "64":
                        {
                            ErrorRetorno = ";;;;;0;" + codigo + " Error en la estructura de la dirección";
                            break;
                        }

                    case "65":
                        {
                            ErrorRetorno = ";;;;;0;" + codigo + " Error por dirección nula";
                            break;
                        }

                    default:
                        {
                            ErrorRetorno = ";;;;;0;" + " Error no identificado";
                            break;
                        }
                }
                return ErrorRetorno;
            }
        }


        public string AproxPredialInterpola(double X1, double X2, double Y1, double Y2, int placa1, int placa2, int placab)
        {
            double X, Y;
            double angulo;
            double distancia;
            double diferencia;
            double distanciatotal;
            double DpuntoN;
            double temporal;
            int temporal2;
            double Xn, Yn;
            double Hn;  // Hipotenusa para el punto n

            // identificar el menor vértice y convertirlo para que quede en el indice 1
            if (X2 < X1)
            {
                temporal = X1;
                X1 = X2;
                X2 = temporal;
                temporal = Y1;
                Y1 = Y2;
                Y2 = temporal;
                temporal2 = placa1;
                placa1 = placa2;
                placa2 = temporal2;
            }

            if (Y2 > Y1)
            {
                X = Math.Abs(X2 - X1);
                Y = Math.Abs(Y2 - Y1);
                distancia = Math.Sqrt(X * X + Y * Y);
                angulo = Math.Asin(Y / distancia);
                diferencia = placa2 - placa1;
                distanciatotal = (100 * distancia) / diferencia;
                DpuntoN = (distanciatotal * placab) / 100;
                if (diferencia > 0)
                {
                    Hn = DpuntoN - (placa1 * distanciatotal / 100);
                    Xn = X1 + Hn * Math.Cos(angulo);
                    Yn = Y1 + Hn * Math.Sin(angulo);
                }
                else
                {
                    Hn = DpuntoN - (placa2 * distanciatotal / 100);
                    Xn = X2 - Hn * Math.Cos(angulo);
                    Yn = Y2 - Hn * Math.Sin(angulo);
                }
            }
            else
            {
                X = Math.Abs(X2 - X1);
                Y = Math.Abs(Y1 - Y2);
                distancia = Math.Sqrt(X * X + Y * Y);
                angulo = Math.Asin(Y / distancia);
                diferencia = placa2 - placa1;
                distanciatotal = (100 * distancia) / diferencia;
                DpuntoN = (distanciatotal * placab) / 100;
                if (diferencia > 0)
                {
                    Hn = DpuntoN - (placa1 * distanciatotal / 100);
                    Xn = X1 + Hn * Math.Cos(angulo);
                    Yn = Y1 - Hn * Math.Sin(angulo);
                }
                else
                {
                    Hn = DpuntoN - (placa2 * distanciatotal / 100);
                    Xn = X2 - Hn * Math.Cos(angulo);
                    Yn = Y2 + Hn * Math.Sin(angulo);
                }
            }
            return Xn + ";" + Yn;
        }



        public string Malla(double coorX1, double coorX2, double coorY1, double coorY2, string cuadrante, int Oplaca, int placa, double De, string Barrio_Der, string Barrio_Izq, string Localidad_izq, string Localidad_Der, string Upz_Izq, string Upz_Der, string Estrato_Der, string Estrato_izq)
        {
            double Adyacente;
            double Opuesto;
            double Hipotenusa;
            double d;   // distancia de la placa, en mts.
            double b, a;
            double b1, a1;
            double Temporal;
            double angulo, angulo1;
            double X = default;
            double Y = default;
            string Barrio, Localidad, Upz, Estrato;

            Barrio = "";
            Localidad = "";
            Estrato = "";
            Upz = "";

            // identificar el menor vértice y convertirlo para que quede en el indice 1
            if (coorX2 < coorX1)
            {
                Temporal = coorX1;
                coorX1 = coorX2;
                coorX2 = Temporal;
                Temporal = coorY1;
                coorY1 = coorY2;
                coorY2 = Temporal;
            }

            Adyacente = Math.Abs(coorX2 - coorX1);
            Opuesto = Math.Abs(coorY2 - coorY1);

            if (cuadrante == "12" | cuadrante == "13" | cuadrante == "14")
            {
                // Noroccidente -  calle, diagonal y avenida calle
                if (Opuesto < 0.0000001)
                {
                    // es una calle totalmente horizontal
                    d = Math.Abs(coorX2 - coorX1) * placa / (double)100;
                    if (De > 0)
                    {
                        // se calcula a una distancia de de la via
                        if (Oplaca == 1)
                        {
                            // placa par
                            X = coorX1 + d;
                            Y = coorY1 + De;
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            X = coorX1 + d;
                            Y = coorY1 - De;
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                    else
                    {
                        // se calcula sobre el eje de la via
                        X = coorX1 + d;
                        Y = coorY1;
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                }
                else
                {
                    // es una calle no vertical
                    Hipotenusa = Math.Sqrt(Adyacente * Adyacente + Opuesto * Opuesto);
                    angulo = Math.Asin(Opuesto / Hipotenusa);
                    d = Hipotenusa * placa / 100;
                    a = d * Math.Cos(angulo);
                    b = d * Math.Sin(angulo);
                    if (De > 0)
                    {
                        // se calcula a una distancia de la via
                        angulo1 = 90 - angulo;
                        b1 = De * Math.Sin(angulo1);
                        a1 = De * Math.Cos(angulo1);
                        if (Oplaca == 1)
                        {
                            // placa es par
                            if (coorY1 > coorY2)
                            {
                                X = (coorX1 + a) + a1;
                                Y = coorY1 - b + b1;
                            }
                            else
                            {
                                X = coorX1 + a - a1;
                                Y = coorY1 + b + b1;
                            }
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Estrato = Estrato_izq;
                            Upz = Upz_Izq;
                        }
                        else
                        {
                            // placa es impar
                            if (coorY1 > coorY2)
                            {
                                X = coorX1 + a - a1;
                                Y = coorY1 - b - b1;
                            }
                            else
                            {
                                X = coorX1 + a + a1;
                                Y = coorY1 + b - b1;
                            }
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Estrato = Estrato_izq;
                            Upz = Upz_Izq;
                        }
                    }
                    else
                    {
                        // se calcula las coordenadas sobre la via
                        if (coorY1 > coorY2)
                        {
                            X = coorX1 + a;
                            Y = coorY1 - b;
                        }
                        else
                        {
                            X = coorX1 + a;
                            Y = coorY1 + b;
                        }
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                }
            }
            else if (cuadrante == "15" | cuadrante == "16" | cuadrante == "17")
            {
                // norte -este - Carrera, transversal, avenida carrera

                if (Adyacente < 0.0000001)
                {
                    // es una carrera totalmente vertical
                    d = Math.Abs(coorY2 - coorY1) * placa / (double)100;
                    if (De > 0)
                    {
                        // se calcula a una distancia de de la via
                        if (Oplaca == 1)
                        {
                            // placa par
                            X = coorX1 + De;
                            Y = coorY1 + d;
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            X = coorX1 - De;
                            Y = coorY1 + d;
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                    else
                    {
                        // se calcula sobre el eje de la via
                        X = coorX1;
                        Y = coorY1 + d;
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                }
                else
                {
                    // es una carrera no vertical
                    Hipotenusa = Math.Sqrt(Adyacente * Adyacente + Opuesto * Opuesto);
                    angulo = Math.Asin(Opuesto / Hipotenusa);
                    d = Hipotenusa * placa / 100;
                    a = d * Math.Cos(angulo);
                    b = d * Math.Sin(angulo);
                    if (De > 0)
                    {
                        // se calcula a una distancia de la via
                        angulo1 = 90 - angulo;
                        b1 = De * Math.Sin(angulo1);
                        a1 = De * Math.Cos(angulo1);
                        if (Oplaca == 1)
                        {
                            // placa es par
                            if (coorY1 > coorY2)
                            {
                                X = coorX2 - a + a1;
                                Y = coorY2 + b + b1;
                            }
                            else
                            {
                                X = coorX2 - a + a1;
                                Y = coorY2 - b - b1;
                            }
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            if (coorY1 > coorY2)
                            {
                                X = (coorX2 - a) - a1;
                                Y = coorY2 + b - b1;
                            }
                            else
                            {
                                X = coorX2 - a - a1;
                                Y = coorY2 - b + b1;
                            }
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                    else
                    {
                        // se calcula las coordenadas sobre la via
                        if (coorY1 > coorY2)
                        {
                            X = coorX2 - a;
                            Y = coorY2 + b;
                        }
                        else
                        {
                            X = coorX1 + a;
                            Y = coorY1 + b;
                        }
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                }
            }
            else if (cuadrante == "32" | cuadrante == "33" | cuadrante == "34")
            {
                // sur -este - calle, diagonal, avenida calle

                if (Opuesto < 0.0000001)
                {
                    // es una calle totalmente horizontal
                    d = Math.Abs(coorX2 - coorX1) * placa / (double)100;
                    if (De > 0)
                    {
                        // se calcula a una distancia de de la via
                        if (Oplaca == 1)
                        {
                            // placa par
                            X = coorX1 + d;
                            Y = coorY1 + De;
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            X = coorX1 + d;
                            Y = coorY1 - De;
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                    else
                    {
                        // se calcula sobre el eje de la via
                        X = coorX1 + d;
                        Y = coorY1;
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                }
                else
                {
                    // es una calle no vertical
                    Hipotenusa = Math.Sqrt(Adyacente * Adyacente + Opuesto * Opuesto);
                    angulo = Math.Asin(Opuesto / Hipotenusa);
                    d = Hipotenusa * placa / 100;
                    a = d * Math.Cos(angulo);
                    b = d * Math.Sin(angulo);
                    if (De > 0)
                    {
                        // se calcula a una distancia de la via
                        angulo1 = 90 - angulo;
                        b1 = De * Math.Sin(angulo1);
                        a1 = De * Math.Cos(angulo1);
                        if (Oplaca == 1)
                        {
                            // placa es par
                            if (coorY1 > coorY2)
                            {
                                X = (coorX1 + a) + a1;
                                Y = coorY1 - b + b1;
                            }
                            else
                            {
                                X = coorX1 + a - a1;
                                Y = coorY1 + b + b1;
                            }
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            if (coorY1 > coorY2)
                            {
                                X = coorX1 + a - a1;
                                Y = coorY1 - b - b1;
                            }
                            else
                            {
                                X = coorX1 + a + a1;
                                Y = coorY1 + b - b1;
                            }
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                    else
                    {
                        // se calcula las coordenadas sobre la via
                        if (coorY1 > coorY2)
                        {
                            X = coorX1 + a;
                            Y = coorY1 - b;
                        }
                        else
                        {
                            X = coorX1 + a;
                            Y = coorY1 + b;
                        }
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                }
            }
            else if (cuadrante == "35" | cuadrante == "36" | cuadrante == "37")
            {
                // sur -este - carrera, tranversal, avenida carrera

                if (Adyacente < 0.0000001)
                {
                    // es una carrera totalmente vertical
                    d = Math.Abs(coorY2 - coorY1) * placa / (double)100;
                    if (De > 0)
                    {
                        // se calcula a una distancia de de la via
                        if (Oplaca == 1)
                        {
                            // placa par
                            X = coorX1 + De;
                            Y = coorY2 - d;
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            X = coorX1 - De;
                            Y = coorY2 - d;
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                    else
                    {
                        // se calcula sobre el eje de la via
                        X = coorX1;
                        Y = coorY2 - d;
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                }
                else
                {
                    // es una carrera no vertical
                    Hipotenusa = Math.Sqrt(Adyacente * Adyacente + Opuesto * Opuesto);
                    angulo = Math.Asin(Opuesto / Hipotenusa);
                    d = Hipotenusa * placa / 100;
                    a = d * Math.Cos(angulo);
                    b = d * Math.Sin(angulo);
                    if (De > 0)
                    {
                        // se calcula a una distancia de la via
                        angulo1 = 90 - angulo;
                        b1 = De * Math.Sin(angulo1);
                        a1 = De * Math.Cos(angulo1);
                        if (Oplaca == 1)
                        {
                            // placa es par
                            if (coorY1 > coorY2)
                            {
                                X = coorX1 + a + a1;
                                Y = coorY1 - b + b1;
                            }
                            else
                            {
                                X = coorX2 - a + a1;
                                Y = coorY2 - b - b1;
                            }
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            if (coorY1 > coorY2)
                            {
                                X = (coorX1 + a) - a1;
                                Y = coorY1 - b - b1;
                            }
                            else
                            {
                                X = coorX2 - a - a1;
                                Y = coorY2 - b + b1;
                            }
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                    else
                    {
                        // se calcula las coordenadas sobre la via
                        if (coorY1 > coorY2)
                        {
                            X = coorX1 + a;
                            Y = coorY1 - b;
                        }
                        else
                        {
                            X = coorX2 - a;
                            Y = coorY2 - b;
                        }
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                }
            }
            else if (cuadrante == "02" | cuadrante == "03" | cuadrante == "04")
            {
                // norte -este - calle, diagonal, avenida calle

                if (Opuesto < 0.0000001)
                {
                    // es una calle totalmente horizontal
                    d = Math.Abs(coorX2 - coorX1) * placa / (double)100;
                    if (De > 0)
                    {
                        // se calcula a una distancia de de la via
                        if (Oplaca == 1)
                        {
                            // placa par
                            X = coorX2 - d;
                            Y = coorY1 + De;
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            X = coorX2 - d;
                            Y = coorY1 - De;
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                    else
                    {
                        // se calcula sobre el eje de la via
                        X = coorX2 - d;
                        Y = coorY1;
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                }
                else
                {
                    // es una calle no vertical
                    Hipotenusa = Math.Sqrt(Adyacente * Adyacente + Opuesto * Opuesto);
                    angulo = Math.Asin(Opuesto / Hipotenusa);
                    d = Hipotenusa * placa / 100;
                    a = d * Math.Cos(angulo);
                    b = d * Math.Sin(angulo);
                    if (De > 0)
                    {
                        // se calcula a una distancia de la via
                        angulo1 = 90 - angulo;
                        b1 = De * Math.Sin(angulo1);
                        a1 = De * Math.Cos(angulo1);
                        if (Oplaca == 1)
                        {
                            // placa es par
                            if (coorY1 > coorY2)
                            {
                                X = (coorX2 - a) + a1;
                                Y = coorY2 + b + b1;
                            }
                            else
                            {
                                X = coorX2 - a - a1;
                                Y = coorY2 - b + b1;
                            }
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            if (coorY1 > coorY2)
                            {
                                X = coorX2 - a - a1;
                                Y = coorY2 + b - b1;
                            }
                            else
                            {
                                X = coorX2 - a + a1;
                                Y = coorY2 - b - b1;
                            }
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                    else
                    {
                        // se calcula las coordenadas sobre la via
                        if (coorY1 > coorY2)
                        {
                            X = coorX2 - a;
                            Y = coorY2 + b;
                        }
                        else
                        {
                            X = coorX2 - a;
                            Y = coorY2 - b;
                        }
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                }
            }
            else if (cuadrante == "05" | cuadrante == "06" | cuadrante == "07")
            {
                // norte -occidente - carrera, transversal, avenida carrera
                if (Adyacente < 0.0000001)
                {
                    // es una carrera totalmente vertical
                    d = Math.Abs(coorY2 - coorY1) * placa / (double)100;
                    if (De > 0)
                    {
                        if (Oplaca == 1)
                        {
                            // placa es par
                            X = coorX1 + De;
                            Y = coorY1 + d;
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            X = coorX1 - De;
                            Y = coorY1 + d;
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                    else
                    {
                        X = coorX1;
                        Y = coorY1 + d;
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                }
                else
                {
                    // es una carrera no vertical
                    Hipotenusa = Math.Sqrt(Adyacente * Adyacente + Opuesto * Opuesto);
                    angulo = Math.Asin(Opuesto / Hipotenusa);
                    d = Hipotenusa * placa / 100;
                    a = d * Math.Cos(angulo);
                    b = d * Math.Sin(angulo);
                    if (De > 0)
                    {
                        // se calcula a una distancia de la via
                        angulo1 = 90 - angulo;
                        b1 = De * Math.Sin(angulo1);
                        a1 = De * Math.Cos(angulo1);
                        if (Oplaca == 1)
                        {
                            // placa es par
                            if (coorY1 > coorY2)
                            {
                                X = coorX2 - a + a1;
                                Y = coorY2 + b + b1;
                            }
                            else
                            {
                                X = coorX1 + a + a1;
                                Y = coorY1 + b - b1;
                            }
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {

                            // placa es impar
                            if (coorY1 > coorY2)
                            {
                                X = (coorX2 - a) - a1;
                                Y = coorY2 + b - b1;
                            }
                            else
                            {
                                X = coorX1 + a - a1;
                                Y = coorY1 + b + b1;
                            }
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                    else
                    {
                        // se calcula las coordenadas sobre la via
                        if (coorY1 > coorY2)
                        {
                            X = coorX2 - a;
                            Y = coorY2 + b;
                        }
                        else
                        {
                            X = coorX1 + a;
                            Y = coorY1 + b;
                        }
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                }
            }
            else if (cuadrante == "22" | cuadrante == "23" | cuadrante == "24")
            {
                // sur -occidente - calle, diagonal, avenida calle

                if (Opuesto < 0.0000001)
                {
                    // es una calle totalmente horizontal
                    d = Math.Abs(coorX2 - coorX1) * placa / (double)100;
                    if (De > 0)
                    {
                        // se calcula a una distancia de de la via
                        if (Oplaca == 1)
                        {
                            // placa par
                            X = coorX2 - d;
                            Y = coorY1 + De;
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            X = coorX2 - d;
                            Y = coorY1 - De;
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                    else
                    {
                        // se calcula sobre el eje de la via
                        X = coorX2 - d;
                        Y = coorY1;
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                }
                else
                {
                    // es una calle no vertical
                    Hipotenusa = Math.Sqrt(Adyacente * Adyacente + Opuesto * Opuesto);
                    angulo = Math.Asin(Opuesto / Hipotenusa);
                    d = Hipotenusa * placa / 100;
                    a = d * Math.Cos(angulo);
                    b = d * Math.Sin(angulo);
                    if (De > 0)
                    {
                        // se calcula a una distancia de la via
                        angulo1 = 90 - angulo;
                        b1 = De * Math.Sin(angulo1);
                        a1 = De * Math.Cos(angulo1);
                        if (Oplaca == 1)
                        {
                            // placa es par
                            if (coorY1 > coorY2)
                            {
                                X = (coorX2 - a) + a1;
                                Y = coorY2 + b + b1;
                            }
                            else
                            {
                                X = coorX2 - a - a1;
                                Y = coorY2 - b + b1;
                            }
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            if (coorY1 > coorY2)
                            {
                                X = coorX2 - a - a1;
                                Y = coorY2 + b - b1;
                            }
                            else
                            {
                                X = coorX2 - a + a1;
                                Y = coorY2 - b - b1;
                            }
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                    else
                    {
                        // se calcula las coordenadas sobre la via
                        if (coorY1 > coorY2)
                        {
                            X = coorX2 - a;
                            Y = coorY2 + b;
                        }
                        else
                        {
                            X = coorX2 - a;
                            Y = coorY2 - b;
                        }
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                    }
                }
            }
            else if (cuadrante == "25" | cuadrante == "26" | cuadrante == "27")
            {
                // sur -oeste - carrera, transversal, avenida carrera

                if (Adyacente < 0.0000001)
                {
                    // es una carrera totalmente vertical
                    d = Math.Abs(coorY2 - coorY1) * placa / (double)100;
                    if (De > 0)
                    {
                        // se calcula a una distancia de de la via
                        if (Oplaca == 1)
                        {
                            // placa par
                            X = coorX1 + De;
                            Y = coorY2 - d;
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            X = coorX1 - De;
                            Y = coorY2 - d;
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                    else
                    {
                        // se calcula sobre el eje de la via
                        X = coorX1;
                        Y = coorY2 - d;
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                }
                else
                {
                    // es una carrera no vertical
                    Hipotenusa = Math.Sqrt(Adyacente * Adyacente + Opuesto * Opuesto);
                    angulo = Math.Asin(Opuesto / Hipotenusa);
                    d = Hipotenusa * placa / 100;
                    a = d * Math.Cos(angulo);
                    b = d * Math.Sin(angulo);
                    if (De > 0)
                    {
                        // se calcula a una distancia de la via
                        angulo1 = 90 - angulo;
                        b1 = De * Math.Sin(angulo1);
                        a1 = De * Math.Cos(angulo1);
                        if (Oplaca == 1)
                        {
                            // placa es par
                            if (coorY1 > coorY2)
                            {
                                X = coorX1 + a + a1;
                                Y = coorY1 - b + b1;
                            }
                            else
                            {
                                X = coorX2 - a + a1;
                                Y = coorY2 - b - b1;
                            }
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            if (coorY1 > coorY2)
                            {
                                X = (coorX1 + a) - a1;
                                Y = coorY1 - b - b1;
                            }
                            else
                            {
                                X = coorX2 - a - a1;
                                Y = coorY2 - b + b1;
                            }
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                    else
                    {
                        // se calcula las coordenadas sobre la via
                        if (coorY1 > coorY2)
                        {
                            X = coorX1 + a;
                            Y = coorY1 - b;
                        }
                        else
                        {
                            X = coorX2 - a;
                            Y = coorY2 - b;
                        }
                        if (Oplaca == 1)
                        {
                            // placa es par
                            Barrio = Barrio_Izq;
                            Localidad = Localidad_izq;
                            Upz = Upz_Izq;
                            Estrato = Estrato_izq;
                        }
                        else
                        {
                            // placa es impar
                            Barrio = Barrio_Der;
                            Localidad = Localidad_Der;
                            Upz = Upz_Der;
                            Estrato = Estrato_Der;
                        }
                    }
                }
            }

            return Localidad + ";" + Upz + ";" + Barrio + ";" + X + ";" + Y + ";" + Estrato;
        }

        public string AproximaMalla(string DireccionCod)
        {

            // *******  Se debe revisar el envio del numero de iteraciones, cuando se envia por archivo.....****

            DataSet RegistroMalla;
            int Iteraciones;  // Numero de iteraciones a realizar en la malla, para aproximar.
            Iteraciones = 3;
            string CodigoEje;
            string CodigoAproximado = default;
            int ValPlaca;
            double Residuo;
            int Tplaca;
            bool encontrado;
            int i;
            CodigoEje = DireccionCod.Substring(0, 14);
            ValPlaca = System.Convert.ToInt32(DireccionCod.Substring(14, 3));
            Residuo = ValPlaca % 2;
            Tplaca = 0;
            if (Residuo == 0)
                Tplaca = 1;  // par
            else
                Tplaca = 0;// impar

            encontrado = false;
            i = 0;
            while (encontrado == false & i <= Iteraciones)
                i = i + 1;
            // realizar la búsqueda en la BD con el código aproximado
            RegistroMalla = SqlHelper.ExecuteDataset(cadenaConexion, "SpConsultarMalla", CodigoAproximado);
            if (RegistroMalla.Tables[0].Rows.Count > 0)
            {
                // llamar la funcion buscar por malla, para retornar las coordenadas
                string cadenaretorno;
                double CoordX1;
                double CoordX2;
                double CoordY1;
                double CoordY2;
                string barrio_I, barrio_D;
                string localidad_I, localidad_D;
                string Estrato_I, Estrato_D;
                string Upz_I, Upz_D;
                string cuadrante1;  // Dos primeros digitos del codigo de la direccion

                CoordX1 = System.Convert.ToDouble(RegistroMalla.Tables[0].Rows[0]["X1"].ToString());
                CoordX2 = System.Convert.ToDouble(RegistroMalla.Tables[0].Rows[0]["X2"].ToString());
                CoordY1 = System.Convert.ToDouble(RegistroMalla.Tables[0].Rows[0]["Y1"].ToString());
                CoordY2 = System.Convert.ToDouble(RegistroMalla.Tables[0].Rows[0]["Y2"].ToString());
                barrio_I = RegistroMalla.Tables[0].Rows[0]["barrio_der"].ToString();
                barrio_D = RegistroMalla.Tables[0].Rows[0]["barrio_izq"].ToString();
                localidad_I = RegistroMalla.Tables[0].Rows[0]["loc_der"].ToString();
                localidad_D = RegistroMalla.Tables[0].Rows[0]["loc_izq"].ToString();
                Upz_I = RegistroMalla.Tables[0].Rows[0]["loc_der"].ToString();
                Upz_D = RegistroMalla.Tables[0].Rows[0]["loc_der"].ToString();
                Estrato_I = RegistroMalla.Tables[0].Rows[0]["loc_der"].ToString();
                Estrato_D = RegistroMalla.Tables[0].Rows[0]["loc_der"].ToString();

                // extraer los dos primeros digitos del codigo de la direccion
                cuadrante1 = CodigoAproximado.Substring(0, 2);

                cadenaretorno = Malla(CoordX1, CoordX2, CoordY1, CoordY2, cuadrante1, Tplaca, ValPlaca, 10, barrio_D, barrio_I, localidad_I, localidad_D, Upz_I, Upz_D, Estrato_D, Estrato_I);
                return cadenaretorno + ";31 Encontrado en Malla Vial Aproximada en " + i + "iteraciones";
            }


            if (encontrado == false)
            {
            }
            else
            {
            }

            return null;
        }
    }
}
