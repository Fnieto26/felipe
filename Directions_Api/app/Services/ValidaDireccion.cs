using Directions_Api.Helpers;
using Microsoft.VisualBasic;
using System.Data;


namespace Directions_Api.Bussines_Logic
{

    public class ValidaDireccion
    {
        private string cadenaConexion; // = ConfigurationManager.ConnectionStrings.Item("conexion").ConnectionString()

        public ValidaDireccion(string cadenaConexion)
        {
            this.cadenaConexion = cadenaConexion;
        }

        private int _Sur = 0;
        private int _Este = 0;


        public int Sur
        {
            get
            {
                return _Sur;
            }
        }

        public int Este
        {
            get
            {
                return _Este;
            }
        }

        public string obtenerCodDirecion(string Direccion)
        {
            Direccion = limpia(Direccion);
            Direccion = ejes(Direccion);

            return validaEjePrincipal(Direccion);
        }


        /// <summary>
        ///     ''' Funcion que permite limpiar la direccion
        ///     ''' </summary>
        ///     ''' <param name="Direccion"></param>
        ///     ''' <returns></returns>
        ///     ''' <remarks></remarks>
        public string limpia(string Direccion)
        {
            string StrDirec;
            // recibir parametro de la cadena  PARAMETRO DE ENTRADA
            StrDirec = Direccion.ToUpper();
            StrDirec = StrDirec.Replace("-", " ");

            // BUSCA # N_ y varios antes de los remplazos en la tabla
            System.Data.DataSet dsSignos = SqlHelper.ExecuteDataset(cadenaConexion, "conSignos");
            foreach (DataRow drSignos in dsSignos.Tables[0].Rows)
                StrDirec = StrDirec.Replace(drSignos["val_errores"].ToString(), drSignos["cargar"].ToString());

            // Separa numeros de letras
            int i;
            string num = "";
            string letra = "";
            string Caraceje1 = "";
            bool numF = false;
            bool letraF = false;
            bool cambioF = false;

            StrDirec = StrDirec.Replace("  ", " ");
            Direccion = "";

            for (i = 0; i <= StrDirec.Length - 1; i++)
            {
                Caraceje1 = StrDirec.Substring(i, 1);
                if (cambioF)
                {
                    if (numF)
                    {
                        Direccion = Direccion + " " + letra;
                        letra = "";
                    }
                    else
                    {
                        Direccion = Direccion + " " + num;
                        num = "";
                    }
                    cambioF = false;
                }

                if (Information.IsNumeric(Caraceje1))
                {
                    num = num + Caraceje1;
                    if (letraF)
                        cambioF = true;
                    numF = true;
                    letraF = false;
                }
                else
                {
                    letra = letra + Caraceje1;
                    if (numF)
                        cambioF = true;
                    letraF = true;
                    numF = false;
                }
            }
            if (numF)
            {
                Direccion = Direccion + " " + letra + " " + num;
                letra = "";
            }
            else
            {
                Direccion = Direccion + " " + num + " " + letra;
                num = "";
            }
            Direccion = Direccion.TrimStart();

            return Direccion;
        }

        public string ejes(string Direccion)
        {
            int posicionE, posicionS;
            _Sur = 0;
            _Este = 0;
            posicionS = 0;
            Direccion = Direccion.ToUpper();
            posicionS = Direccion.IndexOf("SUR");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace("SUR", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }
            Direccion = Direccion.Replace("  ", " ");
            posicionS = 0;
            posicionS = Direccion.IndexOf("S UR");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace("S UR", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" S U R");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" S U R", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" SU R");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" SU R", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" SU R");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" SU R", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" SUS");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" SUS", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" SUC");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" SUC", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" SIR");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" SIR", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" SU ");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" SU ", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" DUR ");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" DUR ", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" AUR ");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" AUR ", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" SR ");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" SR ", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionS = 0;
            posicionS = Direccion.IndexOf(" S ");
            if (posicionS > 0)
            {
                _Sur = 1;
                Direccion = Direccion.Replace(" S ", " ");
                posicionS = 0;
                Direccion = Direccion.Replace("  ", " ");
            }
            Direccion = Direccion.ToUpper();
            posicionE = Direccion.IndexOf("ESTE");
            if (posicionE > 0)
            {
                _Este = 1;
                Direccion = Direccion.Replace("ESTE", "");
                posicionE = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionE = Direccion.IndexOf(" EST");
            if (posicionE > 0)
            {
                _Este = 1;
                Direccion = Direccion.Replace(" EST", "");
                posicionE = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionE = Direccion.IndexOf(" ETE");
            if (posicionE > 0)
            {
                _Este = 1;
                Direccion = Direccion.Replace(" ETE", "");
                posicionE = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionE = Direccion.IndexOf(" ETE");
            if (posicionE > 0)
            {
                _Este = 1;
                Direccion = Direccion.Replace(" ETE", " ");
                posicionE = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionE = Direccion.IndexOf(" E STE ");
            if (posicionE > 0)
            {
                _Este = 1;
                Direccion = Direccion.Replace(" E STE ", " ");
                posicionE = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionE = Direccion.IndexOf(" ESTS ");
            if (posicionE > 0)
            {
                _Este = 1;
                Direccion = Direccion.Replace(" ESTS", "  ");
                posicionE = 0;
                Direccion = Direccion.Replace("  ", " ");
            }

            posicionE = Direccion.IndexOf(" ES");
            if (posicionE > 0)
            {
                _Este = 1;
                Direccion = Direccion.Replace(" ES ", " ");
                posicionE = 0;
                Direccion = Direccion.Replace("  ", " ");
            }
            // Nuevo copiar para servicio ============
            Direccion = Direccion.Replace("BIS", " BIS ");
            Direccion = Direccion.Replace(" BI ", " BIS ");
            Direccion = Direccion.Replace(" BI S ", " BIS ");
            Direccion = Direccion.Replace(" B I S ", " BIS ");
            Direccion = Direccion.Replace(" B IS ", " BIS ");
            Direccion = Direccion.Replace(" BS ", " BIS ");
            // =====================
            Direccion = Direccion.Replace("  ", " ");
            return Direccion;
        }

        public string validaEjePrincipal(string Direccion)
        {
            // =========== ini
            string StrDir2;
            int i;
            string[] vecCampos;
            vecCampos = Direccion.Split(Convert.ToChar(" "));

            System.Data.DataSet dsEje = SqlHelper.ExecuteDataset(cadenaConexion, "conEje");
            foreach (DataRow drEje in dsEje.Tables[0].Rows)
            {
                if (vecCampos[0].Trim() == drEje["Des_EjePrin"].ToString())
                {
                    vecCampos[0] = drEje["CodEje"].ToString();
                    switch (vecCampos[0])
                    {
                        case "CL":
                        case "KR":
                        case "TV":
                        case "AV":
                        case "DG":
                        case "AC":
                        case "AK":
                            {
                                break;
                            }
                    }
                }
            }
            StrDir2 = "";
            if (vecCampos[0].Equals("CL") | vecCampos[0].Equals("KR") | vecCampos[0].Equals("TV") | vecCampos[0].Equals("DG") | vecCampos[0].Equals("AC") | vecCampos[0].Equals("AK"))
            {
                for (i = 0; i <= vecCampos.Length - 1; i++)
                    StrDir2 = StrDir2 + " " + vecCampos[i] + " ";
                StrDir2 = StrDir2.Replace("  ", " ");
                StrDir2 = StrDir2.Trim();
                StrDir2 = StrDir2.Replace("  ", " ");
                // MessageBox.Show(StrDir2)
                return validaTipo1(StrDir2);
            }
            StrDir2 = "";

            System.Data.DataSet dsEjeAvenida = SqlHelper.ExecuteDataset(cadenaConexion, "conValoresAvenida");
            if (vecCampos[0].Equals("AV"))
            {
                foreach (DataRow drEje in dsEje.Tables[0].Rows)
                {
                    if (vecCampos[1].Trim() == drEje["Des_EjePrin"].ToString())
                        vecCampos[1] = drEje["CodEje"].ToString();
                    switch (vecCampos[1])
                    {
                        case "CL":
                            {
                                // MessageBox.Show("es avenida calle StrDir2")
                                vecCampos[0] = "AC";
                                vecCampos[1] = "";
                                for (i = 0; i <= vecCampos.Length - 1; i++)
                                    StrDir2 = StrDir2 + " " + vecCampos[i] + " ";
                                StrDir2 = StrDir2.Replace("  ", " ");
                                StrDir2 = StrDir2.Trim();
                                StrDir2 = StrDir2.Replace("  ", " ");
                                return validaTipo1(StrDir2);
                            }

                        case "KR":
                            {
                                vecCampos[0] = "AK";
                                vecCampos[1] = "";
                                for (i = 0; i <= vecCampos.Length - 1; i++)
                                    StrDir2 = StrDir2 + " " + vecCampos[i] + " ";
                                StrDir2 = StrDir2.Replace("  ", " ");
                                StrDir2 = StrDir2.Trim();
                                StrDir2 = StrDir2.Replace("  ", " ");
                                return validaTipo1(StrDir2);
                            }

                        case "100":
                            {
                                vecCampos[0] = "AC";
                                vecCampos[1] = "100";
                                for (i = 0; i <= vecCampos.Length - 1; i++)
                                    StrDir2 = StrDir2 + " " + vecCampos[i] + " ";
                                StrDir2 = StrDir2.Replace("  ", " ");
                                StrDir2 = StrDir2.Trim();
                                StrDir2 = StrDir2.Replace("  ", " ");
                                return validaTipo1(StrDir2);
                            }

                        case "68":
                            {
                                // evalua si es sur de la 68 
                                if (_Sur == 1)
                                {
                                    vecCampos[0] = "AK";
                                    vecCampos[1] = "068";
                                    for (i = 0; i <= vecCampos.Length - 1; i++)
                                        StrDir2 = StrDir2 + " " + vecCampos[i] + " ";
                                    StrDir2 = StrDir2.Replace("  ", " ");
                                    StrDir2 = StrDir2.Trim();
                                    StrDir2 = StrDir2.Replace("  ", " ");
                                    return validaTipo1(StrDir2);
                                }
                                else
                                {
                                    return "verifique si es AC o AK";
                                }
                            }

                        case "19":
                            {
                                // evalua si es la 19 
                                if (Information.IsNumeric(vecCampos[2]))
                                {
                                    if (Convert.ToInt32(vecCampos[2]) > 50)
                                    {
                                        vecCampos[0] = "AK";
                                        for (i = 0; i <= vecCampos.Length - 1; i++)
                                            StrDir2 = StrDir2 + " " + vecCampos[i] + " ";
                                        StrDir2 = StrDir2.Replace("  ", " ");
                                        StrDir2 = StrDir2.Trim();
                                        StrDir2 = StrDir2.Replace("  ", " ");
                                        return validaTipo1(StrDir2);
                                    }
                                    else
                                    {
                                        vecCampos[0] = "AC";
                                        for (i = 0; i <= vecCampos.Length - 1; i++)
                                            StrDir2 = StrDir2 + " " + vecCampos[i] + " ";
                                        StrDir2 = StrDir2.Replace("  ", " ");
                                        StrDir2 = StrDir2.Trim();
                                        StrDir2 = StrDir2.Replace("  ", " ");
                                        return validaTipo1(StrDir2);
                                    }
                                }

                                break;
                            }
                    }
                }
            }
            return validaTipo2(Direccion);
        }

        private string validaTipo1(string Direccion)
        {
            int i;
            string[] vecCampos;
            // =======
            string b;
            string c = default;
            string a1 = default;
            string c1 = default;
            string a = default;
            // =======
            int[] vecNumeros = new int[13];
            int[] veclongitud = new int[13];
            //int RecoVecto = -1; // nunca se usa
            //int continum = -1; // nunca se usa
            string CodDir = "";
            //int bandcambio = 1; // nunca se usa
            bool cadnum0, cadnum1, cadnum2, cadnum3, cadnum4;
            bool cadnum5, cadnum6, cadnum7, cadnum8, cadnum9, bis;
            string eje1 = Convert.ToString(0);
            int recalvec = 0;

            vecCampos = Direccion.Split(Convert.ToChar(" "));
            // CARGA EL VECTOR vecnumeros con valores
            for (i = 0; i <= vecCampos.Length - 1; i++)
            {
                cadnum0 = vecCampos[i].Contains(Convert.ToString(0));
                cadnum1 = vecCampos[i].Contains(Convert.ToString(1));
                cadnum2 = vecCampos[i].Contains(Convert.ToString(2));
                cadnum3 = vecCampos[i].Contains(Convert.ToString(3));
                cadnum4 = vecCampos[i].Contains(Convert.ToString(4));
                cadnum5 = vecCampos[i].Contains(Convert.ToString(5));
                cadnum6 = vecCampos[i].Contains(Convert.ToString(6));
                cadnum7 = vecCampos[i].Contains(Convert.ToString(7));
                cadnum8 = vecCampos[i].Contains(Convert.ToString(8));
                cadnum9 = vecCampos[i].Contains(Convert.ToString(9));
                bis = vecCampos[i].Contains("BIS");
                if (cadnum0 == true)
                {
                    vecNumeros[i] = 1;
                    recalvec = i;
                }
                if (cadnum1 == true)
                {
                    vecNumeros[i] = 1;
                    recalvec = i;
                }
                if (cadnum2 == true)
                {
                    vecNumeros[i] = 1;
                    recalvec = i;
                }
                if (cadnum3 == true)
                {
                    vecNumeros[i] = 1;
                    recalvec = i;
                }
                if (cadnum4 == true)
                {
                    vecNumeros[i] = 1;
                    recalvec = i;
                }
                if (cadnum5 == true)
                {
                    vecNumeros[i] = 1;
                    recalvec = i;
                }
                if (cadnum6 == true)
                {
                    vecNumeros[i] = 1;
                    recalvec = i;
                }
                if (cadnum7 == true)
                {
                    vecNumeros[i] = 1;
                    recalvec = i;
                }
                if (cadnum8 == true)
                {
                    vecNumeros[i] = 1;
                    recalvec = i;
                }
                if (cadnum9 == true)
                {
                    vecNumeros[i] = 1;
                    recalvec = i;
                }
                if (bis == true)
                    vecNumeros[i] = 3;
                if (vecNumeros[0] != 0 & i == 0)
                {
                }
            }

            var oldVecCampos = vecCampos;
            vecCampos = new string[recalvec + 1];

            // recalculo el vector numeros
            if (oldVecCampos != null)
                Array.Copy(oldVecCampos, vecCampos, Math.Min(recalvec + 1, oldVecCampos.Length));
            // Genera vector de longitudes
            for (i = 0; i <= recalvec; i++)
                veclongitud[i] = vecCampos[i].Length;

            // TODO: MAnejo de error de rutina con valor cortarlo
            if (vecNumeros[0] == 1)
            {
                CodDir = "61 ERROR DE VIA";
                return "62 ERROR EJE PRINCIPAL";
            }
            // ===========
            if (vecNumeros[1] == 1 & vecNumeros[2] == 1 & vecNumeros[3] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[2]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[3]) > 99)
                    return "63";
            }
            if (vecNumeros[1] == 1 & vecNumeros[2] == 1 & vecNumeros[4] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[2]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[4]) > 99)
                    return "63";
            }
            if (vecNumeros[1] == 1 & vecNumeros[2] == 1 & vecNumeros[5] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[2]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[5]) > 99)
                {
                }
            }
            if (vecNumeros[1] == 1 & vecNumeros[2] == 1 & vecNumeros[6] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[2]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[6]) > 99)
                {
                }
            }

            if (vecNumeros[1] == 1 & vecNumeros[3] == 1 & vecNumeros[4] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[3]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[4]) > 99)
                {
                }
            }

            if (vecNumeros[1] == 1 & vecNumeros[4] == 1 & vecNumeros[5] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[4]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[5]) > 99)
                {
                }
            }

            if (vecNumeros[1] == 1 & vecNumeros[5] == 1 & vecNumeros[6] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[5]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[6]) > 99)
                {
                }
            }
            if (vecNumeros[1] == 1 & vecNumeros[3] == 1 & vecNumeros[5] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[3]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[5]) > 99)
                {
                }
            }
            if (vecNumeros[1] == 1 & vecNumeros[4] == 1 & vecNumeros[6] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[4]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[6]) > 99)
                {
                }
            }
            if (vecNumeros[1] == 1 & vecNumeros[5] == 1 & vecNumeros[7] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[5]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[7]) > 99)
                {
                }
            }
            if (vecNumeros[1] == 1 & vecNumeros[3] == 1 & vecNumeros[6] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[3]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[6]) > 99)
                {
                }
            }
            if (vecNumeros[1] == 1 & vecNumeros[4] == 1 & vecNumeros[7] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[4]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[7]) > 99)
                {
                }
            }
            if (vecNumeros[1] == 1 & vecNumeros[5] == 1 & vecNumeros[8] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[5]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[8]) > 99)
                {
                }
            }

            if (vecNumeros[1] == 1 & vecNumeros[3] == 1 & vecNumeros[7] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[3]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[7]) > 99)
                {
                }
            }
            if (vecNumeros[1] == 1 & vecNumeros[4] == 1 & vecNumeros[8] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[4]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[8]) > 99)
                {
                }
            }
            if (vecNumeros[1] == 1 & vecNumeros[5] == 1 & vecNumeros[9] == 1)
            {
                if (Convert.ToInt32(vecCampos[1]) > 250)
                    return "61";
                if (Convert.ToInt32(vecCampos[5]) > 199)
                    return "62";
                if (Convert.ToInt32(vecCampos[9]) > 99)
                {
                }
            }

            // remplazar cadena 0 por valor de la tabla si lo encuentra trae campo en cl
            if (vecCampos[0] == "CL")
            {
                if (_Sur == 1)
                {
                    if (_Este == 1)
                        eje1 = "32";
                    else
                        eje1 = "22";
                }
                else if (_Este == 1)
                    eje1 = "12";
                else
                    eje1 = "02";
            }
            if (vecCampos[0] == "DG")
            {
                if (_Sur == 1)
                {
                    if (_Este == 1)
                        eje1 = "33";
                    else
                        eje1 = "23";
                }
                else if (_Este == 1)
                    eje1 = "13";
                else
                    eje1 = "03";
            }
            if (vecCampos[0] == "AC")
            {
                if (_Sur == 1)
                {
                    if (_Este == 1)
                        eje1 = "34";
                    else
                        eje1 = "24";
                }
                else if (_Este == 1)
                    eje1 = "14";
                else
                    eje1 = "04";
            }

            if (vecCampos[0] == "KR")
            {
                if (_Sur == 1)
                {
                    if (_Este == 1)
                        eje1 = "35";
                    else
                        eje1 = "25";
                }
                else if (_Este == 1)
                    eje1 = "15";
                else
                    eje1 = "05";
            }

            if (vecCampos[0] == "TV")
            {
                if (_Sur == 1)
                {
                    if (_Este == 1)
                        eje1 = "36";
                    else
                        eje1 = "26";
                }
                else if (_Este == 1)
                    eje1 = "16";
                else
                    eje1 = "06";
            }
            if (vecCampos[0] == "AK")
            {
                if (_Sur == 1)
                {
                    if (_Este == 1)
                        eje1 = "37";
                    else
                        eje1 = "27";
                }
                else if (_Este == 1)
                    eje1 = "17";
                else
                    eje1 = "07";
            }

            if (vecNumeros[1] == 1 & vecNumeros[2] == 1 & vecNumeros[3] == 1)
            {
                vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                vecCampos[1] = vecCampos[1].Trim();

                vecCampos[2] = vecCampos[2].PadLeft(3, Convert.ToChar("0"));
                vecCampos[2] = vecCampos[2].Trim();

                vecCampos[3] = vecCampos[3].PadLeft(3, Convert.ToChar("0"));
                vecCampos[3] = vecCampos[3].Trim();
                b = eje1.Trim() + vecCampos[1].Trim() + "-" + "0" + "-" + vecCampos[2].Trim() + "-" + "0" + "-" + vecCampos[3].Trim();
                return b;
            }

            if (vecNumeros[1] == 1 & vecNumeros[3] == 1 & vecNumeros[4] == 1)
            {
                if (veclongitud[2] == 1)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim() + vecCampos[2].TrimEnd();

                    vecCampos[3] = vecCampos[3].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[3] = vecCampos[3].Trim();

                    vecCampos[4] = vecCampos[4].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[4] = vecCampos[4].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "0" + "-" + vecCampos[3].Trim() + "-" + "0" + "-" + vecCampos[4].Trim();
                    return b;
                }
                // =====
                if (veclongitud[2] == 3)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim();

                    vecCampos[3] = vecCampos[3].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[3] = vecCampos[3].Trim();

                    vecCampos[4] = vecCampos[4].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[4] = vecCampos[4].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "-" + "1" + "-" + vecCampos[3].Trim() + "-" + "0" + "-" + vecCampos[4].Trim();
                    return b;
                }
            }

            if (vecNumeros[1] == 1 & vecNumeros[2] == 1 & vecNumeros[4] == 1)
            {
                if (veclongitud[3] == 1)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim();

                    vecCampos[2] = vecCampos[2].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[2] = vecCampos[2].Trim() + vecCampos[3].TrimEnd();

                    vecCampos[4] = vecCampos[4].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[4] = vecCampos[4].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "-" + "0" + "-" + vecCampos[2].Trim() + "0" + "-" + vecCampos[4].Trim();
                    return b;
                }

                if (veclongitud[3] == 3)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim();

                    vecCampos[2] = vecCampos[2].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[2] = vecCampos[2].Trim();

                    vecCampos[4] = vecCampos[4].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[4] = vecCampos[4].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "-" + "0" + "-" + vecCampos[2].Trim() + "-" + "1" + "-" + vecCampos[4].Trim();
                    return b;
                }
            }

            if (vecNumeros[1] == 1 & vecNumeros[4] == 1 & vecNumeros[5] == 1)
            {
                if (veclongitud[2] == 1 & veclongitud[3] == 3)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim() + vecCampos[2].TrimEnd();

                    vecCampos[4] = vecCampos[4].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[4] = vecCampos[4].Trim();

                    vecCampos[5] = vecCampos[5].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[5] = vecCampos[5].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "1" + "-" + vecCampos[4].Trim() + "-" + "0" + "-" + vecCampos[5].Trim();
                    return b;
                }

                if (veclongitud[2] == 3 & veclongitud[3] == 1)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim();

                    vecCampos[4] = vecCampos[4].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[4] = vecCampos[3].Trim() + vecCampos[4].Trim();

                    vecCampos[5] = vecCampos[5].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[5] = vecCampos[5].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "-" + "1" + vecCampos[4].Trim() + "-" + "0" + "-" + vecCampos[5].Trim();
                    return b;
                }
            }

            if (vecNumeros[1] == 1 & vecNumeros[5] == 1 & vecNumeros[6] == 1)
            {
                if (veclongitud[2] == 1 & veclongitud[3] == 3 & veclongitud[4] == 1)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim() + vecCampos[2].TrimEnd();

                    vecCampos[5] = vecCampos[5].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[5] = vecCampos[4].Trim() + vecCampos[5].TrimEnd();

                    vecCampos[6] = vecCampos[6].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[6] = vecCampos[6].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "1" + vecCampos[5].Trim() + "-" + "0" + "-" + vecCampos[6].Trim();
                    return b;
                }
            }

            if (vecNumeros[1] == 1 & vecNumeros[2] == 1 & vecNumeros[5] == 1)
            {
                if (veclongitud[3] == 1 & veclongitud[4] == 3)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim();

                    vecCampos[2] = vecCampos[2].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[2] = vecCampos[2].Trim() + vecCampos[3].TrimEnd();

                    vecCampos[5] = vecCampos[5].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[5] = vecCampos[5].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "-" + "0" + "-" + vecCampos[2].Trim() + "1" + "-" + vecCampos[5].Trim();
                    return b;
                }

                if (veclongitud[3] == 3 & veclongitud[4] == 1)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim();

                    vecCampos[2] = vecCampos[2].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[2] = vecCampos[2].Trim();

                    vecCampos[5] = vecCampos[5].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[5] = vecCampos[4].Trim() + vecCampos[5].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "-" + "0" + "-" + vecCampos[2].Trim() + "-" + "1" + vecCampos[5].Trim();
                    return b;
                }
            }

            if (vecNumeros[1] == 1 & vecNumeros[2] == 1 & vecNumeros[6] == 1)
            {
                if (veclongitud[3] == 1 & veclongitud[4] == 3 & veclongitud[5] == 1)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim();

                    vecCampos[2] = vecCampos[2].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[2] = vecCampos[2].Trim() + vecCampos[3].TrimEnd();

                    vecCampos[6] = vecCampos[6].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[6] = vecCampos[5].Trim() + vecCampos[6].TrimEnd();
                    b = eje1.Trim() + vecCampos[1].Trim() + "-" + "0" + "-" + vecCampos[2].Trim() + "1" + vecCampos[6].Trim();
                    return b;
                }
            }

            if (vecNumeros[1] == 1 & vecNumeros[3] == 1 & vecNumeros[5] == 1)
            {
                if (veclongitud[2] == 1)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim() + vecCampos[2].TrimEnd();
                    a = eje1.Trim() + vecCampos[1].Trim() + "0" + "-";
                }
                if (veclongitud[2] == 3)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim();
                    a1 = eje1.Trim() + vecCampos[1].Trim() + "-" + "1" + "-";
                }
                if (veclongitud[4] == 1)
                {
                    vecCampos[3] = vecCampos[3].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[3] = vecCampos[3].Trim() + vecCampos[4].TrimEnd();

                    vecCampos[5] = vecCampos[5].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[5] = vecCampos[5].Trim();
                    c = vecCampos[3].Trim() + "0" + "-" + vecCampos[5].Trim();
                }
                if (veclongitud[4] == 3)
                {
                    vecCampos[3] = vecCampos[3].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[3] = vecCampos[3].Trim();

                    vecCampos[5] = vecCampos[5].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[5] = vecCampos[5].Trim();
                    c1 = vecCampos[3].Trim() + "-" + "1" + "-" + vecCampos[5].Trim();
                }
                if (veclongitud[2] == 1 & veclongitud[4] == 1)
                {
                    b = a.Trim() + c.Trim();
                    return b;
                }
                if (veclongitud[2] == 1 & veclongitud[4] == 3)
                {
                    b = a.Trim() + c1.Trim();
                    return b;
                }
                if (veclongitud[2] == 3 & veclongitud[4] == 1)
                {
                    b = a1.Trim() + c.Trim();
                    return b;
                }
                if (veclongitud[2] == 3 & veclongitud[4] == 3)
                {
                    b = a1.Trim() + c1.Trim();
                    return b;
                }
            }
            // VERIFICAR LOS 8 CALCULOS QUE FALTAN

            if (vecNumeros[1] == 1 & vecNumeros[4] == 1 & vecNumeros[6] == 1)
            {
                if (veclongitud[2] == 1 & veclongitud[3] == 3 & veclongitud[5] == 1)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim() + vecCampos[2].TrimEnd();

                    vecCampos[4] = vecCampos[4].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[4] = vecCampos[4].Trim() + vecCampos[5].TrimEnd();

                    vecCampos[6] = vecCampos[6].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[6] = vecCampos[6].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "1" + "-" + vecCampos[4].Trim() + "0" + "-" + vecCampos[6].Trim();
                    return b;
                }

                if (veclongitud[2] == 1 & veclongitud[3] == 3 & veclongitud[5] == 3)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim() + vecCampos[2].TrimEnd();

                    vecCampos[4] = vecCampos[4].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[4] = vecCampos[4].Trim();

                    vecCampos[6] = vecCampos[6].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[6] = vecCampos[6].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "1" + "-" + vecCampos[4].Trim() + "-" + "1" + "-" + vecCampos[6].Trim();
                    return b;
                }

                if (veclongitud[2] == 3 & veclongitud[3] == 1 & veclongitud[5] == 1)
                {
                    vecCampos[1] = vecCampos[1].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[1] = vecCampos[1].Trim();

                    vecCampos[4] = vecCampos[4].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[4] = vecCampos[3].Trim() + vecCampos[4].Trim() + vecCampos[5].Trim();

                    vecCampos[6] = vecCampos[6].PadLeft(3, Convert.ToChar("0"));
                    vecCampos[6] = vecCampos[6].Trim();
                    b = eje1.Trim() + vecCampos[1].Trim() + "-" + "1" + vecCampos[4].Trim() + "0" + "-" + vecCampos[6].Trim();
                    return b;
                }
            }

            return "64";
        }

        private string validaTipo2(string Direccion)
        {
            bool Contiene;
            string Direc, ViaEsp, CadNueva, Letraeje, DirecNu;
            int UbicaCad, LongNom, Extrae;
            int EjeSecun = default;
            int i;
            int Numpla = default;
            int a, LogVec;
            string[] vecVias;
            a = 0;
            Direc = Direccion;
            Letraeje = "";
            Direc.ToUpper();
            Contiene = false;
            System.Data.DataSet dsEje = SqlHelper.ExecuteDataset(cadenaConexion, "ConViasEspecial");
            foreach (DataRow drEje in dsEje.Tables[0].Rows)
            {
                Contiene = Direc.Contains(drEje["NOM_ID"].ToString());
                if (Contiene == true)
                {
                    UbicaCad = Direc.IndexOf(drEje["NOM_ID"].ToString());
                    LongNom = drEje["NOM_ID"].ToString().Length;
                    Extrae = UbicaCad + LongNom + 1;
                    CadNueva = Direc.Substring(Extrae);
                    vecVias = CadNueva.Split(Convert.ToChar(" "));
                    // 'Divide en un vector
                    LogVec = vecVias.Length;
                    if (vecVias.Length != 1)
                    {
                        for (i = 0; i <= vecVias.Length - 1; i++)
                        {
                            if (Information.IsNumeric(vecVias[i]))
                            {
                                EjeSecun = Convert.ToInt32(vecVias[i]);

                                if (Information.IsNumeric(vecVias[i + 1]))
                                {
                                    Numpla = Convert.ToInt32(vecVias[i + 1]);
                                    Letraeje = "";
                                }
                                else if (vecVias[i + 1].Length == 1)
                                {
                                    Letraeje = vecVias[i + 1];
                                    if (i + 2 < vecVias.Length && Information.IsNumeric(vecVias[i + 2]))
                                        Numpla = Convert.ToInt32(vecVias[i + 2]);
                                    else if (i + 3 < vecVias.Length && Information.IsNumeric(vecVias[i + 3]))
                                        Numpla = Convert.ToInt32(vecVias[i + 3]);
                                }
                                else
                                    return "Error en placa ViaEspecial";
                                break;
                            }
                        }
                    }
                    else
                        return "Error en placa ViaEspecial";

                    ViaEsp = drEje["NOM_BUS"].ToString();
                    DataSet dsGen = SqlHelper.ExecuteDataset(cadenaConexion, "ConMallaAveEspe");
                    foreach (DataRow drGen in dsGen.Tables[0].Rows)
                    {
                        if (drGen[0].ToString() == ViaEsp)
                        {
                            if (Convert.ToInt32(drGen[5].ToString()) == EjeSecun)
                            {
                                a = a + 1;
                                if (a < 1)
                                {
                                    if (drGen[6].ToString() == Letraeje.ToString())
                                    {
                                        DirecNu = (drGen[1].ToString() + " " + drGen[2].ToString() + " " + drGen[3].ToString() + " " + drGen[4].ToString() + " " + drGen[5].ToString() + " " + drGen[6].ToString() + " " + drGen[7].ToString() + " " + Numpla);
                                        return obtenerCodDirecion(DirecNu);
                                    }
                                }
                                else
                                {
                                    DirecNu = (drGen[1].ToString() + " " + drGen[2].ToString() + " " + drGen[3].ToString() + " " + drGen[4].ToString() + " " + drGen[5].ToString() + " " + drGen[6].ToString() + " " + drGen[7].ToString() + " " + Numpla);
                                    return obtenerCodDirecion(DirecNu);
                                }
                            }
                        }
                    }
                    break;
                }

                Contiene = false;
            }
            return ("ERROR EN DIRECCION");
        }

        public string telefono(string Numtelefono)
        {
            // Busqueda en el predial
            DataSet ds = SqlHelper.ExecuteDataset(cadenaConexion, "SpConsultaTelefono", Numtelefono);

            if (ds.Tables[0].Rows.Count > 0)
            {
                // si existe retorna los valores del telefono
                double coorx = default;
                double coory = default;
                string Barrio = string.Empty;
                int Localidad = default;
                int Upz = default;
                string codigodireccion = string.Empty;
                int estrato = default;

                if (!ds.Tables[0].Rows[0]["coorx"].GetType().ToString().Equals("System.DBNull"))
                    coorx = Convert.ToDouble(ds.Tables[0].Rows[0]["coorx"].ToString());
                if (!ds.Tables[0].Rows[0]["coory"].GetType().ToString().Equals("System.DBNull"))
                    coory = Convert.ToDouble(ds.Tables[0].Rows[0]["coory"].ToString());
                if (!ds.Tables[0].Rows[0]["barrio"].GetType().ToString().Equals("System.DBNull"))
                    Barrio = ds.Tables[0].Rows[0]["barrio"].ToString();
                if (!ds.Tables[0].Rows[0]["localidad"].GetType().ToString().Equals("System.DBNull"))
                    Localidad = Convert.ToInt32(ds.Tables[0].Rows[0]["localidad"].ToString());
                if (!ds.Tables[0].Rows[0]["upz"].GetType().ToString().Equals("System.DBNull"))
                    Upz = Convert.ToInt32(ds.Tables[0].Rows[0]["upz"].ToString());
                if (!ds.Tables[0].Rows[0]["codigodir"].GetType().ToString().Equals("System.DBNull"))
                    codigodireccion = ds.Tables[0].Rows[0]["codigodir"].ToString();
                if (!ds.Tables[0].Rows[0]["estrato"].GetType().ToString().Equals("System.DBNull"))
                    estrato = Convert.ToInt32(ds.Tables[0].Rows[0]["estrato"].ToString());

                return codigodireccion + ";" + Localidad + ";" + Upz + ";" + Barrio + ";" + coorx + ";" + coory + ";" + estrato + ";10";
            }
            else
                return "11 No existe en BD";
        }
    }

}
