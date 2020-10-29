using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.IO;
using System.Net;
using ShamanClases.PanelC;
using System.Reflection;
using System.Configuration;
using ShamanClases;
using zkemkeeper;
using System.Net.Mail;
using System.Data.SqlClient;
using ShamanExpressDLL;
using modFechas = ShamanClases.modFechas;
using modDeclares = ShamanClases.modDeclares;

namespace FichadaRelojUyService
{
    public partial class Service1 : ServiceBase
    {

        Timer t = new Timer();
        private Conexion shamanConexion = new Conexion();
        string m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        string dBServer1 = ConfigurationManager.AppSettings["DBServer1"];

        string origenFichada = ConfigurationManager.AppSettings["origenFichada"];

        string timePool = ConfigurationManager.AppSettings["TimePool"];
        int saveInDB = Convert.ToInt32(ConfigurationManager.AppSettings["SaveInDB"]);
        //bool setByDoc = ConfigurationManager.AppSettings["SetByDoc"] == "1";

        List<Relojes> relojes;

        DateTime fecLastCreateGrilla = DateTime.Now.AddDays(-1);

        public Service1()
        {
            InitializeComponent();
        }

        private string GetValueOf(string relojStringProperty, string propertyName)
        {
            propertyName += "=";

            if (relojStringProperty.Contains(propertyName))
                return relojStringProperty.Replace(propertyName, "");
            return "";
        }

        protected override void OnStart(string[] args)
        {
            new System.Threading.Thread(StartService).Start();
        }
        internal void StartService()
        {
            t.Elapsed += delegate { ElapsedHandler(); };
            t.Interval = Convert.ToInt32(timePool) * 60000;
            t.Start();
            Logger.GetInstance().AddLog(true, "OnStart", "Servicio inicializado.");
            CustomMail mail = new CustomMail(MailType.Information, "Se ha inicializado el servicio de Fichada", "Servicio de Fichada");

            CreateGrilla();

            mail.Send();
            Console.WriteLine("Starting service");
        }

        private void CreateGrilla()
        {

            if (fecLastCreateGrilla.Day != DateTime.Now.Day)
            {
                Logger.GetInstance().AddLog(true, "CreateGrilla", "Corresponde correr CreateGrilla");
                if (ConnectDLL())
                {
                    conGrillaOperativa objConGrillaOperativa = new conGrillaOperativa();
                    objConGrillaOperativa.CreateGrillaView(DateTime.Now, true);
                    Logger.GetInstance().AddLog(true, "CreateGrilla", $"CreateGrillaView({DateTime.Now}, true)");
                    fecLastCreateGrilla = DateTime.Now;
                    modDatabase.cnnsNET.Remove(ShamanExpressDLL.modDeclares.cnnDefault);
                }
            }
        }

        private bool ConnectDLL()
        {
            StartUp init = new StartUp();

            if (init.GetValoresHardkey(false))
            {

                if (init.GetVariablesConexion())
                {

                    if (init.AbrirConexion(ShamanExpressDLL.modDeclares.cnnDefault))
                    {
                        ShamanExpressDLL.modFechas.InitDateVars();
                        Logger.GetInstance().AddLog(true, "setConexionDB", string.Format("Conectado a Database Shaman {0}", ShamanExpressDLL.modDatabase.cnnCatalog));
                        return true;
                    }
                    else
                    {
                        Logger.GetInstance().AddLog(false, "setConexionDB", "No se pudo conectar a base de datos Shaman - " + init.MyLastExec.ErrorDescription);
                    }
                }
                else
                {
                    Logger.GetInstance().AddLog(false, "setConexionDB", "No se pudieron recuperar las variables de conexión - " + init.MyLastExec.ErrorDescription);
                }
            }
            else
            {
                Logger.GetInstance().AddLog(false, "setConexionDB", "No se encuentran los valores HKey - " + init.MyLastExec.ErrorDescription);
            }

            return false;

        }

        protected override void OnPause()
        {
            Logger.GetInstance().AddLog(true, "OnPause", "Se ejecutó el método OnPause, el servicio deja de estar activo.");
            CustomMail mail = new CustomMail(MailType.Information, "Se ha pausado el servicio de Fichada", "Servicio de Fichada");
            mail.Send();
            t.Stop();
        }

        protected override void OnContinue()
        {
            Logger.GetInstance().AddLog(true, "OnPause", "Se ejecutó el método OnContinue, el servicio vuelve a estar activo.");
            t.Start();
        }

        protected override void OnStop()
        {
            Logger.GetInstance().AddLog(true, "OnStop", "Se ejecutó el método OnStop, el servicio deja de estar activo.");
            CustomMail mail = new CustomMail(MailType.Information, "Se ha detenido el servicio de Fichada", "Servicio de Fichada");
            mail.Send();
            t.Stop();
        }

        public void ElapsedHandler()
        {
            CreateGrilla();

            if (relojes == null || relojes.Count == 0)
            {
                if (origenFichada.ToLower() == "txt")
                {
                    ProcesarTXT();
                }
                else
                {
                    if (LeerRelojes())
                        ProcesarRelojes();
                }
            }
            else
                ProcesarRelojes();
        }

        private bool LeerRelojes()
        {
            try
            {
                //List<Relojes> relojes = ConfigurationManager.GetSection("relojes") as List<Relojes>;
                var list = ConfigurationManager.AppSettings.AllKeys.Where(k => k.StartsWith("data.list.")).Select(k => ConfigurationManager.AppSettings[k]).ToList();
                relojes = new List<Relojes>();
                foreach (var item in list)
                {
                    string[] relojString = item.Split(';');
                    Relojes reloj = new Relojes();
                    reloj.Id = Convert.ToInt32(GetValueOf(relojString[0], "Id"));
                    reloj.Descripcion = GetValueOf(relojString[1], "Descripcion");
                    reloj.DireccionIP = GetValueOf(relojString[2], "DireccionIP");
                    reloj.Puerto = Convert.ToInt32(GetValueOf(relojString[3], "Puerto"));
                    reloj.Vaciar = Convert.ToInt32(GetValueOf(relojString[4], "Vaciar"));
                    if (relojString.Count() > 5)
                        reloj.CommPassword = Convert.ToInt32(GetValueOf(relojString[5], "CommPassword"));

                    relojes.Add(reloj);
                }
                return relojes.Count > 0;
            }
            catch (Exception ex)
            {
                Logger.GetInstance().AddLog(false, "LeerRelojes()", ex.Message);
            }
            return false;
        }

        private void ProcesarRelojes()
        {
            try
            {
                if (relojes != null && relojes.Count > 0)
                {
                    Logger.GetInstance().AddLog(true, "ProcesarRelojes()", "Procesando " + relojes.Count + " relojes");

                    foreach (var item in relojes)
                        this.clkZKSoft(item.Id, 1, item.Descripcion, item.DireccionIP, item.Puerto, 0, Convert.ToBoolean(item.Vaciar), item.CommPassword);
                }
                else
                    Logger.GetInstance().AddLog(true, "ProcesarRelojes()", "No se encontraron relojes, reinicie el servicio");

                if (modCache.cnnsCache.Count > 0)
                    modDeclares.ShamanSession.Cerrar(modDeclares.ShamanSession.PID);
            }
            catch (Exception ex)
            {
                Logger.GetInstance().AddLog(false, "ProcesarRelojes()", ex.Message);
            }
        }

        private void ProcesarTXT()
        {
            StreamReader fileReader = null;
            try
            {
                Logger.GetInstance().AddLog(true, "ProcesarTXT()", "incio");

                string pathSource = ConfigurationManager.AppSettings["dataTxt"].Split(';')[0];
                string pathDest = Path.Combine(m_exePath, "marcas_" + modFechas.DateToSql(DateTime.Now).Replace("-", "_") + " " + DateTime.Now.Hour + "." + DateTime.Now.Minute + ".log");
                short eliminar = Convert.ToInt16(ConfigurationManager.AppSettings["dataTxt"].Split(';')[1]);

                File.Move(pathSource, pathDest);

                fileReader = new StreamReader(pathDest);
                List<RelojResponse> relojResponse = new List<RelojResponse>();

                do
                {
                    string vLin = fileReader.ReadLine();

                    if (vLin.Length > 48)
                    {
                        //Logger.GetInstance().AddLog(true, "ProcesarTXT()", "preparandose para leer valores");
                        long numero_empleado;
                        long.TryParse(vLin.Substring(0, 10), out numero_empleado);
                        DateTime fecha = modFechas.NtoD(Convert.ToInt32(vLin.Substring(11, 10).Replace("-", "")));
                        string fechaHora = vLin.Substring(11, 19);
                        string tipo_marca = vLin.Substring(31, 2);
                        string numero_reloj = vLin.Substring(34, 3);
                        string numero_movil = vLin.Substring(42, 8);

                        //Logger.GetInstance().AddLog(true, "ProcesarTXT()", "todos los valores leidos");

                        if (fecha.Year == DateTime.Now.Year)
                        {
                            //Logger.GetInstance().AddLog(true, "ProcesarTXT()", "preparandose para encolar valores");
                            RelojResponse relojResponseItem = new RelojResponse();
                            relojResponseItem.Fich = fechaHora;
                            relojResponseItem.Nro = Convert.ToInt32(numero_reloj);
                            relojResponseItem.SdwEnrollNumber = numero_empleado.ToString();
                            //relojResponseItem.IdwVerifyMode = idwVerifyMode; //TODO: que es?
                            //01 entrada,  02 salida-int, 03 entrada_int, 04 salida
                            int nTipoMarca = Convert.ToInt32(tipo_marca);
                            //switch (nTipoMarca)
                            //{
                            //    case 1:
                            //    case 3:
                            //        relojResponseItem.IdwInOutMode = 0;
                            //        break;
                            //    case 2:
                            //    case 4:
                            //        relojResponseItem.IdwInOutMode = 1;
                            //        break;
                            //    default:
                            //        Logger.GetInstance().AddLog(false, "ProcesarTXT()", string.Format("El valor {0} para tipo de marca es incorrecto, se asigna el valor 0", tipo_marca));
                            //        relojResponseItem.IdwInOutMode = 0;
                            //        break;
                            //}

                            if (nTipoMarca == 1 || nTipoMarca == 4)
                            {
                                relojResponseItem.IdwInOutMode = nTipoMarca == 1 ? 0 : 1;
                                relojResponseItem.IdwWorkcode = Convert.ToInt32(numero_movil);
                                relojResponse.Add(relojResponseItem);
                                //Logger.GetInstance().AddLog(true, "ProcesarTXT()", string.Format("Fichada encolada: Fecha: {0}  Legajo: {1}", relojResponseItem.Fich, relojResponseItem.SdwEnrollNumber));
                            }

                        }

                    }

                } while (!fileReader.EndOfStream);

                Logger.GetInstance().AddLog(true, "ProcesarTXT()", "listo para enviar lista a guardar.");

                SaveInDataBase(null, relojResponse);

                if (eliminar == 1)
                {
                    fileReader.Close();
                    fileReader = null;
                    Logger.GetInstance().AddLog(true, "ProcesarTXT()", "eliminando archivo local. " + pathDest);
                    File.Delete(pathDest);
                }

            }
            catch (Exception ex)
            {
                Logger.GetInstance().AddLog(false, "ProcesarTXT()", ex.Message);
            }
            finally
            {
                // Close streams
                if (fileReader != null)
                    fileReader.Close();
            }
        }

        private bool clkZKSoft(int pRid, int pNro, string pDes, string pDir, int pPor, long pPsw, bool vClean = false, int pCommPassword = 0)
        {
            bool clkZKSoft = false;

            try
            {
                string sdwEnrollNumber = "";
                int idwVerifyMode;
                int idwInOutMode;
                int idwYear;
                int idwMonth;
                int idwDay;
                int idwHour;
                int idwMinute;
                int idwSecond;
                int idwWorkcode = 0;
                string vFic;

                CZKEM Reloj = new CZKEM();
                DevOps devolucionOperacion = new DevOps();

                RelojesIngresos objFichada = new RelojesIngresos();
                List<RelojResponse> relojResponse = new List<RelojResponse>();

                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Conectandose " + pDir + ":" + pPor);

                if (pCommPassword > 0)
                    Reloj.SetCommPassword(pCommPassword);

                if (Reloj.Connect_Net(pDir, pPor))
                {
                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Conectado a " + pDir + ":" + pPor);
                    Reloj.EnableDevice(pNro, false);

                    // ----> Leo Datos
                    if (Reloj.ReadGeneralLogData(pNro))
                    {
                        // SSR_GetGeneralLogData
                        // ----> Leo Datos
                        while (Reloj.SSR_GetGeneralLogData(pNro, out sdwEnrollNumber, out idwVerifyMode, out idwInOutMode, out idwYear, out idwMonth, out idwDay, out idwHour, out idwMinute, out idwSecond, ref idwWorkcode))
                        {
                            if (idwYear == DateTime.Now.Year)
                            {
                                RelojResponse relojResponseItem = new RelojResponse();
                                vFic = string.Format(idwYear.ToString("0000")) + "-" + string.Format(idwMonth.ToString("00")) + "-" + string.Format(idwDay.ToString("00")) + " " + String.Format(idwHour.ToString("00")) + ":" + String.Format(idwMinute.ToString("00")) + ":" + String.Format(idwSecond.ToString("00"));
                                //Logger.GetInstance().AddLog(true, "clkZKSoft()", "Fecha del Registro: " + vFic);
                                relojResponseItem.Fich = vFic;
                                relojResponseItem.Nro = pNro;
                                relojResponseItem.SdwEnrollNumber = sdwEnrollNumber;
                                relojResponseItem.IdwVerifyMode = idwVerifyMode;
                                relojResponseItem.IdwInOutMode = idwInOutMode;
                                relojResponseItem.IdwWorkcode = idwWorkcode;

                                relojResponse.Add(relojResponseItem);
                            }
                        }

                        //SAVE IN DATABASE.
                        SaveInDataBase(pRid, relojResponse);

                        if (vClean)
                        {
                            //Logger.GetInstance().AddLog(true, "clkZKSoft()", "BLOQUE COMENTADO: Vaciar Reloj " + pNro + " Ip: " + pDir + ":" + pPor);

                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "Vaciar Reloj " + pNro + " Ip: " + pDir + ":" + pPor);
                            if (Reloj.ClearGLog(pNro))
                            {
                                Reloj.RefreshData(pNro);
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Se vació RelojId " + pNro + " Ip: " + pDir + ":" + pPor);
                            }
                            else
                            {
                                int idwErrorCode = 0;
                                Reloj.GetLastError(idwErrorCode);
                                Logger.GetInstance().AddLog(false, "clkZKSoft()", "Error al vaciar " + pDir + ":" + pPor + " " + idwErrorCode);
                            }
                        }
                    }
                    else
                        Logger.GetInstance().AddLog(true, "clkZKSoft()", "No hay fichadas en " + pDir + ":" + pPor);

                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Desconectar Reloj " + pDir + ":" + pPor);

                    Reloj.Disconnect();

                    clkZKSoft = true;
                }
                else
                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Sin conexión a " + pDir + ":" + pPor);
            }
            catch (Exception ex)
            {
                Logger.GetInstance().AddLog(false, "clkZKSoft()", ex.Message);
            }
            return clkZKSoft;
        }

        private void SaveInDataBase(int? pRid, List<RelojResponse> relojResponse)
        {

            Logger.GetInstance().AddLog(true, "SaveInDataBase()", string.Format("Hay {0} registros para guardar", relojResponse.Count));
            try
            {
                if (relojResponse.Count > 0)
                {
                    using (SqlConnection con = new SqlConnection(dBServer1))
                    {
                        using (SqlCommand cmd = new SqlCommand("sp_SetFichadaReloj", con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            con.Open();

                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "Envios a Server1: " + dBServer1);
                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "Cantidad de Registros: " + relojResponse.Count);

                            foreach (RelojResponse itemResponse in relojResponse)
                            {
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Revisando Legajo: " + itemResponse.SdwEnrollNumber);
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Fecha del Registro: " + itemResponse.Fich);
                                try
                                {

                                    cmd.Parameters.Add("@reloj", SqlDbType.Int).Value = pRid ?? itemResponse.Nro;
                                    cmd.Parameters.Add("@legajo", SqlDbType.VarChar, 10).Value = itemResponse.SdwEnrollNumber;
                                    cmd.Parameters.Add("@tipomov", SqlDbType.TinyInt).Value = itemResponse.IdwInOutMode;
                                    cmd.Parameters.Add("@fechahora", SqlDbType.DateTime).Value = DateTime.Parse(itemResponse.Fich);
                                    cmd.Parameters.Add("@usuarioId", SqlDbType.BigInt).Value = 0;
                                    //cmd.Parameters.Add("@terminalId", SqlDbType.).Value = 0;
                                    //Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@termin(1alId\", SqlDbType.VarChar).Value = 0;");
                                    SqlParameter p2 = new SqlParameter("@terminalId", SqlDbType.Decimal);
                                    p2.Precision = 18;
                                    p2.Scale = 0;
                                    p2.Value = 0;
                                    cmd.Parameters.Add(p2);

                                    cmd.Parameters.Add("@execRdo", SqlDbType.Int).Direction = ParameterDirection.Output;
                                    cmd.Parameters.Add("@execMsg", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output;
                                    cmd.Parameters.Add("@execId", SqlDbType.BigInt).Direction = ParameterDirection.Output;

                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@reloj\", SqlDbType.VarChar).Value: " + (pRid ?? itemResponse.Nro).ToString());
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@legajo\", SqlDbType.VarChar).Value " + itemResponse.SdwEnrollNumber);
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@tipomov\", SqlDbType.VarChar).Value: " + itemResponse.IdwInOutMode);
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@fechahora\", SqlDbType.VarChar).Value: " + itemResponse.Fich);
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@usuarioId\", SqlDbType.VarChar).Value: 0");
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@terminalId\", SqlDbType.VarChar).Value = 0;");


                                    if (this.saveInDB == 1)
                                    {
                                        cmd.ExecuteNonQuery();
                                        Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.ExecuteNonQuery() run ok");
                                    }
                                    else
                                    {
                                        Logger.GetInstance().AddLog(true, "clkZKSoft()", "saveInDB is false, cmd.ExecuteNonQuery() NOT RUN");
                                    }


                                    //Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters[\"@execRdo\"].Value: " + cmd.Parameters["@execRdo"].Value.ToString());
                                    //Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters[\"@execMsg\"].Value: " + cmd.Parameters["@execMsg"].Value.ToString());
                                    //Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters[\"@execId\"].Value: " + cmd.Parameters["@execId"].Value.ToString());

                                    if (Convert.ToBoolean(cmd.Parameters["@execRdo"].Value))
                                        Logger.GetInstance().AddLog(false, "clkZKSoft()", "SetFichada Error en set fichada (execMsg): " + cmd.Parameters["@execMsg"].Value.ToString());
                                    else
                                        Logger.GetInstance().AddLog(true, "clkZKSoft()", "SetFichada OK");
                                }
                                catch (Exception ex)
                                {
                                    Logger.GetInstance().AddLog(false, "clkZKSoft()", "Excepción en SetFichada: " + ex.Message);
                                }

                                cmd.Parameters.Clear();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.GetInstance().AddLog(false, "SaveInDataBase()", ex.Message);
            }
        }


        //private void FuncionPrueba(int pRid)
        //{
        //    if (ConnectServer(dBServer1))
        //    {
        //        Logger.GetInstance().AddLog(true, "FuncionPrueba()", "pRid: " + pRid.ToString());
        //        Logger.GetInstance().AddLog(true, "FuncionPrueba()", "Envios a Server1: " + dBServer1);
        //        FuncionPrueba2(1, "9109");
        //        modDeclares.ShamanSession.Cerrar(modDeclares.ShamanSession.PID);
        //    }
        //}
        //private static DevOps FuncionPrueba2(int pRid, string SdwEnrollNumber)
        //{
        //    RelojesIngresos objFichada = new RelojesIngresos();
        //    DevOps devolucionOperacion = objFichada.SetFichada(pRid, SdwEnrollNumber, "2018-07-10 15:13:14", "CLOCK");
        //    Logger.GetInstance().AddLog(true, "FuncionPrueba2()", "SetFichada CacheDebug: " + devolucionOperacion.CacheDebug);
        //    if (devolucionOperacion.Resultado)
        //        Logger.GetInstance().AddLog(true, "FuncionPrueba2()", "SetFichada OK");
        //    else
        //        Logger.GetInstance().AddLog(false, "FuncionPrueba2()", "SetFichada Error: " + devolucionOperacion.DescripcionError);
        //    return devolucionOperacion;
        //}
    }
}
