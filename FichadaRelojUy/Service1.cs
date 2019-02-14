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

namespace FichadaRelojUyService
{
    public partial class Service1 : ServiceBase
    {

        Timer t = new Timer();
        private Conexion shamanConexion = new Conexion();
        string m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        string dBServer1 = ConfigurationManager.AppSettings["DBServer1"];

        string timePool = ConfigurationManager.AppSettings["TimePool"];
        //bool setByDoc = ConfigurationManager.AppSettings["SetByDoc"] == "1";

        List<Relojes> relojes;

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
            t.Elapsed += delegate { ElapsedHandler(); };
            t.Interval = 60000;
            t.Start();
            Logger.GetInstance().AddLog(true, "OnStart", "Servicio inicializado.");
            CustomMail mail = new CustomMail(MailType.Information, "Se ha inicializado el servicio de Fichada", "Servicio de Fichada");
            mail.Send();
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
            if (relojes == null || relojes.Count == 0)
            {
                if (LeerRelojes())
                    ProcesarRelojes();
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
                        this.clkZKSoft(item.Id, 1, item.Descripcion, item.DireccionIP, item.Puerto, 0, Convert.ToBoolean(item.Vaciar));
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

        private bool clkZKSoft(int pRid, int pNro, string pDes, string pDir, int pPor, long pPsw, bool vClean = false)
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
                                            cmd.Parameters.Add("@reloj", SqlDbType.Int).Value = pRid;
                                            cmd.Parameters.Add("@legajo", SqlDbType.VarChar, 10).Value = itemResponse.SdwEnrollNumber;                                           
                                            cmd.Parameters.Add("@tipomov", SqlDbType.TinyInt).Value = itemResponse.IdwInOutMode;
                                            cmd.Parameters.Add("@fechahora", SqlDbType.DateTime).Value = DateTime.Parse(itemResponse.Fich);                                          
                                            cmd.Parameters.Add("@usuarioId", SqlDbType.BigInt).Value = 0;
                                            //cmd.Parameters.Add("@terminalId", SqlDbType.).Value = 0;
                                            //Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@terminalId\", SqlDbType.VarChar).Value = 0;");
                                            SqlParameter p2 = new SqlParameter("@terminalId", SqlDbType.Decimal);
                                            p2.Precision = 18;
                                            p2.Scale = 0;
                                            p2.Value = 0;
                                            cmd.Parameters.Add(p2);

                                            cmd.Parameters.Add("@execRdo", SqlDbType.Int).Direction = ParameterDirection.Output;
                                            cmd.Parameters.Add("@execMsg", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output;
                                            cmd.Parameters.Add("@execId", SqlDbType.BigInt).Direction = ParameterDirection.Output;

                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@reloj\", SqlDbType.VarChar).Value: " + pRid.ToString());
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@legajo\", SqlDbType.VarChar).Value " + itemResponse.SdwEnrollNumber);
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@tipomov\", SqlDbType.VarChar).Value: " + itemResponse.IdwInOutMode);
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@fechahora\", SqlDbType.VarChar).Value: " + itemResponse.Fich);
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@usuarioId\", SqlDbType.VarChar).Value: 0");
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters.Add(\"@terminalId\", SqlDbType.VarChar).Value = 0;");

                                            //SqlDataReader reader = cmd.ExecuteReader();
                                            cmd.ExecuteNonQuery();
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.ExecuteNonQuery() run ok");

                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters[\"@execRdo\"].Value: " + cmd.Parameters["@execRdo"].Value.ToString());
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters[\"@execMsg\"].Value: " + cmd.Parameters["@execMsg"].Value.ToString());
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "cmd.Parameters[\"@execId\"].Value: " + cmd.Parameters["@execId"].Value.ToString());

                                            if (Convert.ToBoolean(cmd.Parameters["@execRdo"].Value))
                                                Logger.GetInstance().AddLog(false, "clkZKSoft()", "SetFichada Error en set fichada (execMsg): " + cmd.Parameters["@execMsg"].Value.ToString());
                                            else
                                                Logger.GetInstance().AddLog(false, "clkZKSoft()", "SetFichada OK");
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
