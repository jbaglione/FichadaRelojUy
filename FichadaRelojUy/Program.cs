using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace FichadaRelojUyService
{
    static class Program
    {
        public static void Main(string[] args)
        {
            if (args.FirstOrDefault()?.ToUpper() == "/CONSOLE")
            {
                RunAsConsole();
            }
            else
            {
                RunAsService();
            }
        }
        private static void RunAsConsole()
        {
            Service1 serv = new Service1();
            serv.StartService();

            Console.WriteLine("Running service as console. Press any key to stop.");
            Console.ReadKey();

            serv.Stop();
        }
        private static void RunAsService()
        {
            //#if (!DEBUG)
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
                {
                new Service1()
                };
            ServiceBase.Run(ServicesToRun);
            //#else
            //Service1 myServ = new Service1();
            //// here Process is my Service function
            //// that will run when my service onstart is call
            //// you need to call your own method or function name here instead of Process();
            //#endif
        }
    }
}
