using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
namespace CryptBotRequestDocNoService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {

            //if (Environment.UserInteractive)
            //{
            //    var x = new CryptBotRequestDocNoService();
            //    x.TestStartupAndStop(args);
            //    Console.WriteLine("Run as console, enter any key to exit.");
            //    Console.ReadLine();
            //}
            //else
            //{
            //    ServiceBase[] ServicesToRun;
            //    ServicesToRun = new ServiceBase[]
            //    {
            //        new CryptBotRequestDocNoService()
            //    };
            //    ServiceBase.Run(ServicesToRun);
            //}
#if DEBUG
            CryptBotRequestDocNoService myService = new CryptBotRequestDocNoService();
            myService.OnDebug();
            System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);

#else
            ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                new CryptBotRequestDocNoService()
                };
                ServiceBase.Run(ServicesToRun);
#endif



        }
    }
}
