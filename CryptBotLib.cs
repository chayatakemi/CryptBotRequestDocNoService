using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CryptBotRequestDocNoService
{

    public sealed class CryptBotLib : MarshalByRefObject
    {
        private readonly dynamic _application;
        public CryptBotLib()
        {
            try
            {
                const string progId = "CryptBotLib.Utility";
                _application = Activator.CreateInstance(Type.GetTypeFromProgID(progId));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        /*public void Quit()
        {
            _application.Quit();
        }*/

        public string GetLicenseInfo
        {
            get { return _application.getLicenseInfo("LicenseKey4"); }
        }
    }
}
