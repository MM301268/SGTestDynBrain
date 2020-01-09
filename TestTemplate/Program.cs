using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumWorker
{
    public class Program
    {
        private static SeleniumWorker.SelWorker _mySelenium;

        private static void Main()
        {
            for (var i = 1; i <= 500; i++)
            {
                _mySelenium = new SeleniumWorker.SelWorker("P_Adressänderung");
                var result = SelWorker.SendRequestToRest(DateTime.Now,
                                                               "Adressänderung",
                                                               "BLN",
                                                               "Fax",
                                                               "true",
                                                               "IP Dynamics\nLyonel-Feininger Str. 28\n80807 München").Result;
                Console.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + " " + result.ToString()) ;
            //_mySelenium.Init(new Uri("http://172.17.10.196:4444/wd/hub"), new Uri("https://iis-srv.post-test.de/Webclient/"), "Chrome");
            //_mySelenium.GetBrowserTitle("Voxtron Web Client");
            //_mySelenium.Login("IPD_Agent1", "IPD_Agent1" , "true");
            ////_mySelenium.Login("IPD_Agent" + i.ToString(), "IPD_Agent" + i.ToString(), "true");
            //_mySelenium.AcceptTask(1,1);
            //_mySelenium.Logout();
            //_mySelenium.Terminate();
            }
        }
    }
}
