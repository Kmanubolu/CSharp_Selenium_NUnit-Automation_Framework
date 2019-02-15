using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TimeLiveRegression.org.com.utilities;

namespace TimeLiveRegression.org.com.screen
{
    public class LogIn
    {

        public static String sheetname = "LogIn";
        Common common = CommonManager.getInstance().GetCommon();

        public Boolean SCRLogIn()
        {
            Boolean status = true;
            //Common common = Common();
            status = OpenApp();
            status = common.RunComponent(sheetname, Common.o);
            if (!status)
            {
                status = false;
            }
            return status;
        }


        public Boolean OpenApp()
        {
            ManagerDriver.getInstance().GetDriver().Manage().Window.Maximize();
            ManagerDriver.getInstance().GetDriver().Navigate().GoToUrl(ConfigurationManager.AppSettings[ConfigurationManager.AppSettings["Region"]]);
            return true;
        }

    }
}
