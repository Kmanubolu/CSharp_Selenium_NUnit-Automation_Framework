using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using System.Web;
using System.Net;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;

namespace TimeLiveRegression.org.com.utilities
{
    class LocalDriverFactory
    {
        private static LocalDriverFactory instance = new LocalDriverFactory();

        public static LocalDriverFactory getInstance()
        {
            return instance;
        }

        public IWebDriver CreateNewDriver(String browser)
        {
                IWebDriver driver = null;

                if (browser.ToUpper() == "FIREFOX")
                {
                    driver = new FirefoxDriver();
                }
                else if (browser.ToUpper() == "CHROME")
                {
                    driver = new ChromeDriver();
                }
                else if (browser.ToUpper() == "IE")
                {
                    driver = new InternetExplorerDriver();
                }
            return driver;
        }
    }
}
