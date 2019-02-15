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
using System.Configuration;

namespace TimeLiveRegression.org.com.utilities
{
    class RemoteDriverFactory
    {
        private static RemoteDriverFactory instance = new RemoteDriverFactory();




        public static RemoteDriverFactory getInstance()
        {
            return instance;
        }

        public IWebDriver CreateNewDriver(String browser)
        {
            RemoteWebDriver driver = null;
            DesiredCapabilities caps = null;
            if (browser.ToUpper() == "FIREFOX")
            {
                caps = new DesiredCapabilities();
                caps.SetCapability(CapabilityType.BrowserName, "firefox");
            }
            else if (browser.ToUpper() == "CHROME")
            {
                caps = new DesiredCapabilities();
                caps.SetCapability(CapabilityType.BrowserName, "chrome");
            }
            else if (browser.ToUpper() == "IE")
            {
                caps = new DesiredCapabilities();
                caps.SetCapability(CapabilityType.BrowserName, "internetExplorer");
            }
            driver = new RemoteWebDriver(new Uri("http://localhost:4444/wd/hub"), caps, TimeSpan.FromSeconds(600));
            return driver;
        }
    }
}
