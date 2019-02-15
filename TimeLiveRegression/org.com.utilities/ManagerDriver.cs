using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;

namespace TimeLiveRegression.org.com.utilities
{
    class ManagerDriver
    {
        private static ManagerDriver instance = new ManagerDriver();
       ThreadLocal<IWebDriver> driver = new ThreadLocal<IWebDriver>();

        public static ManagerDriver getInstance()
        {
            return instance;
        }

        
        public void SetDriver(IWebDriver d)
        {
           this.driver.Value = d;
        }

        public IWebDriver GetDriver()
        {
            return driver.Value;
        }
       
    }
}
