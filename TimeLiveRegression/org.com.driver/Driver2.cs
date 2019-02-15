using NUnit.Framework;
using System;
using OpenQA.Selenium;
using System.Configuration;
using TimeLiveRegression.org.com.utilities;
using System.Collections.Generic;
using System.Threading;

namespace TimeLiveRegression
{
    [TestFixture]
    [Parallelizable]
    public class Driver2
    {
        private IWebDriver driver;
        private String browser;
        private String executionType;

        //public static Queue<string> testCaseIDorGroup = new Queue<string>();

        [SetUp]
        public void Init()
        {
            //Console.WriteLine("Hi");
            //testCaseIDorGroup.Enqueue("TC02");

        }

        [Test]
        public void TimeLive2()
        {
            Boolean isExitLoop = false;
            String threadId = System.Threading.Thread.CurrentThread.ManagedThreadId.ToString();
            Console.WriteLine(threadId);

            int x = Convert.ToInt32(ConfigurationManager.AppSettings["ShortSyncTime"]);
            for (int i = 1; i < x; i++)
            {
                if (Driver1.testCaseIDorGroup.Count > 0)
                {
                    break;
                }
                else
                {
                    Thread.Sleep(2000);
                }
            }

            String strTCID = null;
            try
            {
                //testCaseName = testCaseIDorGroup.remove();
                strTCID = Driver1.testCaseIDorGroup.Dequeue();
            }
            catch (Exception e)
            {
                isExitLoop = true;
            }
            if (strTCID == null)
            {
                isExitLoop = true;
            }

            while (!isExitLoop)
            {
                browser = ConfigurationManager.AppSettings["Browsers"];
                executionType = ConfigurationManager.AppSettings["ExecutionType"];

                if (executionType.ToUpper() == "REMOTE")
                {
                    driver = RemoteDriverFactory.getInstance().CreateNewDriver(browser);
                }
                else if (executionType.ToUpper() == "LOCAL")
                {
                    driver = LocalDriverFactory.getInstance().CreateNewDriver(browser);
                }
                ManagerDriver.getInstance().SetDriver(driver);
                Common common = new Common();

                ConfigManager cm = new ConfigManager();
                ThreadCache.getInstance().setConfigManager(cm);

                CommonManager.getInstance().SetCommon(common);
                common.RunTestCase(strTCID);

                try
                {
                    strTCID = Driver1.testCaseIDorGroup.Dequeue();
                }
                catch (Exception e)
                {
                    isExitLoop = true;
                    Driver1.testCaseIDorGroup.Enqueue("Done");
                }
                if (strTCID == null || strTCID == "Done")
                {
                    isExitLoop = true;
                }

            }
        }

        [TearDown]
        public void CleanUp()
        {
           
        }
    }
}