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
    public class Driver1
    {
        private IWebDriver driver;
        private String browser;
        private String executionType;

        public static Queue<string> testCaseIDorGroup = new Queue<string>();

        [SetUp]
        public void Init()
        {
            Boolean status = false;
            HTML.fnSummaryInitialization("Execution Summary Report");
            status = XlsxReader.getInstance().addTestCasesFromDataSheetName(testCaseIDorGroup);
            //testCaseIDorGroup.Enqueue("TC01");
        }

        [Test]
        public void TimeLive1()
        {
            Boolean isExitLoop = false;
            String threadId = System.Threading.Thread.CurrentThread.ManagedThreadId.ToString();
            Console.WriteLine(threadId);

            int x = Convert.ToInt32(ConfigurationManager.AppSettings["ShortSyncTime"]);
            for (int i = 1; i < x; i++)
            {
                if (testCaseIDorGroup.Count > 0)
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
                strTCID = testCaseIDorGroup.Dequeue();
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
                CommonManager.getInstance().SetCommon(common);

                ConfigManager cm = new ConfigManager();
                ThreadCache.getInstance().setConfigManager(cm);

                common.RunTestCase(strTCID);

                try
                {
                    strTCID = testCaseIDorGroup.Dequeue();
                }
                catch (Exception e)
                {
                    isExitLoop = true;
                    testCaseIDorGroup.Enqueue("Done");
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
            int ThreadCount = Convert.ToInt32(ConfigurationManager.AppSettings["ThreadCount"]);
            for (int i = 0; i < 30; i++)
            {
                int j = 0;
                foreach (string item in testCaseIDorGroup)
                {
                    if (item.Equals("Done"))
                    {
                        j = j + 1;
                    }
                }
                if (ThreadCount == j)
                {
                    break;
                }
                Thread.Sleep(1000);
            }
            HTML.fnSummaryCloseHtml(ConfigurationManager.AppSettings["Release"]);
        }
    }
}