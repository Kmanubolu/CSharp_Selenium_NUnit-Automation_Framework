using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;

namespace SeleniumTests
{
    [TestFixture]
    public class Sample
    {
        private IWebDriver driver;
        private StringBuilder verificationErrors;
        private string baseURL;
        private bool acceptNextAlert = true;
        
        [SetUp]
        public void SetupTest()
        {
            driver = new FirefoxDriver();
            baseURL = "https://timelive.livetecs.com/";
            verificationErrors = new StringBuilder();
        }
        
        [TearDown]
        public void TeardownTest()
        {
            try
            {
                driver.Quit();
            }
            catch (Exception)
            {
                // Ignore errors if unable to close the browser
            }
            Assert.AreEqual("", verificationErrors.ToString());
        }
        
        [Test]
        public void TheSampleTest()
        {
            driver.Navigate().GoToUrl(baseURL + "/");
            driver.FindElement(By.Id("CtlLogin1_Login1_UserName")).Clear();
            driver.FindElement(By.Id("CtlLogin1_Login1_UserName")).SendKeys("krishna2@gmail.com");
            driver.FindElement(By.Id("CtlLogin1_Login1_Password")).Clear();
            driver.FindElement(By.Id("CtlLogin1_Login1_Password")).SendKeys("India123!@#");
            driver.FindElement(By.Id("CtlLogin1_Login1_Button1")).Click();
            driver.FindElement(By.CssSelector("#open-menu > span")).Click();
            driver.FindElement(By.Id("R_H2_13")).Click();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_btnAddEmployee")).Click();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_FirstNameTextBox")).Click();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_FirstNameTextBox")).Clear();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_FirstNameTextBox")).SendKeys("Ram");
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_MiddleNameTextBox")).Clear();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_MiddleNameTextBox")).SendKeys("Reddy");
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_LastNameTextBox")).Click();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_LastNameTextBox")).Clear();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_LastNameTextBox")).SendKeys("Redd");
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_EMailAddressTextBox")).Click();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_EMailAddressTextBox")).Clear();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_EMailAddressTextBox")).SendKeys("Ram123@gmail.com");
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_PasswordTextBox")).Clear();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_PasswordTextBox")).SendKeys("Ram123!@#");
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_VerifyPasswordTextbox")).Clear();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_VerifyPasswordTextbox")).SendKeys("Ram123!@#");
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_ddlAccountDepartmentId")).Click();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_ddlEmployeeStatusId")).Click();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_ddlEmployeeStatusId")).Click();
            new Select(driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_ddlEmployeeStatusId"))).SelectByText("Resigned");
            new Select(driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_ddlAccountRoleId"))).SelectByText("User");
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_ddlAccountRoleId")).Click();
            new Select(driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_ddlAccountRoleId"))).SelectByText("Administrator");
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_ddlAccountLocationId")).Click();
            new Select(driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_ddlAccountLocationId"))).SelectByText("Ella");
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_chkIsShowEmployeePicture")).Click();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_chkIsShowEmployeePicture")).Click();
            driver.FindElement(By.Id("C_C_C_CtlAccountEmployeeForm1_FormView1_Add")).Click();
            Assert.AreEqual("Employees", driver.Title);
        }
        private bool IsElementPresent(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        
        private bool IsAlertPresent()
        {
            try
            {
                driver.SwitchTo().Alert();
                return true;
            }
            catch (NoAlertPresentException)
            {
                return false;
            }
        }
        
        private string CloseAlertAndGetItsText() {
            try {
                IAlert alert = driver.SwitchTo().Alert();
                string alertText = alert.Text;
                if (acceptNextAlert) {
                    alert.Accept();
                } else {
                    alert.Dismiss();
                }
                return alertText;
            } finally {
                acceptNextAlert = true;
            }
        }
    }
}
