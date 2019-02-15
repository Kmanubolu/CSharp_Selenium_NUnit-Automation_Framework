using System;
using System.Configuration;
using TimeLiveRegression.org.com.elements;
using System.Reflection;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System.Threading;

namespace TimeLiveRegression.org.com.utilities
{
    public class Common
     {

        private XlsxReader sXL;
        public static Elements o = new Elements();


        public Common()
        {

        }

        #region for RunTestCase ...
        public Boolean RunTestCase(String strTCID)
         {
            ConfigManager cm = ThreadCache.getInstance().getConfigManager();
            Boolean blnRunFlag = true;
            sXL = XlsxReader.getInstance();
            ADODB.Recordset rs = sXL.getRecordSet("select* from[TestCases$] where Execution = 'YES' and TCID = '" + strTCID + "'");
             while (rs.EOF == false)
             {
                string strTestCaseNo = Convert.ToString(rs.Fields["TCID"].Value);
                string strTestCaseName = Convert.ToString(rs.Fields["TestCaseName"].Value);
                HTML.fnInitilization(strTestCaseNo + "-" + strTestCaseName);

                //ConfigurationManager.AppSettings["TCID"] = strTestCaseNo;
                //ConfigurationManager.AppSettings["TestCaseName"] = strTestCaseName;
                cm.setProperty("TCID", strTestCaseNo);
                cm.setProperty("TestCaseName", strTestCaseName);

                for (int i = 0; i < rs.Fields.Count; i++)
                {
                    string strColumnName = rs.Fields[i].Name; 
                    if (strColumnName.Contains("Component"))
                    {
                            string strComponentName = Convert.ToString(rs.Fields[i].Value);
                        if ((!strComponentName.Equals("")))
                        {
                            //ConfigurationManager.AppSettings["ComponentName"] = strComponentName;
                            cm.setProperty("ComponentName", strComponentName);

                            HTML.fnInsertResult("Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), cm.getProperty("ComponentName") , strComponentName + " Component execution should be started", strComponentName + " Component execution has been started", "PASS");

                            string AssemblyName = "TimeLiveRegression";
                            string typeName = "TimeLiveRegression.org.com.screen.{0}, " + AssemblyName;

                            string strClassName = strComponentName;
                            string strCompoName = "SCR" + strComponentName;
                            try
                            {
                                string innerTypeName = string.Format(typeName, strClassName);
                                Type type1 = Type.GetType(innerTypeName);
                                object obj1 = Activator.CreateInstance(type1);
                                MethodInfo methodInfo1 = type1.GetMethod(strCompoName);
                                blnRunFlag = (Boolean)methodInfo1.Invoke(obj1, null);
                            }
                            catch (Exception e)
                            {
                                blnRunFlag = false;
                            }
                            if (blnRunFlag != true)
                            {
                                blnRunFlag = false;
                                break;
                            }
                            else
                            {
                                blnRunFlag = true;
                                HTML.fnInsertResult("Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), cm.getProperty("ComponentName"), strCompoName + " Component execution should be executed successfully", strCompoName + " Component  executed successfully", "PASS");
                            }
                        }
                    }
                }
                if (blnRunFlag)
                {
                    blnRunFlag = true;
                    HTML.fnSummaryInsertTestCase();
                    CommonManager.getInstance().GetCommon().Terminate();
                }
                else
                {
                    blnRunFlag = false;
                    HTML.fnInsertResult("Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), cm.getProperty("ComponentName") , cm.getProperty("ComponentName") + " Component should be executed successfully", cm.getProperty("ComponentName") + " Component not executed successfully", "FAIL");
                    HTML.fnSummaryInsertTestCase();
                    CommonManager.getInstance().GetCommon().Terminate();
                }
                rs.MoveNext();
             }
            rs.Close();
            return blnRunFlag;
         }
        #endregion
        #region for RunComponent...
        public Boolean RunComponent(String sheetname, Elements o)
        {
            ConfigManager cm = ThreadCache.getInstance().getConfigManager();
            Boolean status = false;
            sXL = XlsxReader.getInstance();
            String sql = "select* from[" + sheetname + "$] where TCID = '" + cm.getProperty("TCID") + "'";
            ADODB.Recordset rs = sXL.getRecordSet(sql);
            while (rs.EOF == false)
            {
                for (int i = 0; i < rs.Fields.Count; i++)
                {
                    string strColumnName = rs.Fields[i].Name;
                    string strClassName = strColumnName.Substring(0, 3);
                    if (strClassName.Equals("ele") || strClassName.Equals("edt") || strClassName.Equals("btn") || strClassName.Equals("lst") || strClassName.Equals("fun") || strClassName.Equals("cfu"))
                    {
                        string strValue = rs.Fields[i].Value;
                        if ((!strValue.Equals("")))
                        {

                            status = SafeAction(o.getObject(strColumnName), strValue, strColumnName);
                        }
                        if (!status)
                        {
                            return false;
                        }
                    }
                }
                rs.MoveNext();
            }
            rs.Close();
            return status;
        }
        #endregion

        #region for SafeAction...
        public Boolean SafeAction(By element, String strValue, String ColumnName)
        {
            ConfigManager cm = ThreadCache.getInstance().getConfigManager();
            Boolean blnRunFlag = true;
            IWebElement obj = null;
            Common common = CommonManager.getInstance().GetCommon();
            string strClass = ColumnName.Substring(0, 3);
            IJavaScriptExecutor js = (IJavaScriptExecutor)ManagerDriver.getInstance().GetDriver();

           
            obj = ManagerDriver.getInstance().GetDriver().FindElement(element);
            if (obj.Displayed)
            {
                blnRunFlag = true;
                highlightElement(obj);
            }
            else
            {
                HTML.fnInsertResult("Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), cm.getProperty("ComponentName") + "- Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), "Verification: " + ColumnName + " object should be displayed", "Verification: " + ColumnName + " object is not displayed", "FAIL");
                blnRunFlag = false;
            }
            if (blnRunFlag)
            {
                if (strClass == "edt")
                {
                    obj.SendKeys(strValue);
                    //common.JavaScriptDynamicWait(obj, js);
                    HTML.fnInsertResult("Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), cm.getProperty("ComponentName"), "Data should be entered '" + strValue + "' in " + ColumnName + " editbox", "Data entered '" + strValue + "' in " + ColumnName + " editbox", "PASS");
                }
                if (strClass == "lst")
                {
                    SelectElement ss = new SelectElement(obj);
                    ss.SelectByText(strValue);
                    //common.JavaScriptDynamicWait(obj, js);
                    HTML.fnInsertResult("Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), cm.getProperty("ComponentName"), "Data should be selected '" + strValue + "' from " + ColumnName + " listbox", "Data should be selected '" + strValue + "' from " + ColumnName + " listbox", "PASS");
                }
                if (strClass == "com")
                {
                    obj.Clear();
                    obj.SendKeys(strValue);
                    obj.SendKeys(Keys.Tab);
                    //common.JavaScriptDynamicWait(obj, js);
                    HTML.fnInsertResult("Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), cm.getProperty("ComponentName"), "Data should be selected '" + strValue + "' from " + ColumnName + " listbox", "Data should be selected '" + strValue + "' from " + ColumnName + " listbox", "PASS");
                }
                if (strClass == "ele")
                {
                    obj.Click();
                    //common.JavaScriptDynamicWait(obj, js);
                    HTML.fnInsertResult("Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), cm.getProperty("ComponentName"), "Click on the element '" + ColumnName + "'", "Clicked on the element '" + ColumnName + "'", "PASS");
                }
                if (strClass == "btn")
                {
                    obj.Click();
                    //common.JavaScriptDynamicWait(obj, js);
                    HTML.fnInsertResult("Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), cm.getProperty("ComponentName"), "Click on the element '" + ColumnName + "'", "Clicked on the element '" + ColumnName + "'", "PASS");
                }
            }
            else
            {
                HTML.fnInsertResult("Thread ID - " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(), cm.getProperty("ComponentName"), "Verification: " + ColumnName + " object should be enabled", "Verification: " + ColumnName + " object is disabled", "FAIL");
                blnRunFlag = false;
            }
            return blnRunFlag;
        }
        #endregion
        public void Terminate()
        {
            ManagerDriver.getInstance().GetDriver().Quit();
        }



        public  Boolean JavaScriptDynamicWait(IWebElement sElement, IJavaScriptExecutor js)
        {
            Boolean status = false;
            int x = Convert.ToInt32(ConfigurationManager.AppSettings["VeryLongSyncTime"]);
            for (int i = 1; i <= x; i++) {

	            Boolean isAjaxRunning = Boolean.Parse(js
                              .ExecuteScript(
                                           "return Ext.Ajax.isLoading();") //returns true if ajax call is currently in progress
                              .ToString());
	            if (!isAjaxRunning) {
	            	status = true;
	                   break;
	            }
                Thread.Sleep(1000);//wait for one secnod then check if ajax is completed
	        }
            return status;
	    }

        /**
	     * @function Highlights on current working element or locator
	     * @param driver
	     * @param locator
	     * @throws Exception
	     */
        public void highlightElement(IWebElement obj)
        {
		    if(ConfigurationManager.AppSettings["HighlightElements"].ToLower() == ("true"))
            {
                String attributevalue = "border:10px solid green;";//change border width and colour values if required
                IJavaScriptExecutor executor = (IJavaScriptExecutor)ManagerDriver.getInstance().GetDriver();
                String getattrib = obj.GetAttribute("style");
                executor.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", obj, attributevalue);
                Thread.Sleep(100);
                executor.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", obj, getattrib);
            }
        }
    }

}
