using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
//using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.IE;
using OpenQA;
using OpenQA.Selenium.Support.UI;
using SHDocVw;


namespace GlobalLibrary
{
    /// <summary>
    /// GlobalSeleniumMethods is the class for the assembly GlobalLibrary. Contains methods 
    /// and data to encapsulate the Selenium api method calls to automate webpage controls
    /// </summary>
    class GlobalSeleniumMethods
    {
        ///<summary>Public IWebDriver declaration </summary>
        public IWebDriver Driver;

        ///<summary>Public Declaration for IE </summary>
        public InternetExplorerOptions options;


        /// <summary>
        /// GlobalSeleniumMethods constructor, creates new instance of the Internet Explorer driver,
        /// makes this instance available for the other methods to acess the Driver object.
        /// </summary>
        public GlobalSeleniumMethods()
        {
            //Close All IE browsers before opening a new one
            //CloseAllIEBrowser();

            //instantiate new instance of IWebDriver
            options = new InternetExplorerOptions();
            options.IgnoreZoomLevel = false;
            options.EnableNativeEvents = true;
            //Driver = new InternetExplorerDriver("W:\Projects\Automation\Installation\IEDriver\IEDriverServer_Win32_2.32.3\IEDriverServer.exe")
            Driver = new InternetExplorerDriver(options);
        }


        /// <summary>
        /// GlobalSeleniumMethods constructor passes existing instance of Driver to public member available
        /// for the other methods to acess the Driver object
        /// </summary>
        public GlobalSeleniumMethods(IWebDriver iwdDriver, int intBrowser)
        {
            switch (intBrowser)
            {
                case (int)eBrowser.FireFox:
                    Driver = iwdDriver;
                    break;
                case (int)eBrowser.IE:
                    options = new InternetExplorerOptions();
                    options.IgnoreZoomLevel = true;
                    options.EnableNativeEvents = true;
                    Driver = iwdDriver;
                    break;
            }
        }


        /// <summary>
        /// CloseAllIEBrowser() method is used to close all processes for Internet Explorer
        /// <example> <code>
        /// bRtn = Close AllIEBrowser()
        /// </code> </example>
        /// </summary>
        /// <returns>IWebDriver object</returns>
        public bool CloseAllIEBrowser()
        {
            try
            {
                System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName(GlobalVariables.ieProcess);
                foreach (System.Diagnostics.Process proc in procs)
                {
                    proc.Kill();
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                return false;
            }
        }


        /// <summary>
        /// GetDriver() method is used to return the existing instantiated Driver object
        /// </summary>
        /// <returns>IWebDriver object</returns>
        public IWebDriver GetDriver()
        {
            return Driver;
        }


        /// <summary>
        /// Sets the Native browser events
        /// </summary>
        /// <param name="bValue">true turns on browser events, false truns off browser events</param>
        /// <returns>True/False</returns>
        public bool SetBrowserEvents(bool bValue)
        {
            bool bRtn = true;
            try
            {
                options.EnableNativeEvents = bValue;
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                bRtn = false;
            }
            return bRtn;
        }


        /// <summary>
        /// OpenBrowser launches the IE browser and loads the passed url
        /// </summary>
        /// <param name="url">string, url of webpage that is loaded upon browser execution</param>
        /// <param name="sCtrlWaitFor">Pass a string, control that is waited for when browser loads</param>
        /// <param name="url">Pass a int, html control type, see the Enum called CtrlType</param>
        /// <returns>bool, true if method succeeds, false if method fails</returns>
        public bool OpenBrowser(string sURL, string sCtrlWaitFor, int iWaitForSec, int iCtrlType)
        {
            bool bRtn = true;
            IWebElement iWeRtn;
            try
            {
                // Launch new browser and load passed URL
                Driver.Navigate().GoToUrl(sURL);

                // Wait iWaitForSec (seconds) for sCtrlWaitFor (control) to appear
                iWeRtn = this.WaitForElement(sCtrlWaitFor, iWaitForSec, iCtrlType);

                if (iWeRtn.Equals(null))
                {
                    bRtn = false;
                }
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = OpenBrowser(" + sURL + "," + sCtrlWaitFor + "," + iWaitForSec + "," + iCtrlType + ")");
                Log.Error(ex);
                bRtn = false;
            }
            return bRtn;
        }


        /// <summary>
        /// BrowserPageForward() method performs the same action as pressing the Browser Forward button
        /// </summary>
        /// <return>bool, returns true if no exception is fired, flose if exception is fired</return>
        public bool BrowserPageForward()
        {
            bool bRtn = true;
            try
            {
                Driver.Navigate().Forward();
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                bRtn = false;
            }
            return bRtn;
        }


        /// <summary>
        /// BrowserPageBack() method performs the same action as pressing the Browser Back button
        /// </summary>
        /// <return>bool, returns true if no exception is fired, flose if exception is fired</return>
        public bool BrowserPageBack()
        {
            bool bRtn = true;
            try
            {
                Driver.Navigate().Back();
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                bRtn = false;
            }
            return bRtn;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="strCtrl"></param>
        /// <param name="intWaitSeconds"></param>
        /// <param name="intCtrlType"></param>
        /// <returns></returns>
        public IWebElement WaitForElement(string strCtrl, int intWaitSeconds, int intCtrlType)
        {
            IWebElement element = null;
            IWebElement myDynamicElement = null;

            try
            {
                WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(intWaitSeconds));
                System.Threading.Thread.Sleep(3000);
                myDynamicElement = wait.Until<IWebElement>((d) =>
                {
                    switch (intCtrlType)
                    {
                        case (int)CtrlType.ClassName:
                            element = d.FindElement(By.ClassName(strCtrl));
                            break;
                        case (int)CtrlType.CssSelector:
                            element = d.FindElement(By.CssSelector(strCtrl));
                            break;
                        case (int)CtrlType.ID:
                            element = d.FindElement(By.Id(strCtrl));
                            break;
                        case (int)CtrlType.LinkText:
                            element = d.FindElement(By.LinkText(strCtrl));
                            break;
                        case (int)CtrlType.Name:
                            element = d.FindElement(By.Name(strCtrl));
                            break;
                        case (int)CtrlType.PartialLinkText:
                            element = d.FindElement(By.PartialLinkText(strCtrl));
                            break;
                        case (int)CtrlType.TagName:
                            element = d.FindElement(By.TagName(strCtrl));
                            break;
                        case (int)CtrlType.XPath:
                            element = d.FindElement(By.XPath(strCtrl));
                            break;
                    }
                    return element;

                });
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = WaitForElement(" + strCtrl + "," + intWaitSeconds + "," + intCtrlType + ")");
                Log.Error(ex);
                myDynamicElement = null;
            }
            return myDynamicElement;

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="iWidth"></param>
        /// <param name="iHight"></param>
        /// <returns></returns>
        public bool ScreenResize(int iWidth, int iHight)
        {
            bool bRtn = true;

            try
            {
                Driver.Manage().Window.Size = new System.Drawing.Size(iWidth, iHight);
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = ScreenResize(" + iWidth + "," + iHight + ")");
                Log.Error(ex);
                bRtn = false;
            }
            return bRtn;
        }

        /// <summary>
        /// 
        /// </summary>
        public void Quit()
        {

            try
            {
                Driver.Quit();
            }
            catch (StaleElementReferenceException ex)
            { }
        }


        /// <summary>
        /// PerformAction() method version 1
        /// </Summary>
        public bool PerformAction(string strCtrl, int intCtrlType, int intCtrlClass)
        {
            bool bRtn = true;
            IWebElement iWebRtn;

            try
            {
                IWebElement element = null;

                switch (intCtrlType)
                {
                    case (int)CtrlType.ClassName:
                        element = Driver.FindElement(By.ClassName(strCtrl));
                        break;
                    case (int)CtrlType.CssSelector:
                        element = Driver.FindElement(By.CssSelector(strCtrl));
                        break;
                    case (int)CtrlType.ID:
                        element = Driver.FindElement(By.Id(strCtrl));
                        break;
                    case (int)CtrlType.LinkText:
                        element = Driver.FindElement(By.LinkText(strCtrl));
                        break;
                    case (int)CtrlType.Name:
                        element = Driver.FindElement(By.Name(strCtrl));
                        break;
                    case (int)CtrlType.PartialLinkText:
                        element = Driver.FindElement(By.PartialLinkText(strCtrl));
                        break;
                    case (int)CtrlType.TagName:
                        element = Driver.FindElement(By.TagName(strCtrl));
                        break;
                    case (int)CtrlType.XPath:
                        element = Driver.FindElement(By.XPath(strCtrl));
                        break;
                }

                switch (intCtrlClass)
                {
                    case (int)CtrlClass.Button:
                        element.Click();
                        break;
                    case (int)CtrlClass.SingleClick:
                        element.Click();
                        break;
                    case (int)CtrlClass.DoubleClick:
                        // First attemtp to click
                        bRtn = PerformAction(strCtrl, intCtrlType, (int)CtrlClass.SingleClick);
                        // Second try for clicking if PerformAction returned false
                        if (bRtn.Equals(false))
                        {
                            iWebRtn = WaitForElement(strCtrl, GlobalVariables.defaultWaitTime, (int)CtrlType.ID);
                            if (!iWebRtn.Equals(null))
                            {
                                bRtn = PerformAction(strCtrl, intCtrlType, (int)CtrlClass.SingleClick);
                                if (bRtn.Equals(false))
                                {
                                    bRtn = false;
                                }
                            }
                            else
                            {
                                bRtn = false;
                            }
                        }
                        break;
                    default:
                        element.Click();
                        try
                        {
                            element.Click();
                        }
                        catch (StaleElementReferenceException ex)
                        { }
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = PerformAction(" + strCtrl + ", " + intCtrlType + "," + intCtrlClass + ")");
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }

        /// <summary>
        /// PerformAction() method version 2
        /// </summary>
        /// <param name="strCtrl"></param>
        /// <param name="strValue"></param>
        /// <param name="intCtrlType"></param>
        /// <param name="intCtrlType"></param>
        /// <returns></returns>
        public bool PerformAction(string strCtrl, string strValue, int intCtrlType, int intCtrlClass)
        {
            var bRtn = true;
            IWebElement element = null;
            string strSelected = null;
            string strSet = null;
            SelectElement selectElement = null;
            IWebElement strIWebElement = null;

            try
            {
                switch (intCtrlType)
                {
                    case (int)CtrlType.ClassName:
                        element = Driver.FindElement(By.ClassName(strCtrl));
                        break;
                    case (int)CtrlType.CssSelector:
                        element = Driver.FindElement(By.CssSelector(strCtrl));
                        break;
                    case (int)CtrlType.ID:
                        element = Driver.FindElement(By.Id(strCtrl));
                        break;
                    case (int)CtrlType.LinkText:
                        element = Driver.FindElement(By.LinkText(strCtrl));
                        break;
                    case (int)CtrlType.Name:
                        element = Driver.FindElement(By.Name(strCtrl));
                        break;
                    case (int)CtrlType.PartialLinkText:
                        element = Driver.FindElement(By.PartialLinkText(strCtrl));
                        break;
                    case (int)CtrlType.TagName:
                        element = Driver.FindElement(By.TagName(strCtrl));
                        break;
                    case (int)CtrlType.XPath:
                        element = Driver.FindElement(By.XPath(strCtrl));
                        break;
                    //case (int)CtrlType.Text:
                    //    element = Driver.FindElement(By.Text(strCtrl));
                    //    break;
                }

                switch (intCtrlClass)
                {
                    case (int)CtrlClass.TextBox:
                        //Clear and Populate WebEditbox with passed strValue
                        element.Clear();
                        element.SendKeys(strValue);

                        //strSet = element.GetAttribute("Value");
                        //if (!strSet.Equals(null))
                        //{
                        //    strSet = strSet.Replace("\r\n", "\n");

                        //    //Validate if passed strValue was set
                        //    if (!strSet.Equals(strValue))
                        //    {
                        //        bRtn = false;
                        //    }
                        //}
                        break;
                    case (int)CtrlClass.DropDownList:
                        //Select DropDownList with passed strValue
                        selectElement = new SelectElement(element);
                        selectElement.SelectByText(strValue);

                        element = WaitForElement(strCtrl, GlobalVariables.defaultWaitTime, (int)CtrlClass.ID);
                        if (null == element)
                        {
                            bRtn = false;
                            return bRtn;
                        }
                        break;
                    default:
                        bRtn = false;
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = PerformAction(" + strCtrl + "," + strValue + "," + intCtrlType + "," + intCtrlClass + ")");
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }



        /// <summary>
        /// PerformAction() method version 3
        /// Override of PerformAction() method is used for elements(checkboxes, radio button), you can specify the state of your control 
        /// <example>
        /// <code>
        /// bRtn = PerformAction(element, (int)CheckStatus.Check);
        /// </code>>
        /// </example>
        /// </summary>
        /// <param name="element">Specify the element you want to click of type IWebElement </param>
        /// <param name="checkStatus">Pass the sate you want to control to be in (ex: Check, UnCheck) </param>
        /// <returns>bool true or false</returns>
        public bool PerformAction(IWebElement element, int checkStatus)
        {
            bool bRtn = true;

            try
            {
                //Return true if the state is what you expected
                if (checkStatus.Equals((int)CheckStatus.Check) && element.Selected.Equals(true))
                    return true;
                if (checkStatus.Equals((int)CheckStatus.UnCheck) && element.Selected.Equals(false))
                    return true;

                //Click the checkbox
                element.Click();

                //if isChecked is true and checkStatus is UnCheck
                if (checkStatus.Equals((int)CheckStatus.UnCheck) && element.Selected.Equals(true))
                {
                    //Click the checkbox
                    element.Click();

                    //If state still does not match
                    if (checkStatus.Equals((int)CheckStatus.UnCheck) && element.Selected.Equals(true))
                        return false;
                }

                //If isCheck is true and checkStatus is Checked
                if (checkStatus.Equals((int)CheckStatus.Check) && element.Selected.Equals(false))
                {
                    //Click the checkbox
                    element.Click();

                    //If state still does not match
                    if (checkStatus.Equals((int)CheckStatus.Check) && element.Selected.Equals(false))
                        return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = PerformAction(" + element + "," + checkStatus + ")");
                Log.Error(ex);
                return false;
            }
        }


        /// <summary>
        /// PerformAction() method version 4
        /// Override of PerformAction() method is used mainly for checkboxes and radio buttons, you can specify the state of your control
        /// <example> <code>
        /// bRtn = selObj.PerformAction("SomeContrlName", (int)CtrlType.ID, (int)CtrlClass.Button, (int)CheckStatus.Check);
        /// </code> </example>
        /// </summary>
        /// <param name="strCtrl">Control to Identify </param>
        /// <param name="intCtrlType">Control Type, see the CtrlType enumeration</param>
        /// <param name="intCtrlClass">Control Class, see the CtrlClass enumeration</param>
        /// <param name="checkStatus">Pass the state you want to control to be in (ex: Check, UnCheck)</param>
        /// <returns>bool true or false</returns>
        public bool PerformAction(string strCtrl, int intCtrlType, int intCtrlClass, int checkStatus)
        {
            bool bRtn = true;
            IWebElement element = null;

            try
            {
                switch (intCtrlType)
                {
                    case (int)CtrlType.ClassName:
                        element = Driver.FindElement(By.ClassName(strCtrl));
                        break;
                    case (int)CtrlType.CssSelector:
                        element = Driver.FindElement(By.CssSelector(strCtrl));
                        break;
                    case (int)CtrlType.ID:
                        element = Driver.FindElement(By.Id(strCtrl));
                        break;
                    case (int)CtrlType.LinkText:
                        element = Driver.FindElement(By.LinkText(strCtrl));
                        break;
                    case (int)CtrlType.Name:
                        element = Driver.FindElement(By.Name(strCtrl));
                        break;
                    case (int)CtrlType.PartialLinkText:
                        element = Driver.FindElement(By.PartialLinkText(strCtrl));
                        break;
                    case (int)CtrlType.TagName:
                        element = Driver.FindElement(By.TagName(strCtrl));
                        break;
                    case (int)CtrlType.XPath:
                        element = Driver.FindElement(By.XPath(strCtrl));
                        break;
                    //case (int)CtrlType.Text:
                    //    element = Driver.FindElement(By.Text(strCtrl));
                    //    break;
                }

                bRtn = PerformAction(element, checkStatus);

                if (bRtn.Equals(false))
                {
                    bRtn = false;
                    return bRtn;
                }

                return true;
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = PerformAction(" + strCtrl + "," + intCtrlType + "," + intCtrlClass + "," + checkStatus + ")");
                Log.Error(ex);
                return false;
            }
        }


        /// <summary>
        /// The getValue() method gets the value of control (it supports TextBox and DropDownList)
        /// <example>
        /// Example on how to use method
        /// <code>
        /// String returnedValue = "";
        /// bRtn = selObj.getValue("SomeControlName", ref returnedValue, (int)CtrlType.ID, (int) CtrlClass.TextBox);
        /// </code>
        /// </example>
        /// </summary>
        /// <param name="strCtrl">Control to identify</param>
        /// <param name="returnValue">Returns the String Value set for the control</param>
        /// <param name="intCtrlType">Control Type, see CtrlType enumeration</param>
        /// <param name="intCtrlClass">Contrl Class, see CtrlClass enumeration</param>
        /// <returns></returns>
        public bool getValue(string strCtrl, ref string returnValue, int intCtrlType, int intCtrlClass)
        {
            var bRtn = true;
            IWebElement element = null;
            string strSelected = null;
            string strSet = null;
            SelectElement selectElement = null;
            IWebElement strIWebElement = null;

            try
            {
                switch (intCtrlType)
                {
                    case (int)CtrlType.ClassName:
                        element = Driver.FindElement(By.ClassName(strCtrl));
                        break;
                    case (int)CtrlType.CssSelector:
                        element = Driver.FindElement(By.CssSelector(strCtrl));
                        break;
                    case (int)CtrlType.ID:
                        element = Driver.FindElement(By.Id(strCtrl));
                        break;
                    case (int)CtrlType.LinkText:
                        element = Driver.FindElement(By.LinkText(strCtrl));
                        break;
                    case (int)CtrlType.Name:
                        element = Driver.FindElement(By.Name(strCtrl));
                        break;
                    case (int)CtrlType.PartialLinkText:
                        element = Driver.FindElement(By.PartialLinkText(strCtrl));
                        break;
                    case (int)CtrlType.TagName:
                        element = Driver.FindElement(By.TagName(strCtrl));
                        break;
                    case (int)CtrlType.XPath:
                        element = Driver.FindElement(By.XPath(strCtrl));
                        break;
                    //case (int)CtrlType.Text:
                    //    element = Driver.FindElement(By.Text(strCtrl));
                    //    break;
                }

                switch (intCtrlClass)
                {
                    case (int)CtrlClass.TextBox:
                        returnValue = element.GetAttribute("Value");
                        break;

                    case (int)CtrlClass.DropDownList:
                        // Get the selected dropdown text
                        selectElement = new SelectElement(element);
                        strIWebElement = selectElement.SelectedOption;
                        returnValue = strIWebElement.Text;
                        break;

                    default:
                        bRtn = false;
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = getValue(" + strCtrl + "," + returnValue + "," + intCtrlType + "," + intCtrlClass + ")");
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="strCtrl"></param>
        /// <param name="rowNum"></param>
        /// <param name="colNum"></param>
        /// <param name="returnValue"></param>
        /// <returns></returns>
        public bool getValue(string strCtrl, int rowNum, int colNum, ref string returnValue)
        {
            bool bRtn = false;

            try
            {
                // Get all row in the grid
                var elems = Driver.FindElements(By.XPath("//table[@id='" + strCtrl + "']/tbody/tr"));

                // Make sure elems has the data
                if (elems.Count < rowNum + 1)
                    return false;

                // Get the row based on rowNum
                var rowElem = elems[rowNum];

                // Get all columns for the Row
                var colms = rowElem.FindElements(By.XPath("td"));

                // Verify if it has the columns we need to get the data for 
                if (colms.Count < colNum)
                    return false;

                // Save the value in return
                returnValue = colms[colNum - 1].Text;

                return true;
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = PerformAction(" + strCtrl + "," + rowNum + "," + colNum + "," + returnValue + ")");
                Log.Error(ex);
                bRtn = false;
            }
            return bRtn;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="strCtrl"></param>
        /// <param name="intCtrlType"></param>
        /// <returns></returns>
        public bool doseControlExist(string strCtrl, int intCtrlType)
        {
            bool bRtn = true;
            IWebElement element = null;

            try
            {
                switch (intCtrlType)
                {
                    case (int)CtrlType.ClassName:
                        element = Driver.FindElement(By.ClassName(strCtrl));
                        break;
                    case (int)CtrlType.CssSelector:
                        element = Driver.FindElement(By.CssSelector(strCtrl));
                        break;
                    case (int)CtrlType.ID:
                        element = Driver.FindElement(By.Id(strCtrl));
                        break;
                    case (int)CtrlType.LinkText:
                        element = Driver.FindElement(By.LinkText(strCtrl));
                        break;
                    case (int)CtrlType.Name:
                        element = Driver.FindElement(By.Name(strCtrl));
                        break;
                    case (int)CtrlType.PartialLinkText:
                        element = Driver.FindElement(By.PartialLinkText(strCtrl));
                        break;
                    case (int)CtrlType.TagName:
                        element = Driver.FindElement(By.TagName(strCtrl));
                        break;
                    case (int)CtrlType.XPath:
                        element = Driver.FindElement(By.XPath(strCtrl));
                        break;
                    //case (int)CtrlType.Text:
                    //    element = Driver.FindElement(By.Text(strCtrl));
                    //    break;
                }
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = doseControlExist(" + strCtrl + "," + intCtrlType + ")");
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="isDisabled"></param>
        /// <param name="strCtrl"></param>
        /// <param name="intCtrlType"></param>
        /// <param name="intCtrlClass"></param>
        /// <returns></returns>
        public bool isCtrlDisabled(ref bool isDisabled, string strCtrl, int intCtrlType, int intCtrlClass)
        {
            bool bRtn = true;
            isDisabled = true;

            try
            {
                IWebElement element = null;

                switch (intCtrlType)
                {
                    case (int)CtrlType.ClassName:
                        element = Driver.FindElement(By.ClassName(strCtrl));
                        break;
                    case (int)CtrlType.CssSelector:
                        element = Driver.FindElement(By.CssSelector(strCtrl));
                        break;
                    case (int)CtrlType.ID:
                        element = Driver.FindElement(By.Id(strCtrl));
                        break;
                    case (int)CtrlType.LinkText:
                        element = Driver.FindElement(By.LinkText(strCtrl));
                        break;
                    case (int)CtrlType.Name:
                        element = Driver.FindElement(By.Name(strCtrl));
                        break;
                    case (int)CtrlType.PartialLinkText:
                        element = Driver.FindElement(By.PartialLinkText(strCtrl));
                        break;
                    case (int)CtrlType.TagName:
                        element = Driver.FindElement(By.TagName(strCtrl));
                        break;
                    case (int)CtrlType.XPath:
                        element = Driver.FindElement(By.XPath(strCtrl));
                        break;
                    //case (int)CtrlType.Text:
                    //    element = Driver.FindElement(By.Text(strCtrl));
                    //    break;
                }

                if (element.Enabled.Equals(true))
                {
                    isDisabled = false;
                }
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = isCtrlDisabled(" + isDisabled + "," + strCtrl + "," + intCtrlType + "," + intCtrlClass + ")");
                Log.Error(ex);
                bRtn = false;

            }

            return bRtn;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="strlCtrl"></param>
        /// <param name="intCtrlType"></param>
        /// <param name="intCtrlClass"></param>
        /// <returns></returns>
        public bool isChecked(string strCtrl, int intCtrlType, bool isChecked)
        {
            bool bRtn = true;
            isChecked = false;

            try
            {
                IWebElement element = null;

                switch (intCtrlType)
                {
                    case (int)CtrlType.ClassName:
                        element = Driver.FindElement(By.ClassName(strCtrl));
                        break;
                    case (int)CtrlType.CssSelector:
                        element = Driver.FindElement(By.CssSelector(strCtrl));
                        break;
                    case (int)CtrlType.ID:
                        element = Driver.FindElement(By.Id(strCtrl));
                        break;
                    case (int)CtrlType.LinkText:
                        element = Driver.FindElement(By.LinkText(strCtrl));
                        break;
                    case (int)CtrlType.Name:
                        element = Driver.FindElement(By.Name(strCtrl));
                        break;
                    case (int)CtrlType.PartialLinkText:
                        element = Driver.FindElement(By.PartialLinkText(strCtrl));
                        break;
                    case (int)CtrlType.TagName:
                        element = Driver.FindElement(By.TagName(strCtrl));
                        break;
                    case (int)CtrlType.XPath:
                        element = Driver.FindElement(By.XPath(strCtrl));
                        break;
                    //case (int)CtrlType.Text:
                    //    element = Driver.FindElement(By.Text(strCtrl));
                    //    break;
                }

                if (element.Selected)
                {
                    isChecked = true;
                }
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = isChecked(" + strCtrl + "," + intCtrlType + "," + isChecked + ")");
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="strText"></param>
        /// <returns></returns>
        public bool VerifyText(string strText)
        {
            bool bResult = true;

            try
            {
                // Check if page contains strText
                bResult = Driver.PageSource.Contains(strText);
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = VerifyText(" + strText + ")");
                Log.Error(ex);
            }

            return bResult;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="strMainMenu"></param>
        /// <param name="strSubMenu"></param>
        /// <param name="intCtrlType"></param>
        /// <param name="intCtrlClass"></param>
        /// <returns></returns>
        public bool HoverMenu(string strMainMenu, string strSubMenu, int intMainType, int intSubType)
        {
            bool bResult = true;
            IWebElement menu = null;
            IWebElement menuOption = null;

            try
            {
                // Wait for Main Menu option to be available
                menu = WaitForElement(strMainMenu, GlobalVariables.defaultWaitTime, intMainType);
                if (menu == null)
                {
                    return bResult;
                }

                // Get the element that shows menu with the mouseOver event
                switch (intMainType)
                {
                    case (int)CtrlType.ClassName:
                        menu = Driver.FindElement(By.ClassName(strMainMenu));
                        break;
                    case (int)CtrlType.CssSelector:
                        menu = Driver.FindElement(By.CssSelector(strMainMenu));
                        break;
                    case (int)CtrlType.ID:
                        menu = Driver.FindElement(By.Id(strMainMenu));
                        break;
                    case (int)CtrlType.LinkText:
                        menu = Driver.FindElement(By.LinkText(strMainMenu));
                        break;
                    case (int)CtrlType.Name:
                        menu = Driver.FindElement(By.Name(strMainMenu));
                        break;
                    case (int)CtrlType.PartialLinkText:
                        menu = Driver.FindElement(By.PartialLinkText(strMainMenu));
                        break;
                    case (int)CtrlType.TagName:
                        menu = Driver.FindElement(By.TagName(strMainMenu));
                        break;
                    case (int)CtrlType.XPath:
                        menu = Driver.FindElement(By.XPath(strMainMenu));
                        break;
                    //case (int)CtrlType.Text:
                    //    element = Driver.FindElement(By.Text(strCtrl));
                    //    break;
                }

                // The element that I want to click (hidden)
                switch (intSubType)
                {
                    case (int)CtrlType.ClassName:
                        menuOption = Driver.FindElement(By.ClassName(strSubMenu));
                        break;
                    case (int)CtrlType.CssSelector:
                        menuOption = Driver.FindElement(By.CssSelector(strSubMenu));
                        break;
                    case (int)CtrlType.ID:
                        menuOption = Driver.FindElement(By.Id(strSubMenu));
                        break;
                    case (int)CtrlType.LinkText:
                        menuOption = Driver.FindElement(By.LinkText(strSubMenu));
                        break;
                    case (int)CtrlType.Name:
                        menuOption = Driver.FindElement(By.Name(strSubMenu));
                        break;
                    case (int)CtrlType.PartialLinkText:
                        menuOption = Driver.FindElement(By.PartialLinkText(strSubMenu));
                        break;
                    case (int)CtrlType.TagName:
                        menuOption = Driver.FindElement(By.TagName(strSubMenu));
                        break;
                    case (int)CtrlType.XPath:
                        menuOption = Driver.FindElement(By.XPath(strSubMenu));
                        break;
                    //case (int)CtrlType.Text:
                    //    element = Driver.FindElement(By.Text(strCtrl));
                    //    break;
                }

                // Build and perform the mouseOver with Advance User Interactions API
                OpenQA.Selenium.Interactions.Actions builder = new OpenQA.Selenium.Interactions.Actions(Driver);
                builder.MoveToElement(menu).Build().Perform();

                // Then click on when menu option is visible\
                menuOption.Click();

                // This is done to handle the single click and double click issue
                try
                {
                    menuOption.Click();
                }
                catch (StaleElementReferenceException ex)
                { }
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = HoverMenu(" + strMainMenu + ", " + strSubMenu + ", " + intMainType + ", " + intMainType + ", " + intSubType + ")");
                Log.Error(ex);
            }

            return bResult;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="strTableXPath"></param>
        /// <param name="intTextColm"></param>
        /// <param name="strColmText"></param>
        /// <param name="tableCtrl"></param>
        /// <param name="checkStatus"></param>
        /// <param name="selCtrl"></param>
        /// <returns></returns>
        public bool SelectGridItem(string strTableXPath, int intTextColm, string strColmText, string tableCtrl, int checkStatus, int selCtrl)
        {
            bool bRtn = false;
            IWebElement elemObj;
            int countElem = 0;
            bool isChecked = false;

            try
            {
                // Get all elements in the elems variable
                var elems = Driver.FindElements(By.XPath("//table[@id='" + strTableXPath + "']/tbody/tr"));
                // Wait for 10 second to make sure all the elements are saved in the variable
                System.Threading.Thread.Sleep(1000);
                // Looping through the elems variable box
                foreach (var rowElem in elems)
                {
                    var cells = rowElem.FindElements(By.XPath("td"));
                    System.Threading.Thread.Sleep(1000);

                    if (cells.Count > intTextColm)
                    {
                        if (cells[intTextColm].Text.Equals(strColmText))
                        {
                            if (selCtrl.Equals((int)TableGrid.Grid))
                            {
                                elemObj = cells[0].FindElement(By.TagName("a"));
                                if (elemObj.Selected)
                                {
                                    isChecked = true;
                                }
                                if (checkStatus.Equals((int)CheckStatus.Check) && isChecked.Equals(false))
                                {
                                    elemObj.Click();
                                    System.Threading.Thread.Sleep(1000);
                                }
                                if (checkStatus.Equals((int)CheckStatus.UnCheck) && isChecked.Equals(true))
                                {
                                    elemObj.Click();
                                    System.Threading.Thread.Sleep(1000);
                                }
                                if (checkStatus.Equals((int)CheckStatus.JustClick))
                                {
                                    elemObj.Click();
                                    System.Threading.Thread.Sleep(1000);
                                }
                            }
                            else
                            {
                                // In case of a table, strColmText must be a control Id, ex: ct100_PB_ManagerSelectionList_
                                // The countElement starts with 0 and is incremented through the table
                                elemObj = cells[0].FindElement(By.Id(tableCtrl + Convert.ToString(countElem)));
                                if (elemObj.Selected)
                                {
                                    isChecked = true;
                                }
                                if (checkStatus.Equals((int)CheckStatus.Check) && isChecked.Equals(false))
                                {
                                    elemObj.Click();
                                    System.Threading.Thread.Sleep(1000);
                                }
                                if (checkStatus.Equals((int)CheckStatus.UnCheck) && isChecked.Equals(true))
                                {
                                    elemObj.Click();
                                    System.Threading.Thread.Sleep(1000);
                                }
                                if (checkStatus.Equals((int)CheckStatus.JustClick))
                                {
                                    elemObj.Click();
                                    System.Threading.Thread.Sleep(1000);
                                }
                            }
                            bRtn = true;
                            break;

                            // The above break will stop code execution and the below code will not run.
                            // Currently testing a work around.

                            // This is done to handle some single click and double click issue
                            try
                            {
                                IWebElement iWeb = cells[0].FindElement(By.TagName("a"));
                                if (!iWeb.Equals(null))
                                {
                                    cells[0].FindElement(By.TagName("a")).Click();
                                    break;
                                }

                            }
                            catch (StaleElementReferenceException ex)
                            { }
                        }
                        countElem++;
                    }
                }
            }
            catch (Exception ex)
            {

                Log.Info("Method Params = SelectGridItem(" + strTableXPath + ", " + intTextColm + ", " + strColmText + ", " + tableCtrl + ", " + checkStatus + "," + selCtrl + ")");
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="iBtnPress"></param>
        /// <returns></returns>
        public bool ClickIEPopup(int iBtnPress)
        {
            bool bRtn = true;

            try
            {
                switch (iBtnPress)
                {
                    case (int)BtnClick.OK:
                        Driver.SwitchTo().Alert().Accept();
                        break;
                    case (int)BtnClick.Cancel:
                        Driver.SwitchTo().Alert().Dismiss();
                        break;
                    default:
                        bRtn = false;
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = ClickIEPopup(" + iBtnPress + ")");
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }


        /// <summary>
        /// CopyPopupURLAndClose() method copy's the URL of the popup window and returns it via the reference string
        /// <example>Example on how to call method </example>
        /// <code>
        /// selObj.CopyPopupURLAndClose("BankRoutingCodeConfirmationPage.aspx", ref outURL);
        /// </code>
        /// </summary>
        /// <param name="strPopupURLText"></param>
        /// <param name="strOutURL"></param>
        /// <returns></returns>
        public bool CopyPopupURLAndClose(string strPopupURLText, ref string strOutURL)
        {
            bool bRtn = true;

            try
            {
                foreach (InternetExplorer ie in new ShellWindows())
                {
                    if (ie.LocationURL.Contains(strPopupURLText))
                    {
                        strOutURL = ie.LocationURL;
                        ie.Quit();
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = copyPopupURLAndClose(" + strPopupURLText + "," + strOutURL + ")");
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileNAme"></param>
        /// <returns></returns>
        public bool TakeScreenShot(ref string fileNAme)
        {
            fileNAme = fileNAme + DateTime.Now.TimeOfDay.ToString().Replace(':', '_').Replace('.', '_') + ".jpg";

            try
            {
                // Declare screenshot object
                Screenshot screenShot = ((ITakesScreenshot)Driver).GetScreenshot();
                screenShot.SaveAsFile(fileNAme, System.Drawing.Imaging.ImageFormat.Jpeg); // take screenshot of current screen

                return true;
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = TakeScreenShot(" + fileNAme + ")");
                Log.Error(ex);
                return false;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool FileUpload(string filePath)
        {
            bool bRtn = true;

            try
            {
                // Gets the file element in the current webbrowser page and inserts given file pass and sends Enter key
                var fileObj = Driver.FindElement(By.XPath(GlobalVariables.iEBrowserFileTypeCtrl));
                fileObj.SendKeys(filePath);
            }
            catch (Exception ex)
            {
                Log.Info("Method Params = FileUpload(" + filePath + ")");
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }




        /// <summary>
        /// 
        /// </summary>
        /// <param name="folderLocation"></param>
        /// <param name="fileName"></param>
        /// <param name="Ldata"></param>
        /// <returns></returns>
        public bool CreateFile(string folderLocation, string fileName, List<string> Ldata)
        {
            string filePath = folderLocation + @"\" + fileName;

            try
            {
                // Create Folder if do not exist
                if (!System.IO.Directory.Exists(folderLocation))
                    System.IO.Directory.CreateDirectory(folderLocation);

                // Initialise stream object with file
                using (System.IO.StreamWriter wr = new System.IO.StreamWriter(filePath, true))
                {
                    foreach (string value in Ldata)
                    {
                        var sb = new StringBuilder();
                        sb.Append(value);
                        wr.WriteLine(sb.ToString());
                    }

                    wr.Close();
                }

                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                return false;
            }
        }




        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool DeleteFile(string filePath)
        {
            bool bRtn = true;

            try
            {
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                    return bRtn;
                }

            }
            catch (Exception ex)
            {
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }



    }



        public static class ExtensionMethods
        {
            public static IWebElement FindElementOnPage(this IWebDriver webDriver, By by)
            {
                OpenQA.Selenium.Remote.RemoteWebElement element = (OpenQA.Selenium.Remote.RemoteWebElement)webDriver.FindElement(by);
                var hack = element.LocationOnScreenOnceScrolledIntoView;
                return element;
            }

        }
    
}
