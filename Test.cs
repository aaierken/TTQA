using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
//using OpenQA.Selenium.IE;

namespace GlobalLibrary
{
    class Test
    {
        bool bRtn = true;
        IWebElement iWeRtn;
        public static GlobalEngine ReportEngine = new GlobalEngine();
        public static GlobalMethods GlobalMethods = new GlobalMethods();

        static void Main(string[] args)
        {
            //Set up
            var ToAddress = "mnm703@gmail.com";
            var Subject = "Selenium Automation Auto Generated";
            var Body = "Please open the attached Selenium Automation Result Report with IE 9 or lower version!\n\n\n If you have question, please contact with Frank @ 703-944-2969. \n\n Thanks, \nSelenium - QA Automation Team";
            var AttachmentPath = @"C:\ReportTest.htm";
            var ReportName = "ReportTest.htm";

            SimpleScript_Test();
            System.Threading.Thread.Sleep(3000);
            ReportEngine.GenerateHtmlReport(ReportName);

            GlobalMethods.SendEmail(ToAddress, Subject, Body, AttachmentPath);


                        
            //LoggerTest();
                        
            //OpenBrowserMethod_Test();

            //ScreenResize_Test();


            
        }


        public static void LoggerTest()
        {
            Log.Info("This is testing for 'Info' logging");
            Log.Error("This is testing for 'Error' logging");
            Log.Debug("This is testing for 'Debug' logging");
        }

        public static void OpenBrowserMethod_Test()
        {
            //instantiating GlobalSeleniumMethods Class
            GlobalSeleniumMethods selObj = new GlobalSeleniumMethods();

            //Opens Browser
            selObj.OpenBrowser("www.google.com", "gbqfba", 10, (int)CtrlClass.Button);

            //Close the browser
            //selObj.CloseAllIEBrowser();

            

            selObj.Quit();

        }

        public static void ScreenResize_Test()
        {
            GlobalSeleniumMethods selObj = new GlobalSeleniumMethods();
            selObj.ScreenResize(1024, 680);
        }

        public static void SimpleScript_Test()
        {

            GlobalSeleniumMethods selObj = new GlobalSeleniumMethods();            
            bool bRtn;                      
            int n =1;
            string storyTitle = "Sample TestCase Automation with Selenium Framework in C#";
            //string storyTitle2 = "StoryTile 2 Testing creating new section";
            Log.Info("***********************************************" + storyTitle + " ========> Started *****************************************************************");
            

            try
            {

                //Opens Browser
                bRtn = selObj.OpenBrowser("www.quickenloans.com", "//*[@id='navigation_primary_home']", 5, (int)CtrlType.XPath);
                ReportEngine.AddToReport(n++.ToString(), "Goto 'www.quickenloans.com' site", bRtn, "", storyTitle);
                
                bRtn = selObj.HoverMenu("navigation_primary_refinance","//a[contains(text(),'Lower Your Payment')]", (int)CtrlType.ID, (int)CtrlType.XPath);
                ReportEngine.AddToReport(n++.ToString(), "Goto 'Refinance>Lower Your Payment' page ", bRtn, "", storyTitle);

                
                //bRtn = selObj.PerformAction("//a[contains(text(),'Home')]",(int)CtrlType.XPath, (int)CtrlClass.SingleClick);
                //ReportEngine.AddToReport(n++.ToString(), "Click on 'Home' tab ", bRtn, "", storyTitle);

                //bRtn = selObj.HoverMenu("//a[contains(text(),'Refinance')]", "//a[contains(text(),'Lower Your Payment')]", (int)CtrlType.XPath, (int)CtrlType.XPath);
                //ReportEngine.AddToReport(n++.ToString(), "Goto 'Refinance>Lower Your Payment' page", bRtn, "", storyTitle);

                bRtn = selObj.PerformAction("//*[@id='referral_FirstName']","FirstName Frank", (int)CtrlType.XPath, (int)CtrlClass.TextBox);
                ReportEngine.AddToReport(n++.ToString(), "Enters 'Frank' in the FirstName field ", bRtn, "", storyTitle);

                bRtn = selObj.PerformAction("referral_LastName", "LastName Musabay", (int)CtrlType.ID, (int)CtrlClass.TextBox);
                ReportEngine.AddToReport(n++.ToString(), "Enters 'Musabay' in the LastName field ", bRtn, "", storyTitle);

                bRtn = selObj.PerformAction("referral[Zipcode]", "92126", (int)CtrlType.Name, (int)CtrlClass.TextBox);
                ReportEngine.AddToReport(n++.ToString(), "Enters '92126' value in the Zip: field ", bRtn, "", storyTitle);

            }
            catch(Exception ex)
            {
                Log.Error(ex);
                bRtn = false;                
                selObj.Quit();
            }

            selObj.Quit();
            Log.Info("***********************************************" + storyTitle + " ========> Ended *****************************************************************");

        }





    }
}
