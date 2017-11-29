using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HTMLReportEngine;
using OpenQA.Selenium;
using System.Diagnostics;

namespace GlobalLibrary
{
    class GlobalEngine
    {
        private string firstSectionStoryID;
        private string otherSectionStoryID;
        private string firstStoryTitleID="";
        private int statusValidationPass;
        private int statusValidationFailed;
        //static Section project = new Section("StoryTitle", "StoryID: ");

        private DataTable myTable;
        private Report report;
        private Exception myException;
        private DataSet myDataSet;

        private string FirstSectionStoryID
        {
            get { return firstSectionStoryID; }
            set { firstSectionStoryID = value; }
        }

        private string OtherSectionStoryID
        {
            get { return otherSectionStoryID; }
            set { otherSectionStoryID = value; }
        }

        private string FirstStoryTitleID
        {
            get { return firstStoryTitleID; }
            set { firstStoryTitleID = value; }
        }

        private int StatusValidationPass
        {
            get { return statusValidationPass; }
            set { statusValidationPass = value; }
        }

        private int StatusValidationFailed
        {
            get { return statusValidationFailed; }
            set { statusValidationFailed = value; }
        }

        /// <summary>
        /// ReportEngine() method is used to add steps 
        /// </summary>
        public GlobalEngine()
        {
            myTable = new DataTable();
            report = new Report();
            myDataSet = new DataSet();
            myException = new Exception();
        }


        private void addReportColumns()
        {
            if (myTable.Columns.Count.Equals(0))
            {
                var StoryExecutiontimeStart = DateTime.Now.ToString();

                myTable.Columns.Add("Step #");
                myTable.Columns.Add("Step Ended");
                myTable.Columns.Add("Step Name");
                myTable.Columns.Add("Validation");
                myTable.Columns.Add("Error");
                myTable.Columns.Add("Pass");
                myTable.Columns.Add("Failed");
                myTable.Columns.Add("StoryTitle");
            }
        }

        public void AddToReport(string iStepNumber, string sStepName, bool sValidation, string sException, string sStoryTitle)
        {
            var StepExecutionTime = DateTime.Now.ToString();
            var countStepNumer = iStepNumber;

            if(!sException.Equals(null) || !sException.Equals(""))
            {
                //var sExceptionResult = sException.ToString();
            }
            else
            {
                sException = "";
            }

            if(sValidation.Equals(true))
            {
                var sValidationResult = "Passed";
                statusValidationPass = 1;
                statusValidationFailed = 0;
                addReportColumns();
                iStepNumber = "Step #" + iStepNumber;

                var myRowCount = myTable.Rows.Count;
                if (myRowCount.Equals(0))
                {
                    firstStoryTitleID = sStoryTitle;
                    myTable.Rows.Add(iStepNumber, StepExecutionTime, sStepName, sValidationResult, sException, statusValidationPass, StatusValidationFailed, firstStoryTitleID);
                    //myTable.Rows.[0].AcceptChanges();
                    CreateSection();
                }
                else
                {
                    firstStoryTitleID = sStoryTitle;
                    var emptyRow = myTable.Rows.Count +1;
                    var myNewRow = myTable.Rows.Add(iStepNumber, StepExecutionTime, sStepName, sValidationResult, sException, statusValidationPass, StatusValidationFailed, sStoryTitle);
                }
            }
            else
            {   
                GlobalSeleniumMethods selObj = new GlobalSeleniumMethods();
                IWebDriver webdriver = null;
                string ScreenShot ="";
                webdriver = selObj.GetDriver();
                selObj.TakeScreenShot(ref ScreenShot);
                string ScreenShotLink = "<a href=\"" + ScreenShot + "\">Error_Snapshot</a" ;
                sException = ScreenShotLink;

                var sValidationResult = "Failed";
                statusValidationPass = 0;
                statusValidationFailed = 1;
                addReportColumns();
                iStepNumber = "Step #" + iStepNumber;

                var myRowCount = myTable.Rows.Count;
                if (myRowCount.Equals(0))
                {
                    firstSectionStoryID = sStoryTitle;
                    myTable.Rows.Add(iStepNumber, StepExecutionTime, sStepName, sValidationResult, sException, statusValidationPass, StatusValidationFailed, firstSectionStoryID);
                    //myTable.Rows.[0].AcceptChanges();
                    CreateSection();
                }
                else
                {
                    firstStoryTitleID = sStoryTitle;
                    var emptyRow = myTable.Rows.Count +1;
                    var myNewRow = myTable.Rows.Add(iStepNumber, StepExecutionTime, sStepName, sValidationResult, sException, statusValidationPass, StatusValidationFailed, firstStoryTitleID);
                }
                //if (!iStepNumber.Equals("Error"))
                //{
                //    AssertStatement(sValidationFalse,sStepName);
                //}
            }
        }

        public void AddToReport(string iStepNumber, string sStepName, IWebElement IWebElement, string sException, string sStoryTitle)
        {
            if(IWebElement.Equals(null))
            {
                AddToReport(iStepNumber,sStepName,false,sException,sStoryTitle);
            }
            else
            {
                AddToReport(iStepNumber,sStepName,true,sException,sStoryTitle);
            }
        }

        private void CreateSection()
        {
            Section project = new Section("StoryTitle", "Story: ");
            report.Sections.Add(project);
            project.GradientBackground = true;
            project.BackColor = System.Drawing.Color.Gray;

            // Add Total Fields
            int TotalPassed = report.TotalFields.Add("Pass");
            int TotolFailed = report.TotalFields.Add("Failed");

            project.IncludeChart = true;
            project.ChartShowBorder = true;

            // Set chart properties
            project.ChartChangeOnField = "Pass";
            // project.ChartValueField = "Pass";
            project.ChartChangeOnField = "Failed";
        }

        private void AssertStatement(bool validation, string description)
        {
            if (validation.Equals(false))
            {
                Assert.IsTrue(validation, description + "--- Failed");
            }
            
        }

        public void GenerateHtmlReport(string reportName)
        {
            var ReportDate = DateTime.Now.ToString();

            report.ReportFields.Add(new Field("Step #", "Step #", 70, ALIGN.LEFT));
            report.ReportFields.Add(new Field("Step Ended", "Step Ended", 140, ALIGN.LEFT));
            report.ReportFields.Add(new Field("Step Name", "Step Name", 0, ALIGN.LEFT));
            report.ReportFields.Add(new Field("Validation", "Validation",100,ALIGN.CENTER));
            report.ReportFields.Add(new Field("Error", "Error",100,ALIGN.CENTER));
            report.ReportFields.Add(new Field("Pass", "Pass",50,ALIGN.LEFT));
            report.ReportFields.Add(new Field("Failed", "Failed",50,ALIGN.LEFT));
            report.ReportFields.Add(new Field("StoryTitle", "StoryTitle",400, ALIGN.LEFT));
            
            // Create a DataSet and fill it before using this code
            myDataSet.Tables.Add(myTable);
            report.ReportSource = myDataSet;
            bool bRtn;

            report.ReportTitle = "Selenium Automation Result Report on: " + ReportDate;
            try
            {
                bRtn = report.SaveReport(@"c:\ReportTest.htm");
            }
            catch (Exception ex)
            { }
        }

        public TimeSpan MeasureExecTime(Action action)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            action();
            TimeSpan stepExecutionTime = stopwatch.Elapsed;
            return stepExecutionTime;
        }




                
       

    }
}
