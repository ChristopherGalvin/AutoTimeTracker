using System.Configuration;
using System.Collections.Specialized;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;

namespace Project_PeopleHRLogin
{
    class Program
    {

        public static CurrentRuntime.TimeType? currentTimeType;


        static void Main(string[] args)
        {


            TimeSpan startIn = new TimeSpan(8, 0, 0); 
            TimeSpan endIn = new TimeSpan(10, 0, 0);

            TimeSpan startOut = new TimeSpan(17, 0, 0);
            TimeSpan endOut = new TimeSpan(18, 0, 0);



            TimeSpan now = DateTime.Now.TimeOfDay;

            if ((now > startIn) && (now < endIn))
            {
                currentTimeType = CurrentRuntime.TimeType.TimeIn;
            }
            else if ((now > startOut) && (now < endOut))
            {
                currentTimeType = CurrentRuntime.TimeType.TimeOut;
            }


            currentTimeType = CurrentRuntime.TimeType.TimeIn;

            if (currentTimeType != null)
            {
                CurrentRuntime currentRunTime = new CurrentRuntime(currentTimeType);

                if (currentRunTime.Run)
                {
                    var chromeOptions = new ChromeOptions();
                    //chromeOptions.AddArguments("headless");
                    using (var browser = new ChromeDriver(chromeOptions))
                    {

                        browser.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);


                        // Log into the PeopleHR system
                        browser.Url = ConfigurationManager.AppSettings["URL"];
                        browser.FindElementById("txtEmailId").SendKeys(ConfigurationManager.AppSettings["Username"]);
                        browser.FindElementById("txtPassword").SendKeys(ConfigurationManager.AppSettings["Password"]);
                        browser.FindElementById("btnLogin").Click();


                        //Navigate to the details page
                        browser.Url = ConfigurationManager.AppSettings["URL"] + "Pages/LeftSegment/MyDetails.aspx";

                        Thread.Sleep(2000);

                        browser.ExecuteScript("document.querySelector('[data-date=\"" + (DateTime.Now).ToString("yyyy-MM-dd") + "\"]').click();");


                        try
                        {
                            if (!currentRunTime.AtHome)
                            {
                                // Write date in and out

                                //If its run in afternoon and still not filled out, then...
                                string currentTimeInValue = browser.FindElementById("txtTimeInHH1").GetAttribute("value");
                                string logText = "";


                                if (currentTimeType == CurrentRuntime.TimeType.TimeIn || currentTimeInValue == "HH")
                                {
                                    browser.FindElementById("txtTimeInHH1").Clear();
                                    browser.FindElementById("txtTimeInHH1").SendKeys(currentRunTime.InHour);
                                    browser.FindElementById("txtTimeInMM1").Clear();
                                    browser.FindElementById("txtTimeInMM1").SendKeys(currentRunTime.InMinute);

                                    logText +=  "(IN: " + currentRunTime.InHour + ":" + currentRunTime.InMinute + ")";

                                }



                                if (currentTimeType == CurrentRuntime.TimeType.TimeOut)
                                {
                                    browser.FindElementById("txtTimeOutHH1").Clear();
                                    browser.FindElementById("txtTimeOutHH1").SendKeys(currentRunTime.OutHour);
                                    browser.FindElementById("txtTimeOutMM1").Clear();
                                    browser.FindElementById("txtTimeOutMM1").SendKeys(currentRunTime.OutMinute);
                                    logText += "(OUT: " + currentRunTime.OutHour + ":" + currentRunTime.OutMinute + ")";
                                }

                                currentRunTime.WriteToTextFile(logText, false);

                                browser.FindElementById("aSave").Click();
                            }
                            else
                            {

                                string logText = "";

                                if (currentTimeType == CurrentRuntime.TimeType.TimeIn)
                                {
                                    browser.FindElementById("aProjectTimesheet").Click();
                                    browser.FindElementById("ProjectTimesheetList_aAddNew").Click();
                                    browser.FindElementById("ddlProject_ddlManagedList");

                                    // DropDowns
                                    var projectSelectElement = browser.FindElementById("ddlProject_ddlManagedList");
                                    var projectSelect = new SelectElement(projectSelectElement);
                                    projectSelect.SelectByText("Work from home");

                                    Thread.Sleep(100);


                                    var taskSelectElement = browser.FindElementById("ddlProjectTask_ddlManagedList");
                                    var taskSelect = new SelectElement(taskSelectElement);
                                    taskSelect.SelectByText("Technical support");


                                    browser.FindElementById("txtStartTimeHH").Clear();
                                    browser.FindElementById("txtStartTimeHH").SendKeys("08");
                                    browser.FindElementById("txtStartTimeMM").Clear();
                                    browser.FindElementById("txtStartTimeMM").SendKeys("45");
                                    browser.FindElementById("txtEndTimeHH").Clear();
                                    browser.FindElementById("txtEndTimeHH").SendKeys("17");
                                    browser.FindElementById("txtEndTimeMM").Clear();
                                    browser.FindElementById("txtEndTimeMM").SendKeys("15");

                                    browser.FindElementById("aProjectTimesheetSave").Click();


                                    logText +=  "(IN: " + currentRunTime.InHour + ":" + currentRunTime.InMinute + ")";
                                    logText +=  "(OUT: " + currentRunTime.OutHour + ":" + currentRunTime.OutMinute + ")";


                                    currentRunTime.WriteToTextFile(logText + " - AT HOME", false);
                                }

                            }


                            Thread.Sleep(5000);
                        }
                        catch (Exception e)
                        {
                            string logText = e.ToString();
                            currentRunTime.WriteToTextFile(logText, true);
                        }


                    }
                    
                }

            }

        }

        private static string RandomTime(int startAt, int endAt)
        {
            Random rand = new Random();
            int returnValue = rand.Next(startAt, endAt);
            return returnValue.ToString();
        }


       
        
    }


    public class CurrentRuntime
    {
        public string Day;
        public bool AtHome;
        public string InHour;
        public string InMinute;
        public string OutHour;
        public string OutMinute;
        public bool Run = false;
        private TimeType? _currentType;


        public CurrentRuntime(TimeType? currentType)
        {
            Day = DateTime.Now.DayOfWeek.ToString();
            if (ConfigurationManager.AppSettings["DaysHome"].Contains(Day.Substring(0, 3)))
            {
                AtHome = true;
            }

            if(DateTime.Now.DayOfWeek != DayOfWeek.Saturday && DateTime.Now.DayOfWeek != DayOfWeek.Sunday)
            {
                Run = true;
            }


            InHour = "08";
            InMinute = RandomTime(20, 44);

            OutHour = "17";
            OutMinute = RandomTime(15, 30);

            _currentType = currentType;


        }

        private static string RandomTime(int startAt, int endAt)
        {
            Random rand = new Random();
            int returnValue = rand.Next(startAt, endAt);
            return returnValue.ToString();
        }


        public void WriteToTextFile(string text, bool isError)
        {
            using (var strW = new StreamWriter("log.txt", true))
            {
                if (!isError)
                {
                    strW.WriteLine("(" + DateTime.Now.ToString() + ") SUCCESS - " + text);
                }
                else
                {
                    strW.WriteLine("(" + DateTime.Now.ToString() + ") ERROR - " + text);
                }
                
            }
        }

        public enum TimeType
        {
            TimeIn = 0,
            TimeOut = 1
        }

    }


    public class Selectors
    {
        public static By SelectorByAttributeValue(string p_strAttributeName, string p_strAttributeValue)
        {
            return (By.XPath(String.Format("//*[@{0} = '{1}']",
                                           p_strAttributeName,
                                           p_strAttributeValue)));
        }
    }


}
