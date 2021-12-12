using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using ClosedXML.Excel;
using NUnit.Framework;

namespace SightlyTestProject
{
    public class Tests
    {
        IWebDriver driver;
        string url = "https://staging-newtargetview.sightly.com/";
        string masterFile = "../../../ComparisonFiles/PerformanceDetail-AdWords-Ad_Group_Mapping-Master.xlsx";
        string fileDownLoad = @"C:\Downloads\Tests";
        string user = "qa-tester@qa-test.com";
        string password = "sightlyqatest";

        
        [SetUp]
        public void Setup()
        {
            ChromeOptions chromeOptions = new ChromeOptions();
            chromeOptions.AddUserProfilePreference("download.default_directory", fileDownLoad);

            driver = new ChromeDriver(chromeOptions);
            driver.Manage().Window.Maximize();
            driver.Manage().Cookies.DeleteAllCookies();
        }

        [Test]
        public void Test1()
        {
            // Navigate to the url
            driver.Navigate().GoToUrl(url);

            // Set an Explicit Wait of 10 seconds. This will be used to wait for an element to be found before attempting to click on it.
            WebDriverWait wait10 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            // Log into the website using the provided credentials
            driver.FindElement(By.XPath("//input[@type=\'text\']")).SendKeys(user);
            driver.FindElement(By.XPath("//input[@type=\'password\']")).SendKeys(password);
            driver.FindElement(By.ClassName("login-button")).Click();

            // Click on the Reports icon
            wait10.Until(driver => driver.FindElement(By.Id("header-reports")));
            driver.FindElement(By.Id("header-reports")).Click();

            // Click on All to Display all reports; Added a 500ms sleep as neither an implicit wait or explicit wait would allow for the XPath to be found
            Thread.Sleep(500);
            driver.FindElement(By.XPath("//div[3]/a")).Click();

            // Click on the checkbox for Ad Group Mapping - TEST IOG
            wait10.Until(driver => driver.FindElement(By.XPath("(//input[@value='[object Object]'])[2]")));
            driver.FindElement(By.XPath("(//input[@value='[object Object]'])[2]")).Click();

            // Click on the Create Report Button
            wait10.Until(driver => driver.FindElement(By.ClassName("create-report-text")));
            driver.FindElement(By.ClassName("create-report-text")).Click();

            // Click the Performance Detail Radial button
            wait10.Until(driver => driver.FindElement(By.XPath("//input[@value='detail']")));
            driver.FindElement(By.XPath("//input[@value='detail']")).Click();

            // Select the Report Options 
            driver.FindElement(By.CssSelector(".granularitySelect")).SendKeys("Daily" + Keys.Enter);
            driver.FindElement(By.CssSelector(".timePeriodSelect")).SendKeys("All Time" + Keys.Enter);

            // Run the Report
            driver.FindElement(By.ClassName("run-report-button")).Click();

            // Wait for file to download
            Thread.Sleep(2000);

            // Get the file name of the most recent file in the download folder
            var files = new DirectoryInfo(fileDownLoad).GetFiles("*.*");
            string latestfile = "";

            DateTime lastupdated = DateTime.MinValue;

            foreach (FileInfo file in files)
            {
                if (file.LastWriteTime > lastupdated)
                {
                    lastupdated = file.LastWriteTime;
                    latestfile = file.Name;
                }
            }

            string expectedFilePath = Path.Combine(fileDownLoad, latestfile);

            // Get the file path for the Master File
            string masterFilePath = Path.GetFullPath(masterFile);
            Console.Write(masterFilePath);
            
            // Create a list of the data in the downloaded file
            var list1 = new List<string>();
            using (var wbook1 = new XLWorkbook(expectedFilePath))
            {
                var ws1 = wbook1.Worksheet(1);
                list1 = (List<string>)ws1.Range("A5:S100")
                    .CellsUsed()
                    .Select(c => c.Value.ToString())
                    .ToList();
            }

            // Create a list for the data in the master file for only the column names and the corresponding data
            var list2 = new List<string>();
            using (var wbook2 = new XLWorkbook(masterFilePath))
            {
                var ws2 = wbook2.Worksheet(1);
                list2 = (List<string>)ws2.Range("A5:S100")
                    .CellsUsed()
                    .Select(c => c.Value.ToString())
                    .ToList();
            }

            // Check that the data in the lists are equal and in the proper sequence. Delete downloaded report if True.
            bool check = list1.SequenceEqual(list2);
            if (check is true)
            {
                File.Delete(expectedFilePath);
            }
            Assert.IsTrue(check);
                       
        }
        [TearDown]
        public void tear_Down()
        {
            driver.Quit();
        }
    }
}