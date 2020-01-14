using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports.Reporter.Configuration;
using System.IO;
using System.Net;

namespace SeleniumWorker
{
    [TestClass]
    public class Adressaenderung
    {
        #region TestContext
        public TestContext TestContext { get; set; }
        private SeleniumWorker.SelWorker _mySelenium;
        private readonly static ExtentReports Extent = new ExtentReports();
        private static ExtentHtmlReporter _htmlReporter;
        #endregion

        [ClassInitialize]
        public static void InitTestClass(TestContext testContext)
        {
            var programDirectory = AppDomain.CurrentDomain.BaseDirectory;
            programDirectory = Path.GetFullPath(Path.Combine(programDirectory, @"..\..\..\"));
            var reportDirectory = Path.Combine(programDirectory + "Report");

            if (File.Exists(reportDirectory) == false)
            {
                _ = Directory.CreateDirectory(reportDirectory);
            }

            _htmlReporter = new ExtentHtmlReporter(reportDirectory + @"\P_AdressaenderungDetailed.html");
            Extent.AttachReporter(_htmlReporter);
            Extent.AddSystemInfo("Operationg System: ", "Windows Server 2016");
            var hostname = Dns.GetHostName();
            Extent.AddSystemInfo("Hostname: ", hostname);
            Extent.AddSystemInfo("Browser: ", "Chrome");
        }

        [TestMethod]
        [TestCategory("Functional")]
        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV",
                    "|DataDirectory|\\testdata\\P_AdressaenderungDetailed.csv",
                    "P_AdressaenderungDetailed#csv",
                    DataAccessMethod.Sequential),
                    DeploymentItem("testdata\\P_AdressaenderungDetailed.csv")]
        [DeploymentItem("avcodec-57.dll")]
        [DeploymentItem("avdevice-57.dll")]
        [DeploymentItem("avfilter-6.dll")]
        [DeploymentItem("avformat-57.dll")]
        [DeploymentItem("avutil-55.dll")]
        [DeploymentItem("postproc-54.dll")]
        [DeploymentItem("RazorEngine.dll")]
        [DeploymentItem("swresample-2.dll")]
        [DeploymentItem("swscale-4.dll")] 
        public void P_AdressaenderungDetailed()
        {
            
                                 
            var test = Extent.CreateTest(TestContext.DataRow["TestCaseName"].ToString(), "This test is about testing usecase Adressänderung within DynamicBrain regression testing");



            test.Log(Status.Info, "Step1: Test case has started");

            _mySelenium = new SeleniumWorker.SelWorker(TestContext.DataRow["TestCaseName"].ToString());
            if (null != _mySelenium)
            {

                test.Log(Status.Pass, "Step2: Selenium instance could by serialized");
            }
            _mySelenium.ChangeLogFileName("RollingLogFileAppender", TestContext.DataRow["TestCaseName"].ToString());
            // Check weather to send task to Rest
            if (TestContext.DataRow["CreateBorObject"].ToString() == "true")
            {
                var result = SelWorker.SendRequestToRest(DateTime.Now,
                                                           TestContext.DataRow["Category"].ToString(),
                                                           TestContext.DataRow["Location"].ToString(),
                                                           TestContext.DataRow["Channel"].ToString(),
                                                           TestContext.DataRow["Rescan"].ToString(),
                                                           TestContext.DataRow["Sender"].ToString()).Result;

                test.Log(Status.Pass, "Step3: Task (Bor-Object) has been created: " + TestContext.DataRow["Category"].ToString() + " ContactID: " + result.ToString());
                Assert.IsTrue(result != null);
            }
            _mySelenium.Init(new Uri(TestContext.DataRow["NodeURL"].ToString()), new Uri(TestContext.DataRow["URLTestObject"].ToString()), TestContext.DataRow["Browser"].ToString());
            test.Log(Status.Pass, "Step4: Selenium-Initialization has been performed" );
            var browserTitle = _mySelenium.GetBrowserTitle("Voxtron Web Client");
            Assert.IsTrue(browserTitle);

            var ScreenShotImagePath1 = SelWorker.GetScreenShot(ref _mySelenium._driver);
            test.Log(Status.Info, "Snapshot below: " + test.AddScreenCaptureFromPath(ScreenShotImagePath1, "CheckBrowserName"));

            test.Log(Status.Pass, "Step5: Browser Title: Voxtron Web Client: " + browserTitle);
            _mySelenium.Login(TestContext.DataRow["UserId"].ToString(), TestContext.DataRow["Password"].ToString(), TestContext.DataRow["VoxWebCltStatus"].ToString());
            test.Log(Status.Pass, "Step6: Login has been performed!");
            test.Log(Status.Info, "User:          " + TestContext.DataRow["UserId"].ToString());
            test.Log(Status.Info, "Password:      " + TestContext.DataRow["Password"].ToString());
            test.Log(Status.Info, "Client status: " + TestContext.DataRow["VoxWebCltStatus"].ToString());

            test.Log(Status.Info, "Snapshot below: " + test.AddScreenCaptureFromPath(SelWorker.GetScreenShot(ref _mySelenium._driver), "CheckStatusAfterLogin"));

            _mySelenium.AcceptTask(1, int.Parse(TestContext.DataRow["ContactCodeItemIdx"].ToString()));
            test.Log(Status.Pass, "Step 7: Task has been accpeted");
            _mySelenium.Logout();
            test.Log(Status.Info, "Snapshot below: " + test.AddScreenCaptureFromPath(SelWorker.GetScreenShot(ref _mySelenium._driver), "CheckStatusAfterLogout"));
            test.Log(Status.Pass, "Step 8: Logout has been performed");
            _mySelenium.Terminate();
            test.Log(Status.Pass, "Step 9: Selenium Terminated successful");



            //Create video in another thread
            var InstanceCaller = new Thread(new ThreadStart(() => _mySelenium.CreateVideo()));
            InstanceCaller.Start();

            test.Log(Status.Pass, "Steph 10: Video has been created");

            test.Log(Status.Pass, "Test case has successfully passed");

        }

        [ClassCleanup()]
        public static void Cleanup()
        {
            Extent.Flush();
        }
    }
}
