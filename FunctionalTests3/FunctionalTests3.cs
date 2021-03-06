using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SeleniumWorker
{
    [TestClass]
    public class FunctionalTests3
    {
        #region TestContext
        public TestContext TestContext { get; set; }
        private SeleniumWorker.SelWorker _mySelenium;
        #endregion

        [TestMethod]
        [TestCategory("Functional")]
        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV",
                    "|DataDirectory|\\testdata\\P_Functional3.csv",
                    "P_Functional3#csv",
                    DataAccessMethod.Sequential),
                    DeploymentItem("testdata\\P_Functional3.csv")]
        [DeploymentItem("avcodec-57.dll")]
        [DeploymentItem("avdevice-57.dll")]
        [DeploymentItem("avfilter-6.dll")]
        [DeploymentItem("avformat-57.dll")]
        [DeploymentItem("avutil-55.dll")]
        [DeploymentItem("postproc-54.dll")]
        [DeploymentItem("swresample-2.dll")]
        [DeploymentItem("swscale-4.dll")]
        public void P_Functional3()
        {
            _mySelenium = new SeleniumWorker.SelWorker(TestContext.DataRow["TestCaseName"].ToString());
            _mySelenium.ChangeLogFileName("RollingLogFileAppender", TestContext.DataRow["TestCaseName"].ToString());
            if (TestContext.DataRow["CreateBorObject"].ToString() == "true")
            {
                var result = SelWorker.SendRequestToRest(DateTime.Now,
                                                           TestContext.DataRow["Category"].ToString(),
                                                           TestContext.DataRow["Location"].ToString(),
                                                           TestContext.DataRow["Channel"].ToString(),
                                                           TestContext.DataRow["Rescan"].ToString(),
                                                           TestContext.DataRow["Sender"].ToString()).Result;
                Assert.IsTrue(result != null);
            }
            _mySelenium.Init(new Uri(TestContext.DataRow["NodeURL"].ToString()), new Uri(TestContext.DataRow["URLTestObject"].ToString()), TestContext.DataRow["Browser"].ToString());
            Assert.IsTrue(_mySelenium.GetBrowserTitle("Voxtron Web Client"));
            _mySelenium.Login(TestContext.DataRow["UserId"].ToString(), TestContext.DataRow["Password"].ToString(), TestContext.DataRow["VoxWebCltStatus"].ToString());
            _mySelenium.AcceptTask(0, int.Parse(TestContext.DataRow["ContactCodeItemIdx"].ToString()));
            _mySelenium.Logout();
            _mySelenium.Terminate();
            var InstanceCaller = new Thread(new ThreadStart(() => _mySelenium.CreateVideo()));
            InstanceCaller.Start();
        }
    }
}
