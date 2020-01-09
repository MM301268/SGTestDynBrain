using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SeleniumWorker
{
    [TestClass]
    public class PerformanceTest
    {
        #region TestContext
        public TestContext TestContext { get; set; }
        private SeleniumWorker.SelWorker _mySelenium;
        #endregion

        [TestMethod]
        [TestCategory("Actual")]
        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV",
                    "|DataDirectory|\\testdata\\PerformanceTest1.csv",
                    "PerformanceTest1#csv",
                    DataAccessMethod.Sequential),
                    DeploymentItem("testdata\\PerformanceTest1.csv")]
        [DeploymentItem("avcodec-57.dll")]
        [DeploymentItem("avdevice-57.dll")]
        [DeploymentItem("avfilter-6.dll")]
        [DeploymentItem("avformat-57.dll")]
        [DeploymentItem("avutil-55.dll")]
        [DeploymentItem("postproc-54.dll")]
        [DeploymentItem("swresample-2.dll")]
        [DeploymentItem("swscale-4.dll")]
        public void PerformanceTest1()
        {
            _mySelenium = new SeleniumWorker.SelWorker("CreateBorObject");
            SelWorker.ChangeLogFileNameWoDate("RollingLogFileAppender", "CreateBorObject");
            if (TestContext.DataRow["CreateBorObject"].ToString() == "true")
            {
                var result = SelWorker.SendRequestToRest(DateTime.Now,
                                                           TestContext.DataRow["Category"].ToString(),
                                                           TestContext.DataRow["Location"].ToString(),
                                                           TestContext.DataRow["Channel"].ToString(),
                                                           TestContext.DataRow["Rescan"].ToString(),
                                                           TestContext.DataRow["Sender"].ToString()).Result;
            }

        }

    }
}
