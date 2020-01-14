namespace SeleniumWorker
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Text;
    using System.Threading.Tasks;
    using Accord.Video.FFMPEG;
    using log4net;
    using Newtonsoft.Json;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Firefox;
    using OpenQA.Selenium.IE;
    using OpenQA.Selenium.Remote;
    using OpenQA.Selenium.Support.UI;

    public class SelWorker : IDisposable
    {
        #region Declaration Variables
        private static readonly TimeSpan WaitNewItemsTimeout = TimeSpan.FromSeconds(15);
        private static readonly TimeSpan ImplicitWaitTimeout = TimeSpan.FromSeconds(20);
        private static readonly TimeSpan PageLoadTimeout = TimeSpan.FromSeconds(60);
        private static readonly ILog Logger = LogManager.GetLogger(typeof(SelWorker));
        private static readonly List<Bitmap> ImageLst = new List<Bitmap>();
        private static string _category;
        private static string _externalID;
        private static string _contactID;
        private readonly string _testcase;
        //private readonly ZZLoggerUtil _logutil = new ZZLoggerUtil();
        private readonly string _date = DateTime.Now.ToString("yyyMMdd-HHmm", DateTimeFormatInfo.InvariantInfo);
        private WebDriverWait _wait;
        public  IWebDriver _driver;
        private Datatypes.WaitingMonitorStruct _recordWM;
        private List<Datatypes.WaitingMonitorStruct> _listWM;
        private static readonly CultureInfo CultureInfoGerman = new CultureInfo("de-DE", false);
        

        #endregion

        #region Constructor
        public SelWorker(string testcaseName) => _testcase = testcaseName;
        #endregion

        #region ChangeLogFileName
        public void ChangeLogFileName(string appenderName, string newFilename)
        {
            var repository = log4net.LogManager.GetRepository();
            foreach (var appender in repository.GetAppenders())
            {
                if (string.Compare(appender.Name, appenderName, new System.StringComparison()) == 0 && appender is log4net.Appender.FileAppender)
                {
                    var fileAppender = (log4net.Appender.FileAppender)appender;
                    var directoryName = Path.GetDirectoryName(fileAppender.File);
                    fileAppender.File = System.IO.Path.Combine(directoryName + "//", newFilename + "_" + _date + ".log");
                    fileAppender.ActivateOptions();
                }
            }
        }
        #endregion

        #region ChangeLogFileNameWoDate
        public static void ChangeLogFileNameWoDate(string appenderName, string newFilename)
        {
            var repository = log4net.LogManager.GetRepository();
            foreach (var appender in repository.GetAppenders())
            {
                if (string.Compare(appender.Name, appenderName, new System.StringComparison()) == 0 && appender is log4net.Appender.FileAppender)
                {

                    var fileAppender = (log4net.Appender.FileAppender)appender;
                    var directoryName = Path.GetDirectoryName(fileAppender.File);
                    fileAppender.File = System.IO.Path.Combine(directoryName + "//", newFilename + ".log");
                    fileAppender.ActivateOptions();
                }
            }
        }
        #endregion

        #region Init
        public void Init(Uri nodeURL, Uri url, string browserName)
        {
            Logger.Debug(ZZLoggerUtil.CreateUnderlines("Starting Testcase: " + _testcase));
            Logger.Debug("Starting Testcase: " + _testcase);
            Logger.Debug("Init-Start");

            switch (browserName)
            {
                case "Chrome":
                {
                    var options = new ChromeOptions();
                    options.AddArguments("--disable-gpu");
                    options.AddArguments("--disable-extensions");
                    options.AddArguments("--window-size=1920, 1080");
                    _ = options.ToCapabilities();

                    //options.AddAdditionalCapability("platform", "WIN10", true);
                    options.AddAdditionalCapability("browsername", browserName, true);

                    _driver = new RemoteWebDriver(nodeURL, options.ToCapabilities());
                    Logger.Debug("Remote driver Chrome has been instanciated");
                    break;
                }
                case "Firefox":
                {
                    var options = new FirefoxOptions
                    {
                        Profile = new FirefoxProfile()
                    };
                    _ = options.ToCapabilities();
                    options.AddAdditionalCapability("browsername", browserName, true);

                    options.Profile.SetPreference("insecure_field_warning.contextual.enabled", false);
                    options.Profile.SetPreference("security.insecure_field_warning.contextual.enabled", false);

                    _driver = new RemoteWebDriver(nodeURL, options.ToCapabilities());
                    Logger.Debug("Remote driver Firefox has been instanciated");
                    break;
                }
                case "Internet Explorer":
                {
                    var ieOptions = new InternetExplorerOptions
                    {
                        IntroduceInstabilityByIgnoringProtectedModeSettings = true,
                        IgnoreZoomLevel = true
                    };
                    try
                    {
                        _driver = new RemoteWebDriver(nodeURL, ieOptions.ToCapabilities(), PageLoadTimeout);
                    }
                    catch (Exception ex)
                    {
                        _ = ex.ToString();
                    }
                    break;
                }

            }

            _driver.Manage().Window.Position = new System.Drawing.Point(0, 0);
            Logger.Debug("_driver position set to: " + _driver.Manage().Window.Position.ToString());

            _driver.Manage().Window.Size = new Size(1920, 1080);
            Logger.Debug("_driver window size set to: " + _driver.Manage().Window.Size.ToString());

            _driver.Manage().Cookies.DeleteAllCookies();
            Logger.Debug("Cookies has been deleted");

            _driver.Manage().Timeouts().ImplicitWait = ImplicitWaitTimeout;
            Logger.Debug("ImplicitWaitTimeout set to: " + ImplicitWaitTimeout.ToString());

            _driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(60);
            Logger.Debug("PageLoadTimeout set to: " + PageLoadTimeout.ToString());

            _wait = new WebDriverWait(_driver, WaitNewItemsTimeout);
            Logger.Debug("_driver _wait for Objects: " + WaitNewItemsTimeout.ToString());

            _driver.Navigate().GoToUrl(url);
            Logger.Debug("Navigation to URL: " + url);

            _ = System.Threading.Thread.Yield();

            Logger.Debug("Init End");
        }
        #endregion

        #region IsElementPresent
        private bool IsElementPresent(By by)
        {
            try
            {
                _driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        #endregion

        #region Login
        public void Login(string argusername, string argpassword, string argLoginStatus)
        {

            Logger.Debug("Login Start");
            try
            {
                _ = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id("txtUsername")));
                Logger.Debug("Waiting for text field username");

                ClearNFillElement(ref _driver, _driver.FindElement(By.Id("txtUsername")), argusername);
                Logger.Debug("Username: " + argusername);

                ClearNFillElement(ref _driver, _driver.FindElement(By.Id("txtPassword")), argpassword);
                Logger.Debug("Password: " + argusername);

                ClickElement(ref _driver, _driver.FindElement(By.Id("btnOK")));
                Logger.Debug("OK-Button was clicked");

                if (IsElementPresent(By.Id("btnAlreadyLoggedInOK")))
                {
                    try
                    {
                        _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
                        Logger.Debug("ImplicitWaitTimeout set to: " + ImplicitWaitTimeout.ToString());
                        _driver.FindElement(By.Id("btnAlreadyLoggedInOK")).Click();
                        Logger.Debug("Already logged in dialog comitted");
                    }
                    catch(ElementNotInteractableException ex)
                    {
                        _ = ex.ToString();
                        _driver.Manage().Timeouts().ImplicitWait = ImplicitWaitTimeout;
                        Logger.Debug("ImplicitWaitTimeout set to: " + ImplicitWaitTimeout.ToString());
                    }
                }

                _ = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id("cboInitialStatus")));
                Logger.Debug("Wait for initial status selection dialog");

                SelectComboBox(_driver, _driver.FindElement(By.Id("cboInitialStatus")), "cboInitialStatus_list_", argLoginStatus);

                Logger.Debug("Client initial status: " + argLoginStatus);

                var telefonCheckbox = _driver.FindElement(By.Id("chkQueue4"));
                if (telefonCheckbox.Selected)
                {
                    //_driver.FindElement(By.CssSelector("#divQueue3 label")).Click();
                    ClickElement(ref _driver, _driver.FindElement(By.CssSelector("#divQueue4 label")));
                }
                Logger.Debug("Phone deselected");

                ClickElementFades(ref _driver, _driver.FindElement(By.CssSelector(".table-buttons:nth-child(5) > .btn-default")));
                Logger.Debug("OK button pressed successful");

                _ = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Id("btnPauseSubMenu")));
                Logger.Debug("Wait for button pause");

                _ = System.Threading.Thread.Yield();
            }
            catch (Exception ex)
            {
                Logger.Fatal(ex.ToString());
            }
            finally
            {
                Logger.Debug("Login End");
            }
        }
        #endregion

        #region GetBrowserTitle
        public bool GetBrowserTitle(string title)
        {
            Logger.Debug("Browser Title: " + title);
            return _driver.Title.Contains(title);
        }
        #endregion

        #region ToggleWaitingMonitor
        public void ToggleWaitingMonitor()
        {
            try
            {
                _ = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id("btnWaitingMonitor")));
                Logger.Debug("Waiting Monitor is clickable");

                ClickElement(ref _driver, _driver.FindElement(By.Id("btnWaitingMonitor")));
                Logger.Debug("Waiting Monitor clicked successful");

                _ = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Id("ui-id-7")));
            }
            catch (Exception ex)
            {
                Logger.Fatal(ex.ToString());
            }
            _ = System.Threading.Thread.Yield();
            //System.Threading.Thread.Sleep(500);
        }
        #endregion

        #region ShowWaitMonObjects
        public void ShowWaitMonObjects()
        {
            try
            {
                _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);
                _ = System.Threading.Thread.Yield();
                //Check weather Contact Center calls are displayed
                try
                {
                    var switchCaller = _driver.FindElement(By.CssSelector("div[class='monitor-filter cc-call-incoming active']"));
                    Logger.Debug("Waiting Monitor shows calls");
                }
                catch (NoSuchElementException ex)
                {
                    Logger.Debug("Waiting Monitor does NOT show calls by default");
                    ClickElement(ref _driver, _driver.FindElement(By.CssSelector("div[class='monitor-filter cc-call-incoming']")));
                    Logger.Debug("Waiting Monitor enabled to show calls");
                    _ = ex.ToString();
                }
                //Check weather Contact Center objects are displayed
                try
                {
                    var switchCaller = _driver.FindElement(By.CssSelector("div[class='monitor-filter cc-object-incoming active']"));
                    Logger.Debug("Waiting Monitor enabled to show calls");
                }
                catch (NoSuchElementException ex)
                {
                    Logger.Debug("Waiting Monitor does NOT show objects by default");
                    ClickElement(ref _driver, _driver.FindElement(By.CssSelector("div[class='monitor-filter cc-object-incoming']")));
                    Logger.Debug("Waiting Monitor enabled to shows objects");
                    _ = ex.ToString();
                }
                _driver.Manage().Timeouts().ImplicitWait = ImplicitWaitTimeout;
                _ = System.Threading.Thread.Yield();
            }
            catch (Exception ex)
            {
                Logger.Fatal(ex.ToString());
            }
        }
        #endregion

        #region ReadWaitingMonitor
        public void ReadWaitingMonitor()
        {
            IReadOnlyCollection<IWebElement> rowElements = _driver.FindElements(By.XPath("//*[@id='dialogWaitingMonitor']/div[2]/div/table[2]/tbody/tr"));
            Logger.Debug("Number of elements found " + rowElements.Count);

            _recordWM = new Datatypes.WaitingMonitorStruct();
            _listWM = new List<Datatypes.WaitingMonitorStruct>();
            for (var i = 1; i <= rowElements.Count; i++)
            {
                ElementHighlightON(ref _driver, _driver.FindElement(By.XPath("//*[@id='dialogWaitingMonitor']/div[2]/div/table[2]/tbody")));
                AddScreenShotToList(ref _driver);
                _recordWM.Type = _driver.FindElement(By.XPath("//*[@id='dialogWaitingMonitor']/div[2]/div/table[2]/tbody/tr[" + i + "]/td[2]")).Text;
                _recordWM.From = _driver.FindElement(By.XPath("//*[@id='dialogWaitingMonitor']/div[2]/div/table[2]/tbody/tr[" + i + "]/td[3]")).Text;
                _recordWM.To = _driver.FindElement(By.XPath("//*[@id='dialogWaitingMonitor']/div[2]/div/table[2]/tbody/tr[" + i + "]/td[4]")).Text;
                _recordWM.Status = _driver.FindElement(By.XPath("//*[@id='dialogWaitingMonitor']/div[2]/div/table[2]/tbody/tr[" + i + "]/td[7]")).Text;
                _recordWM.WatingTime = _driver.FindElement(By.XPath("//*[@id='dialogWaitingMonitor']/div[2]/div/table[2]/tbody/tr[" + i + "]/td[8]")).Text;
                _recordWM.Skill = _driver.FindElement(By.XPath("//*[@id='dialogWaitingMonitor']/div[2]/div/table[2]/tbody/tr[" + i + "]/td[9]")).Text;
                _recordWM.Queue = _driver.FindElement(By.XPath("//*[@id='dialogWaitingMonitor']/div[2]/div/table[2]/tbody/tr[" + i + "]/td[10]")).Text;
                _listWM.Add(_recordWM);
                Logger.Debug(_recordWM.Type + "," + _recordWM.From + "," + _recordWM.To + "," + _recordWM.Status + "," + _recordWM.WatingTime + "," + _recordWM.Skill + "," + _recordWM.Queue);
                ElementHighlightOFF(ref _driver, _driver.FindElement(By.XPath("//*[@id='dialogWaitingMonitor']/div[2]/div/table[2]/tbody/tr[" + i + "]")));
                AddScreenShotToList(ref _driver);
            }
        }
        #endregion

        #region CheckWatingMonitorLoops
        public bool CheckWatingMonitorLoops(string skill, int argiterations)
        {
            var resFindChkWaitMon = false;
            for (var i = 1; i <= argiterations; i++)
            {
                //Check if usecase is in _waiting monitor
                resFindChkWaitMon = CheckWaitingMonitor(skill);
                if (resFindChkWaitMon)
                {
                    break;
                }
            }
            return resFindChkWaitMon;
        }
        #endregion

        #region SendRequestToRest
        public static async Task<string> SendRequestToRest(DateTimeOffset creationTimestamp, string category, string location, string channel, string rescan, string sender)
        {
            var dt = DateTime.Now; // Or whatever
            var s = dt.ToString("yyyyMMddHHmmss",CultureInfoGerman);
            //Create Request
            var req = new RouteRequest
            {
                ExternalId = "A_" + s,
                CreationTimestamp = creationTimestamp
            };
            req.Tags.Add("scan_date", creationTimestamp.ToString(CultureInfoGerman));
            req.Tags.Add("category", category);
            req.Tags.Add("location", location);
            req.Tags.Add("channel", channel);
            req.Tags.Add("rescan", rescan);
            req.Tags.Add("sender", sender);

            //Convert to Json
            var myJson = JsonConvert.SerializeObject(req);

            using (var client = new HttpClient())
            {
                try
                {
                    //Routerrequest => DEV-RZ1-SQL-01.dev-onipd.de
                    var response = await client.PostAsync(new Uri("https://pst-rz1-dwr-01.post-test.de:1443/routerequest"), new StringContent(myJson, Encoding.UTF8, "application/json")).ConfigureAwait(true);
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        var respJson = await response.Content.ReadAsStringAsync().ConfigureAwait(true);
                        _contactID = respJson;
                        _category = category;
                        _externalID = req.ExternalId;
                        Logger.Debug("Object-Id gernerated by Rest-API: " + JsonConvert.DeserializeObject<RouteResponse>(respJson).ContactId);
                        response.Dispose();
                        return JsonConvert.DeserializeObject<RouteResponse>(respJson).ContactId;
                    }
                    else
                    {
                        Logger.Fatal(response.ToString());
                        response.Dispose();
                        return null;
                    }

                }
                catch (Exception ex)
                {
                    Logger.Fatal(ex.ToString());
                    return null;
                }
            }
        }
        #endregion

        #region GoOnline
        public void GoOnline()
        {
            ClickElement(ref _driver, _driver.FindElement(By.Id("btnPauseSubMenu")));
            Logger.Debug("WebClient went online");

        }
        #endregion

        #region GetStaleElemById
        public IWebElement GetStaleElemById(string id)
        {
            try
            {
                return _driver.FindElement(By.XPath(id));
            }
            catch (Exception ex)
            {
                _ = ex.ToString();
                Logger.Debug("Attempting to recover from StaleElementReferenceException ...");
                return GetStaleElemById(id);
            }
        }
        #endregion

        #region AcceptTask
        public void AcceptTask(int numberOfIterations, int contactCodeItemIdx)
        {
            var i = 0;
            System.Threading.Thread.Sleep(3000);
            do
            {
                try
                {
                    HandleFrame0();

                    IReadOnlyCollection<IWebElement> contactProperties = _driver.FindElements(By.XPath(".//div[@class='dataStyle']/span"));
                    var contactProp_Category = contactProperties.ElementAt(10);
                    var contactProp_ExternalId = contactProperties.ElementAt(13);
                    var stringComparision = new StringComparison();
                    while ((string.Compare(contactProp_Category.Text, _category, stringComparision) != 0) || (string.Compare(contactProp_ExternalId.Text, _externalID, stringComparision) != 0))
                    {
                        ClickAcceptContact();

                        ClickContactCode(contactCodeItemIdx);

                        ClickContactHandled();

                        SwitchBackToParentFrame();

                        StopWrapUp();

                        HandleFrame0();

                        contactProperties = _driver.FindElements(By.XPath(".//div[@class='dataStyle']/span"));
                        contactProp_Category = contactProperties.ElementAt(10);
                        contactProp_ExternalId = contactProperties.ElementAt(13);
                        stringComparision = new StringComparison();
                    }


                    ClickAcceptContact();

                    ClickContactCode(contactCodeItemIdx);

                    ClickContactHandled();

                    SwitchBackToParentFrame();

                    StopWrapUp();
                }
                catch (Exception ex)
                {
                    //_ = ex.ToString();
                    Logger.Debug("This may not be an error in all cases!!!");
                    Logger.Debug(ex.ToString());
                    break;
                }
                if (numberOfIterations != 0)
                {
                    i++;
                }
                else
                {
                    i = -1;
                }
            }
            while (i < numberOfIterations);
        }
        #endregion

        #region ClickAcceptContact
        public void ClickAcceptContact()
        {
            _ = System.Threading.Thread.Yield();
            System.Threading.Thread.Sleep(500);
            ClickElement(ref _driver, _driver.FindElement(By.CssSelector(".btn-contact")));
            Logger.Debug("Accepted Task in WebClient");
        }
        #endregion

        #region ClickContactCode
        public void ClickContactCode(int contactCodeItemIdx)
        {
            _ = System.Threading.Thread.Yield();
            System.Threading.Thread.Sleep(500);
            if (contactCodeItemIdx >= 0)
            {
                IReadOnlyCollection<IWebElement> contactCodes = _driver.FindElements(By.XPath("//*[contains(@id,'treeItem')]/div/label"));
                var contactCode = contactCodes.ElementAt(contactCodeItemIdx);
                ClickElement(ref _driver, contactCode);
                Logger.Debug("Clicked on contact code im WebClient " + contactCode.Text);
            }
            else
            {
                Logger.Debug("Contact Code not clicked");
            }
        }
        #endregion

        #region ClickContactHandled
        public void ClickContactHandled()
        {
            _ = System.Threading.Thread.Yield();
            System.Threading.Thread.Sleep(500);
            ClickElement(ref _driver, _driver.FindElement(By.CssSelector(".btn-contact-handled")));
            Logger.Debug("Finished task im WebClient");
        }
        #endregion

        #region SwitchBackToParentFrame
        public void SwitchBackToParentFrame()
        {
            _ = System.Threading.Thread.Yield();
            System.Threading.Thread.Sleep(500);
            _ = _driver.SwitchTo().ParentFrame();
            Logger.Debug("Switched back to parent frame");
        }
        #endregion

        #region StopWrapUp
        public void StopWrapUp()
        {
            _ = System.Threading.Thread.Yield();
            System.Threading.Thread.Sleep(500);
            _ = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.CssSelector(".btn-wrapup-stop")));
            ClickElementFades(ref _driver, _driver.FindElement(By.CssSelector(".btn-wrapup-stop")));
            Logger.Debug("Stopped wrapup");

            _ = System.Threading.Thread.Yield();
            System.Threading.Thread.Sleep(2000);
        }
        #endregion

        #region HandleFrame0
        public void HandleFrame0()
        {
            _ = System.Threading.Thread.Yield();
            _ = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[contains(@id,'iframeCustomSubPage')]")));
            Logger.Debug("Frame 0 has been displayed!");

            System.Threading.Thread.Sleep(500);
            var myframe = GetStaleElemById("//*[contains(@id,'iframeCustomSubPage')]");
            _ = _driver.SwitchTo().Frame(myframe);
            Logger.Debug("Frame 0 clicked");
        }
        #endregion

        #region Logout
        public void Logout()
        {
            try
            {
                _ = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id("btnLogoff")));
                Logger.Debug("Button Logoff has appeared");

                ClickElementFades(ref _driver, _driver.FindElement(By.Id("btnLogoff")));

                Logger.Debug("Button Logoff has been pressed successful");

                _ = System.Threading.Thread.Yield();
                System.Threading.Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                Logger.Fatal(ex.ToString());
                Terminate();
            }
        }
        #endregion

        #region Terminate
        public void Terminate()
        {
            Logger.Debug("Termination Start");
            AddScreenShotToList(ref _driver);
            try
            {
                _driver.Close();
                Logger.Debug("_driver closed");
            }
            catch(Exception ex)
            {
                Logger.Fatal("Problem in closing driver");
                Logger.Fatal(ex.ToString());
            }
            try
            {
                _driver.Quit();
                Logger.Debug("_driver terminated");
            }
            catch(Exception ex)
            {
                Logger.Fatal("Problem in quitting driver");
                Logger.Fatal(ex.ToString());
            }
            Logger.Debug("Termination End");

            Logger.Debug("Ending Testcase: " + _testcase);
            Logger.Debug(ZZLoggerUtil.CreateUnderlines("Ending Testcase: " + _testcase));

        }
        #endregion

        #region AddScreenShotToList
        public static void AddScreenShotToList(ref IWebDriver driver)
        {
            try
            {
                var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                var image = (Bitmap)new ImageConverter().ConvertFrom(screenshot.AsByteArray);
                ImageLst.Add(image);
            }
            catch (Exception ex)
            {
                Logger.Debug(ex.ToString());
            }
        }
        #endregion

        #region GetScreenShot
        public static string GetScreenShot(ref IWebDriver driver)
        {
            var programDirectory = AppDomain.CurrentDomain.BaseDirectory;
            programDirectory = Path.GetFullPath(Path.Combine(programDirectory, @"..\..\..\"));
            var screenShotDirectory = Path.Combine(programDirectory + "Screenshots");

            if (File.Exists(screenShotDirectory) == false)
            {
                _ = Directory.CreateDirectory(screenShotDirectory);
            }
            
            try
            {
                var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                var date = DateTime.Now.ToString("yyyMMdd-HHmmssff", DateTimeFormatInfo.InvariantInfo);
                var localpath = screenShotDirectory  + "\\" +  date + ".png";
                screenshot.SaveAsFile(localpath, ScreenshotImageFormat.Png);
                return localpath;
            }
            catch (Exception ex)
            {
                Logger.Debug(ex.ToString());
                return null;
            }
        }
        #endregion

        #region CreateVideo
        public void CreateVideo()
        {
            var width = 1920;
            var height = 1080;
            var newimage = new Bitmap(width, height);
            if (width % 2 == 1)
            {
                width++;
            }
            if (height % 2 == 1)
            {
                height++;
            }
            try
            {
                var videoFileWriter = new VideoFileWriter();
                var programDirectory = AppDomain.CurrentDomain.BaseDirectory;
                programDirectory = Path.GetFullPath(Path.Combine(programDirectory, @"..\..\..\"));
                var videoDirectory = Path.Combine(programDirectory + "\\Video\\");
                if (File.Exists(videoDirectory) == false)
                {
                    _ = Directory.CreateDirectory(videoDirectory);
                }

                videoFileWriter.Open(videoDirectory + _testcase + "_" + _date + ".avi", width, height, 25, VideoCodec.MPEG4);
                foreach (var image in ImageLst)
                {
                    newimage = (Bitmap)new Bitmap(image, new Size(1920, 1080));
                    for (var i = 0; i < 10; i++)
                    {
                        videoFileWriter.WriteVideoFrame(newimage);
                    }
                }
                videoFileWriter.Close();
            }
            catch (Exception ex)
            {
                Logger.Fatal(ex.ToString());
            }
            finally
            {
                newimage.Dispose();
            }
        }
        #endregion

        #region ClearNFillElement
        public static void ClearNFillElement(ref IWebDriver driver, IWebElement element, string argString)
        {
            if (element != null)
            {
                AddScreenShotToList(ref driver);
                ElementHighlightON(ref driver, element);
                AddScreenShotToList(ref driver);
                element.Clear();
                AddScreenShotToList(ref driver);
                element.SendKeys(argString);
                AddScreenShotToList(ref driver);
                ElementHighlightOFF(ref driver, element);
                AddScreenShotToList(ref driver);
            }
        }
        #endregion

        #region ClickElement
        public static void ClickElement(ref IWebDriver driver, IWebElement element)
        {
            if (element != null)
            {
                ElementHighlightON(ref driver, element);
                AddScreenShotToList(ref driver);
                element.Click();
                AddScreenShotToList(ref driver);
                ElementHighlightOFF(ref driver, element);
                AddScreenShotToList(ref driver);
            }
        }
        #endregion

        #region ClickElementFades
        public static void ClickElementFades(ref IWebDriver driver, IWebElement element)
        {
            if (element != null)
            {
                ElementHighlightON(ref driver, element);
                AddScreenShotToList(ref driver);
                element.Click();
                AddScreenShotToList(ref driver);

            }
        }
        #endregion

        #region ElementToText
        public static string ElementToText(ref IWebDriver driver, IWebElement element)
        {
            if (element != null)
            {
                ElementHighlightON(ref driver, element);
                AddScreenShotToList(ref driver);
                ElementHighlightOFF(ref driver, element);
                AddScreenShotToList(ref driver);
                return element.Text;
            }
            return string.Empty;
        }
        #endregion

        #region SelectComboBox
        public void SelectComboBox(IWebDriver driver, IWebElement element, string elementID, string loginStatus)
        {
            //var initialStatusDropDown = element;
            ElementHighlightON(ref driver, element);
            AddScreenShotToList(ref driver);
            _driver.FindElement(By.Id("cboInitialStatus")).Click();
            _ = System.Threading.Thread.Yield();
            SelectDropDownElement(ref driver, elementID, loginStatus);
            AddScreenShotToList(ref driver);
            ElementHighlightOFF(ref driver, element);
            AddScreenShotToList(ref driver);
        }
        #endregion

        #region ElementHighlightON
        public static void ElementHighlightON(ref IWebDriver driver, IWebElement element)
        {
            var js = (IJavaScriptExecutor)driver;
            _ = js.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", element, "color: red; border: 5px solid red;");
        }
        #endregion

        #region ElementHighlightOFF
        public static void ElementHighlightOFF(ref IWebDriver driver, IWebElement element)
        {

            var js = (IJavaScriptExecutor)driver;
            _ = js.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", element, string.Empty);

        }
        #endregion

        #region SelectDropDownElement
        public static void SelectDropDownElement(ref IWebDriver driver, string elementID, string argLoginStatus)
        {
            switch (argLoginStatus)
            {
                case "Free":
                {
                    var fullElementDescription = elementID + "FREE";
                    driver.FindElement(By.Id(fullElementDescription)).Click();
                    break;
                }
                case "Pause (Autom. Abgemeldet)":
                {
                    var fullElementDescription = elementID + "1";
                    driver.FindElement(By.Id(fullElementDescription)).Click();
                    break;
                }
                case "Pause (Bei Start VCC)":
                {
                    var fullElementDescription = elementID + "2";
                    driver.FindElement(By.Id(fullElementDescription)).Click();
                    break;
                }
                case "Pause (Coaching)":
                {
                    var fullElementDescription = elementID + "3";
                    driver.FindElement(By.Id(fullElementDescription)).Click();
                    break;
                }
                case "Pause (Mittagspause)":
                {
                    var fullElementDescription = elementID + "4";
                    driver.FindElement(By.Id(fullElementDescription)).Click();
                    break;
                }
                case "Pause (R-Beleg)":
                {
                    var fullElementDescription = elementID + "5";
                    driver.FindElement(By.Id(fullElementDescription)).Click();
                    break;
                }
                case "Pause (Teamdialog)":
                {
                    var fullElementDescription = elementID + "6";
                    driver.FindElement(By.Id(fullElementDescription)).Click();
                    break;
                }
            }
        }
        #endregion

        #region CheckWaitingMonitor
        private bool CheckWaitingMonitor(string skill)
        {
            ReadWaitingMonitor();
            var found = false;
            foreach (var record in _listWM)
            {
                if ((record.Skill == skill) && (record.WatingTime.Length == 3))
                {
                    found = true;
                    Logger.Debug("Bor-Object found!!! Skill: " + record.Skill + " _waiting Time: " + record.WatingTime);
                }
            }
            return found;
        }
        #endregion

        #region Displose
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // dispose managed resources
                _driver.Close();
            }
            // free native resources
        }
        #endregion

        #region Dispose without Param
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }

}
