using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using System.Collections;
using System.Threading;
using System.Configuration;
using System.IO;
using Assert = NUnit.Framework.Assert;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Remote;
using System.Net;
using System.Web;


namespace Bookstore.Setup
{
    public class Browser
    {


        static IWebDriver _webDriver;
        protected IWebDriver driver;
        private const int DriverTimeOut = 30;
        public IWebDriver GetBrowserDriver(string browsername)

        {
            IWebDriver driver = null;
            driver = new ChromeDriver();
            if (driver != null)
            {
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(DriverTimeOut));
            }

            return driver;

        }
        protected IWebDriver StartWebDriver()

        {
            _webDriver = new ChromeDriver();
            return _webDriver;

       }
        protected void GoToApplication(string url)
        {

            _webDriver.Url = url;

        }

        protected void StopTests()
        {

            _webDriver.Quit();

        }

        protected void MaximizeWindow()
        {
            _webDriver.Manage().Window.Maximize();
        }

        public virtual void ImplicitWait()
        {
            var implicitlyWait = _webDriver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(30));
            // ReSharper disable once NotResolvedInText
            if (implicitlyWait == null) throw new ArgumentNullException("implicitlyWait");
        }

        protected void SwitchToWindow(string handle)
        {

            //Switches to another window
            _webDriver.SwitchTo().Window(handle);

        }
        protected IWebDriver Driver
        {

            get
            {
                return _webDriver;

            }

        }
        public static IEnumerable<string> BrowserToRunWith()
        {
            string[] browsers =
            { "firefox", "chrome"};

            foreach (string b in browsers)
            {
                yield return b;

            }
        }


    }
}
