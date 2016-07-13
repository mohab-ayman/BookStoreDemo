using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Threading;
using System.Configuration;
using System.IO;
using System.Linq;
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
using Bookstore.Setup;
using Bookstore.Pages;

namespace BookstoreTestScript
{
    [TestFixture]
    public class LoginTest : Browser
    {
        private string _url = ConfigurationManager.AppSettings["baseUrl"];
        private IWebDriver _driver;
        private static string _DriveLocation;
        int row1 = 2;
        const int useridCol = 1;
        const int passwordCol = 2;
        [Test]
        [TestCaseSource(typeof(Browser), "BrowserToRunWith")]
        public void Logintobookstorehome(string browsername)
        {
            try
            {
                _driver = GetBrowserDriver(browsername);
                _driver.Navigate().GoToUrl(_url);
                _driver.Manage().Window.Maximize();
                _driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(30));
                var login = new Login(_driver);
                _DriveLocation = BasicPageActions.GetLocalDrive();
                string excelDataPath = _DriveLocation + "\\BookStoreData.xlsx";
                var package = new ExcelPackage(new FileInfo(excelDataPath));
                var workSheet = package.Workbook.Worksheets[1];
                login.logintobookstore(_driver, workSheet, row1, useridCol, passwordCol);
                _driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(30));
                //var exptitle = _driver.Title;
                //Assert.IsTrue(login.IsApplicationsTitleDisplayed());
                _driver.Close();



            }
            catch (Exception ex)
            {
                {
                    Console.WriteLine(ex.StackTrace);
                }
                throw;
            }

        }
        public LoginTest()
        {

        }

    }
    }


