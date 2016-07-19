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
    [TestClass]
    public class RegistrationTest : Browser
    {
        private string _url = ConfigurationManager.AppSettings["baseUrl"];
        private IWebDriver _driver;
        private static string _DriveLocation;
        const int login = 1;
        const int mpassword=2;
        const int confirmpass =3;
        const int firstname =4;
        const int lastname =5;
        const int email =6;
        const int address =7;
        const int phone =8;
        const int creditCardtype =9;
        const int creditcardnumber =10;
        [TestMethod]
        [TestCaseSource(typeof(Browser), "BrowserToRunWith")]
        public void Registermembers(string browsername)
        {
            try
            {
                _driver = GetBrowserDriver(browsername);
                _driver.Navigate().GoToUrl(_url);
                _driver.Manage().Window.Maximize();
                _driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(60));
                var reg = new Registration(_driver);
                var workSheet = reg.Readfromexcelsheet();
                var start = workSheet.Dimension.Start;
                var end = workSheet.Dimension.End;
                for (int row1 = 2; row1 <= end.Row; row1++)
                {

                    reg.registernewmember(_driver, workSheet, row1, login, mpassword, confirmpass, firstname, lastname, email, address, phone, creditCardtype, creditcardnumber);
                }
                
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
        public RegistrationTest()
        {

        }
    }
      
}
