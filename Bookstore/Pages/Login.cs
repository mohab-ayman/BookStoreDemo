﻿using System;
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
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Remote;
using System.Net;
using System.Web;
using Bookstore.Setup;

namespace Bookstore.Pages
{
   public class Login:Browser
    {
        #region
        [FindsBy(How = How.Id, Using = "Header_Menu_Field1")]
        private IWebElement _btnsignin;
        [FindsBy(How = How.Id, Using = "Login_name")]
        private IWebElement _txtUserId;
        [FindsBy(How = How.Id, Using = "Login_password")]
        private IWebElement _txtUserPassword;
        [FindsBy(How = How.Id, Using = "Login_login")]
        private IWebElement _btnLogin;
        [FindsBy(How = How.Id, Using = "MemberForm_Title")]
        private IWebElement _Userinfotitle;
        
        bool displayed = true;
        private static string _DriveLocation;
        private WebDriverWait _wait;
        private IWebElement _element;
        #endregion

        public Login(IWebDriver driver)
        {
            if (Driver != null) driver = Driver;
            PageFactory.InitElements(driver, page: this);
        }
        public void logintobookstore(IWebDriver driver, ExcelWorksheet workSheet, int row1, int useridCol,
          int passwordCol)
        {
            
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Header_Menu_Field1")));
            _btnsignin.Click();
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Login_name")));
            var username = workSheet.Cells[row1, useridCol].Text;
            _txtUserId.SendKeys(username);
            var password = workSheet.Cells[row1, passwordCol].Text;
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementExists(By.Id("Login_password")));
            _txtUserPassword.SendKeys(password);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Login_login")));
            _btnLogin.Click();
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("MemberForm_Title")));
            
            
        }

        public ExcelWorksheet Readfromexcelsheet ()
        {
            _DriveLocation = BasicPageActions.GetLocalDrive();
            string excelDataPath = _DriveLocation + "\\BookStoreData.xlsx";
            var package = new ExcelPackage(new FileInfo(excelDataPath));
            return package.Workbook.Worksheets[1];


        }
        public bool IsUserInformationTitleDisplayed()
        {
            return _Userinfotitle.Displayed;

        }


    }
}
