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
    public class Administration:Browser
    {
        #region
        [FindsBy(How = How.XPath, Using = ".//*[@id='Header_Menu_Admin']/img")]
        private IWebElement _btnAdministration;
        [FindsBy(How = How.Id, Using = "Form_Field1")]
        private IWebElement _btnMembers;
        [FindsBy(How = How.Id, Using = "Search_name")]
        private IWebElement _txtSearch;
        [FindsBy(How = How.XPath, Using = ".//*[@id='Search_search_button']")]
        private IWebElement _btnSearch;
        [FindsBy(How = How.XPath, Using = ".//*[@id='Members_Repeater_ctl01_Members_member_login']")]
        private IWebElement _lnkMember;
        [FindsBy(How = How.Id, Using = "Record_member_login")]
        private IWebElement _lnkRecmember;
        [FindsBy(How = How.Id, Using = "Members_delete")]
        private IWebElement _btnMdelete;
        [FindsBy(How = How.Id, Using = "Login_name")]
        private IWebElement _txtAUserId;
        [FindsBy(How = How.Id, Using = "Login_password")]
        private IWebElement _txtAUserPassword;
        [FindsBy(How = How.Id, Using = "Login_login")]
        private IWebElement _btnALogin;
        private static string _DriveLocation;
        private WebDriverWait _wait;
        private IWebElement _element;
        #endregion

        public Administration(IWebDriver driver)
        {
            if (Driver != null) driver = Driver;
            PageFactory.InitElements(driver, page: this);
        }
        public ExcelWorksheet Readfromexcelsheet()
        {
            _DriveLocation = BasicPageActions.GetLocalDrive();
            string excelDataPath = _DriveLocation + "\\BookStoreData.xlsx";
            var package = new ExcelPackage(new FileInfo(excelDataPath));
            return package.Workbook.Worksheets[2];
        }
        public void deletemember(IWebDriver driver, ExcelWorksheet workSheet, int row1, int login)
        {

            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='Header_Menu_Admin']/img")));
            _btnAdministration.Click();
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Login_name")));
            _txtAUserId.SendKeys("admin");
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Login_password")));
            _txtAUserPassword.SendKeys("admin");
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Login_login")));
            _btnALogin.Click();
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Form_Field1")));
            _btnMembers.Click();
            IWebElement searchgrid = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Members_holder")));
            List<IWebElement> gridrows = searchgrid.FindElements(By.TagName("tr")).ToList();
            var rcountb4delete = gridrows.Count;
            //   var mlogin = workSheet.Cells[row1, login].Text;
            var mlogin = "user1";
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Search_name")));
            _txtSearch.SendKeys(mlogin);
            Thread.Sleep(1000);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='Search_search_button']")));
            _btnSearch.Click();
            Thread.Sleep(1000);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='Members_Repeater_ctl01_Members_member_login']")));
            _lnkMember.Click();
            Thread.Sleep(1000);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Record_member_login")));
            _lnkRecmember.Click();
            Thread.Sleep(1000);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Members_delete")));
            _btnMdelete.Click();
            Thread.Sleep(1000);
             searchgrid = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Members_holder")));
            List<IWebElement> grows = searchgrid.FindElements(By.TagName("tr")).ToList();
             var rcountaftrdelete = grows.Count;
            Assert.Greater(rcountb4delete, rcountaftrdelete);
               
        }
    }
}
