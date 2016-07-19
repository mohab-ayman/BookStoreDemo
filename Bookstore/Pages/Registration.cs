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
    public class Registration : Browser
    {
        #region
        [FindsBy(How = How.Id, Using = "Header_Menu_Reg")]
        private IWebElement _btnregister;
        [FindsBy(How = How.Id, Using = "Reg_member_login")]
        private IWebElement _txtMLogin;
        [FindsBy(How = How.Id, Using = "Reg_member_password")]
        private IWebElement _txtMPassword;
        [FindsBy(How = How.Id, Using = "Reg_member_password2")]
        private IWebElement _txtMConfirmPass;
        [FindsBy(How = How.Id, Using = "Reg_first_name")]
        private IWebElement _txtMFname;
        [FindsBy(How = How.Id, Using = "Reg_last_name")]
        private IWebElement _txtMLname;
        [FindsBy(How = How.Id, Using = "Reg_email")]
        private IWebElement _txtMEmail;
        [FindsBy(How = How.Id, Using = "Reg_address")]
        private IWebElement _txtMAddress;
        [FindsBy(How = How.Id, Using = "Reg_phone")]
        private IWebElement _txtMPhone;
        [FindsBy(How = How.Id, Using = "Reg_card_type_id")]
        private IWebElement _lsttMCreditCardType;
        [FindsBy(How = How.XPath, Using = " .//*[@id='Reg_card_number']")]
        private IWebElement _txtMCreditCardNum;
        [FindsBy(How = How.XPath, Using = " .//*[@id='Reg_insert']")]
        private IWebElement _btnMRegister;
        [FindsBy(How = How.Id, Using = " Reg_cancel")]
        private IWebElement _btnMCancel;
        private static string _DriveLocation;
        private WebDriverWait _wait;
        private IWebElement _element;
        private SelectElement _select;
        #endregion

        public Registration(IWebDriver driver)
        {
            if (Driver != null) driver = Driver;
            PageFactory.InitElements(driver, page: this);
        }
        public void registernewmember(IWebDriver driver, ExcelWorksheet workSheet1, int row1, int login,
        int mpassword, int confirmpass, int firstname, int lastname, int email, int address, int phone, int creditcardtype, int creditcardnumber)
        {

            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Header_Menu_Reg")));
            _btnregister.Click();
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Reg_member_login")));
            var mlogin = workSheet1.Cells[row1, login].Text;
            _txtMLogin.SendKeys(mlogin);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Reg_member_password")));
            var mpass = workSheet1.Cells[row1, mpassword].Text;
            _txtMPassword.SendKeys(mpass);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Reg_member_password2")));
            var mcpass = workSheet1.Cells[row1, confirmpass].Text;
            _txtMConfirmPass.SendKeys(mcpass);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Reg_first_name")));
            var mfname = workSheet1.Cells[row1, firstname].Text;
            _txtMFname.SendKeys(mfname);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Reg_last_name")));
            var mlname = workSheet1.Cells[row1, lastname].Text;
            _txtMLname.SendKeys(mlname);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Reg_email")));
            var mEmail = workSheet1.Cells[row1, email].Text;
            _txtMEmail.SendKeys(mEmail);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Reg_address")));
            var mAddress = workSheet1.Cells[row1, address].Text;
            _txtMAddress.SendKeys(mAddress);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Reg_phone")));
            var mPhone = workSheet1.Cells[row1, phone].Text;
            _txtMPhone.SendKeys(mPhone);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Reg_card_type_id")));
            _select = new SelectElement(_element);
            IList<IWebElement> mcTypes = _select.Options;
            var mCtype = workSheet1.Cells[row1, creditcardtype].Text;
            foreach (IWebElement x in mcTypes)
            {
                if (mCtype.Trim().Equals(x.Text.Trim()))
                {
                    _select.SelectByText(x.Text);
                    break;

                }

            }
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='Reg_card_number']")));
            var mCnum = workSheet1.Cells[row1, creditcardnumber].Text;
            _txtMCreditCardNum.SendKeys(mCnum);
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(500));
            _element = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='Reg_insert']")));
            _btnMRegister.Click();
        }
        public ExcelWorksheet Readfromexcelsheet()
        {
            _DriveLocation = BasicPageActions.GetLocalDrive();
            string excelDataPath = _DriveLocation + "\\BookStoreData.xlsx";
            var package = new ExcelPackage(new FileInfo(excelDataPath));
            return package.Workbook.Worksheets[2];
        }

    }
}
