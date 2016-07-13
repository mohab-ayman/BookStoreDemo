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
   public class BasicPageActions:Browser
    {
        public void NavigateToSite()
        {
            //  StartWebDriver();
            var url = ConfigurationManager.AppSettings["baseUrl"];
            GoToApplication(url);
            MaximizeWindow();
        }


        public void CloseTest()
        {
            StopTests();
        }

        public static string GetLocalDrive()
        {
            var drivePath = System.IO.DriveInfo.GetDrives().GetValue(0).ToString();
            return drivePath;
        }
    }
}
