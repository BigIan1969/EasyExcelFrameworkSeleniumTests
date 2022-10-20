using EasyExcelFramework;
using EasyExcelFrameworkSelenium;
using System.Reflection;
using NUnit;

namespace SeleniumUnitTests
{
    [TestClass]
    public class TestLocators
    {
        [TestMethod]
        public void TestLocatorWorksheet()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\Test@Locators.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\ClickButton.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();

        }
        [TestMethod]
        public void TestLink()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\LinkAlert.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\LinkAlert.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
        }
        [TestMethod]
        public void TestPartialLink()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\LinkAlertPartial.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\LinkAlert.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
        }
        [TestMethod]
        public void TestLabelSelector()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\TestLabelLocator.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSendkeys.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
        }
        [TestMethod]
        public void TestText()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\TextAlert.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\LinkAlert.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
        }
    }
}