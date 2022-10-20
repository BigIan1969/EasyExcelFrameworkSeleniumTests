using EasyExcelFramework;
using EasyExcelFrameworkSelenium;
using System.Reflection;
using NUnit;

namespace SeleniumUnitTests
{
    [TestClass]
    public class TestAlerts
    {
        [TestMethod]
        public void TestAcceptAlert()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\AcceptAlert.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\ClickButton.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();

        }
        [TestMethod]
        public void TestDismissAlert()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\DismissAlert.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\ClickButton.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();

        }
    }
}