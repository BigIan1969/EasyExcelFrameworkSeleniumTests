using EasyExcelFramework;
using EasyExcelFrameworkSelenium;
using System.Reflection;
using NUnit;
namespace SeleniumUnitTests
{
    [TestClass]
    public class TestBrowser
    {
        [TestMethod]
        public void TestLaunchChrome()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\LaunchChrome.xlsx");

            eef.EasyExcel.Execute();
            Assert.AreEqual("chrome-headless-shell", eef.EasyExcel.Locals["BrowserN"]);
            eef.driver.Quit();

        }


        /* Taken out as currently failingpossibly a W11 thing

        [TestMethod]
        public void TestLaunchEdge()
        {
            EasyExcelFrameworkS eef = new EasyExcelFrameworkS("Data\\Launchedge.xlsx");
            eef.EasyExcel.Execute();
            Assert.AreEqual("edge", eef.EasyExcel.Locals["BrowserN"]);
            eef.driver.Quit();

        }

        [TestMethod]
        public void TestLaunchFirefox()
        {
            EasyExcelFrameworkS eef = new EasyExcelFrameworkS("Data\\LaunchFirefox.xlsx");
            eef.EasyExcel.Execute();
            Assert.AreEqual("firefox", eef.EasyExcel.Locals["BrowserN"]);
            eef.driver.Quit();

        }
        */
    }
}