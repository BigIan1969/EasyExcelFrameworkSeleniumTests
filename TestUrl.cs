using EasyExcelFramework;
using EasyExcelFrameworkSelenium;
using System.Reflection;
using NUnit;

namespace SeleniumUnitTests
{
    [TestClass]
    public class TestUrl
    {
        [TestMethod]
        public void TestGotoURL()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\GotoURLGoogle.xlsx");
            eef.EasyExcel.Execute();
            Assert.AreEqual("https://www.google.com/", eef.EasyExcel.Locals["NewURL"]);
            eef.driver.Quit();

        }
    }
}