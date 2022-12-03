using EasyExcelFramework;
using EasyExcelFrameworkSelenium;
using System.Reflection;
using NUnit;

namespace SeleniumUnitTests
{
    [TestClass]
    public class TestIframe
    {
        [TestMethod]
        public void TestSwitchFrameByIndex()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputIframeByIndex.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\Iframe.html";
            eef.EasyExcel.Execute();
            eef.EasyExcel.finish();
            eef.driver.Quit();
            Assert.AreEqual("Ian", eef.EasyExcel.Globals["FirstName"]); ;
            Assert.AreEqual("Holdsworth", eef.EasyExcel.Globals["LastName"]);
            Assert.AreEqual("Ian1", eef.EasyExcel.Globals["FirstName1"]);
            Assert.AreEqual("Holdsworth1", eef.EasyExcel.Globals["LastName1"]);
            Assert.AreEqual("Ian2", eef.EasyExcel.Globals["FirstName2"]);
            Assert.AreEqual("Holdsworth2", eef.EasyExcel.Globals["LastName2"]);

        }
        [TestMethod]
        public void TestSwitchFrameByName()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputIframeByName.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\Iframe.html";
            eef.EasyExcel.Execute();
            eef.EasyExcel.finish();
            eef.driver.Quit();
            Assert.AreEqual("Ian", eef.EasyExcel.Globals["FirstName"]); ;
            Assert.AreEqual("Holdsworth", eef.EasyExcel.Globals["LastName"]);
            Assert.AreEqual("Ian1", eef.EasyExcel.Globals["FirstName1"]);
            Assert.AreEqual("Holdsworth1", eef.EasyExcel.Globals["LastName1"]);
            Assert.AreEqual("Ian2", eef.EasyExcel.Globals["FirstName2"]);
            Assert.AreEqual("Holdsworth2", eef.EasyExcel.Globals["LastName2"]);

        }
        [TestMethod]
        public void TestSwitchFrameByID()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputIframeByID.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\Iframe.html";
            eef.EasyExcel.Execute();
            eef.EasyExcel.finish();
            eef.driver.Quit();
            Assert.AreEqual("Ian", eef.EasyExcel.Globals["FirstName"]); ;
            Assert.AreEqual("Holdsworth", eef.EasyExcel.Globals["LastName"]);
            Assert.AreEqual("Ian1", eef.EasyExcel.Globals["FirstName1"]);
            Assert.AreEqual("Holdsworth1", eef.EasyExcel.Globals["LastName1"]);
            Assert.AreEqual("Ian2", eef.EasyExcel.Globals["FirstName2"]);
            Assert.AreEqual("Holdsworth2", eef.EasyExcel.Globals["LastName2"]);

        }
        [TestMethod]
        public void TestSwitchFrameByElement()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputIframeByElement.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\Iframe.html";
            eef.EasyExcel.Execute();
            eef.EasyExcel.finish();
            eef.driver.Quit();
            Assert.AreEqual("Ian", eef.EasyExcel.Globals["FirstName"]); ;
            Assert.AreEqual("Holdsworth", eef.EasyExcel.Globals["LastName"]);
            Assert.AreEqual("Ian1", eef.EasyExcel.Globals["FirstName1"]);
            Assert.AreEqual("Holdsworth1", eef.EasyExcel.Globals["LastName1"]);
            Assert.AreEqual("Ian2", eef.EasyExcel.Globals["FirstName2"]);
            Assert.AreEqual("Holdsworth2", eef.EasyExcel.Globals["LastName2"]);

        }
        //[TestMethod]
        //public void TestSwitchFrameParent()
        //{
        //    EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputIframeParent.xlsx");
        //    eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\IframeNested.html";
        //    eef.EasyExcel.Execute();
        //    eef.EasyExcel.finish();
        //    eef.driver.Quit();
        //    Assert.AreEqual("Ian", eef.EasyExcel.Globals["FirstName"]); ;
        //    Assert.AreEqual("Holdsworth", eef.EasyExcel.Globals["LastName"]);
        //    Assert.AreEqual("Ian1", eef.EasyExcel.Globals["FirstName1"]);
        //    Assert.AreEqual("Holdsworth1", eef.EasyExcel.Globals["LastName1"]);
        //    Assert.AreEqual("Ian2", eef.EasyExcel.Globals["FirstName2"]);
        //    Assert.AreEqual("Holdsworth2", eef.EasyExcel.Globals["LastName2"]);

        //}
    }
}