using EasyExcelFramework;
using EasyExcelFrameworkSelenium;
using System.Reflection;
using NUnit;
using OpenQA.Selenium;

namespace SeleniumUnitTests
{
    [TestClass]
    public class TestInputControls
    {
        [TestMethod]
        public void TestClickButton()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\ClickButton.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\ClickButton.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();

        }
        [TestMethod]
        public void TestClickNonExistentButton()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\ClickNonExistentButton.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\ClickButton.html";
            Exception ex = Assert.ThrowsException<InvalidSelectorException>(() => eef.EasyExcel.Execute());
            eef.driver.Quit();

        }
        [TestMethod]
        public void TestTypeNonExistentInput()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\TypeNonExistentInput.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\ClickButton.html";
            Exception ex = Assert.ThrowsException<InvalidSelectorException>(() => eef.EasyExcel.Execute());
            eef.driver.Quit();

        }
        [TestMethod]
        public void TestInput()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputSendkeys.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSendkeys.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("Ian", eef.EasyExcel.Locals["FirstName"]);
            Assert.AreEqual("Holdsworth", eef.EasyExcel.Locals["LastName"]);
        }
        [TestMethod]
        public void TestClear()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputClear.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSendkeys.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("", eef.EasyExcel.Locals["FirstName"]);
            Assert.AreEqual("", eef.EasyExcel.Locals["LastName"]);
        }
        [TestMethod]
        public void TestSubmit()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\SubmitForm.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSendkeys.html";
            try
            {
                eef.EasyExcel.Execute();
            }
            catch (NullReferenceException) //discard expected error Chrome headless issue
            { }
            catch
            {
                eef.driver.Quit();
                Assert.AreEqual("Your file couldn’t be accessed", eef.EasyExcel.Globals["ErrorMessage"]);
            }
        }
        [TestMethod]
        public void TestCheck()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputCheckbox.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputCheckbox.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("true", eef.EasyExcel.Locals["Vehicle1"]);
        }
        [TestMethod]
        public void TestUncheck()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputCheckbox.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputCheckbox.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreNotEqual("true", eef.EasyExcel.Locals["Vehicle2"]);
        }
        [TestMethod]
        public void TestRadio()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputRadiobuttons.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputRadiobuttons.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("false", eef.EasyExcel.Locals["Age1"]);
            Assert.AreEqual("true", eef.EasyExcel.Locals["Age2"]);
        }
        [TestMethod]
        public void TestSelectByText()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputSelectByText.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSelect.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("saab", eef.EasyExcel.Locals["Car"]);
        }
        [TestMethod]
        public void TestSelectByValue()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputSelectByValue.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSelect.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("opel", eef.EasyExcel.Locals["Car"]);
        }
        [TestMethod]
        public void TestSelectByIndex()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputSelectByIndex.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSelect.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("audi", eef.EasyExcel.Locals["Car"]);
        }

        [TestMethod]
        public void TestDeselectByText()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputDeSelectByText.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSelectMulti.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("", eef.EasyExcel.Locals["Car"]);
        }
        [TestMethod]
        public void TestDeselectByValue()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputDeSelectByValue.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSelectMulti.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("", eef.EasyExcel.Locals["Car"]);
        }
        [TestMethod]
        public void TestDeselectByIndex()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputDeSelectByIndex.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSelectMulti.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("", eef.EasyExcel.Locals["Car"]);
        }
        [TestMethod]
        public void TestSelectAll()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputSelectAll.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSelectMulti.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("Volvo|Saab|Opel|Audi", eef.EasyExcel.Locals["Car"]);
        }
        [TestMethod]
        public void TestDeselectAll()
        {
            EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium eef = new EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium("Data\\InputDeSelectAll.xlsx");
            eef.EasyExcel.Globals["TargetURL"] = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Data\\InputSelectMulti.html";
            eef.EasyExcel.Execute();
            eef.driver.Quit();
            Assert.AreEqual("", eef.EasyExcel.Locals["Car"]);
        }

    }
}