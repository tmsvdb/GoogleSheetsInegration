using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beardiegames.GoogleSheetsIntegration.Tests
{
    [TestClass()]
    public class SheetManagerTests
    {
        [TestMethod()]
        public void SheetManagerTest()
        {
            // Test1: create a new sheetmanager
            SheetManager manager = new SheetManager("client_id.json", "Test SheetManager");
            // chech if the namager has a running sheetservice
            Assert.IsTrue(manager.isRunning);
        }

        [TestMethod()]
        public void GetSheetTest()
        {
            // Prep: create a manager to test
            SheetManager manager = new SheetManager("client_id.json", "Test GetSheet");

            // Test1: create and get a new sheet
            Sheet sheet1 = manager.GetSheet("1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            // Result:
            // does it return a working sheet object
            Assert.IsNotNull(sheet1);
            // is the manager tracking the new sheet
            Assert.AreEqual<int>(manager.numberOfSheets, 1);
            // id the sheet returned the same as the requested sheet
            Assert.AreEqual<string>(sheet1.Id, "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc"); 

            // Test2: add the same sheet twice
            Sheet sheet2 = manager.GetSheet("1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            // Result:
            // does it return a working sheet object
            Assert.IsNotNull(sheet2);
            // did the manager create/added a new sheet or use the same sheet by checking the amount of sheets
            Assert.AreEqual<int>(manager.numberOfSheets, 1);
            // is the requested sheet the same object as the sheet created in the previous test
            Assert.AreSame(sheet1, sheet2);

            // Test3: create and get a second new sheet
            Sheet sheet3 = manager.GetSheet("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms");
            // Result:
            // does it return a working sheet object
            Assert.IsNotNull(sheet3);
            // is the manager tracking the new sheet
            Assert.AreEqual<int>(manager.numberOfSheets, 2);
            // is the requested sheet the same object as the sheet created in the previous test
            Assert.AreNotSame(sheet1, sheet3);
            // id the sheet returned the same as the requested sheet
            Assert.AreEqual<string>(sheet3.Id, "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms");
        }
    }
}