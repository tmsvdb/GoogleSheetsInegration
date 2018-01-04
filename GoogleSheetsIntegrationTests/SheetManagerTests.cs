using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Beardiegames.GoogleSheetsIntegration;

namespace Beardiegames.GoogleSheetsIntegration.Tests
{
    [TestClass()]
    public class SheetManagerTests
    {
        [TestMethod()]
        public void SheetManagerTest()
        {
            SheetManager manager = new SheetManager("client_id.json", "Test SheetManager");
            Assert.IsTrue(manager.isRunning);
        }

        [TestMethod()]
        public void GetSheetTest()
        {
            SheetManager manager = new SheetManager("client_id.json", "Test GetSheet");

            Sheet sheet1 = manager.GetSheet("1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.IsNotNull(sheet1);
            Assert.AreEqual<int>(manager.numberOfSheets, 1);
            Assert.AreEqual<string>(sheet1.Id, "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc"); 

            Sheet sheet2 = manager.GetSheet("1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.IsNotNull(sheet2);
            Assert.AreEqual<int>(manager.numberOfSheets, 1);
            Assert.AreSame(sheet1, sheet2);

            Sheet sheet3 = manager.GetSheet("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms");
            Assert.IsNotNull(sheet3);
            Assert.AreEqual<int>(manager.numberOfSheets, 2);
            Assert.AreNotSame(sheet1, sheet3);
            Assert.AreEqual<string>(sheet3.Id, "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms");
        }
    }
}