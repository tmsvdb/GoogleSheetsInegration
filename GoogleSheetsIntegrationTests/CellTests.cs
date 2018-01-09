using Microsoft.VisualStudio.TestTools.UnitTesting;
using Beardiegames.GoogleSheetsIntegration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beardiegames.GoogleSheetsIntegration.Tests
{
    [TestClass()]
    public class CellTests
    {
        [TestMethod()]
        public void Cell_CellTest()
        {
            CellID cid = new CellID(3, 3);
            Cell cell = new Cell(cid, "CellTest");
            Assert.AreEqual<CellID>(cid, cell.Id);
            Assert.AreEqual<string>("CellTest", cell.Value);
        }

        [TestMethod()]
        public void Cell_CompareTest()
        {
            Cell cell = new Cell(new CellID(3, 6), "CompareTest");
            Assert.IsTrue(cell.Compare(3, 6));
        }

        [TestMethod()]
        public void Cell_CompareTest1()
        {
            Cell cell = new Cell(new CellID(3, 6), "CompareTest1");
            Assert.IsTrue(cell.Compare("G4"));
        }

        [TestMethod()]
        public void Cell_CompareColumnTest()
        {
            Cell cell = new Cell(new CellID(3, 6), "CompareColumnTest");
            Assert.IsTrue(cell.CompareColumn(6));
        }

        [TestMethod()]
        public void Cell_CompareRowTest()
        {
            Cell cell = new Cell(new CellID(3, 6), "CompareRowTest");
            Assert.IsTrue(cell.CompareRow(3));
        }
    }
}