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
    public class CellIDTests
    {
        [TestMethod()]
        public void CellIDTest()
        {
            CellID cid = new CellID(12, 28);
            Assert.AreEqual<int>(12, cid.rowIndex, "cid row = " + cid.rowIndex.ToString());
            Assert.AreEqual<int>(28, cid.colIndex, "cid col = " + cid.colIndex.ToString());
        }

        [TestMethod()]
        public void ToStringFormatTest()
        {
            CellID cid1 = new CellID(3, 3);
            Assert.AreEqual<string>("C3", cid1.ToStringFormat());

            CellID cid2 = new CellID(12, 3);
            Assert.AreEqual<string>("C12", cid2.ToStringFormat());

            CellID cid3 = new CellID(3, 28);
            Assert.AreEqual<string>("AB3", cid3.ToStringFormat());

            CellID cid4 = new CellID(12, 28);
            Assert.AreEqual<string>("AB12", cid4.ToStringFormat());
        }

        [TestMethod()]
        public void FromStringFormatTest()
        {
            CellID cid1 = CellID.FromStringFormat("C3");
            Assert.AreEqual<int>(3, cid1.rowIndex);
            Assert.AreEqual<int>(3, cid1.colIndex);

            CellID cid2 = CellID.FromStringFormat("C12");
            Assert.AreEqual<int>(12, cid2.rowIndex);
            Assert.AreEqual<int>(3, cid2.colIndex);

            CellID cid3 = CellID.FromStringFormat("AB3");
            Assert.AreEqual<int>(3, cid3.rowIndex);
            Assert.AreEqual<int>(28, cid3.colIndex);

            CellID cid4 = CellID.FromStringFormat("AB12");
            Assert.AreEqual<int>(12, cid4.rowIndex);
            Assert.AreEqual<int>(28, cid4.colIndex);
        }

        [TestMethod()]
        public void ColumnToNumberTest()
        {
            Assert.AreEqual<int>(3, CellID.ColumnToNumber("C"));
            Assert.AreEqual<int>(28, CellID.ColumnToNumber("AB"));
            Assert.AreEqual<int>(702, CellID.ColumnToNumber("ZZ"));
        }

        [TestMethod()]
        public void NumberToColumnTest()
        {
            Assert.AreEqual<string>("C", CellID.NumberToColumn(3));
            Assert.AreEqual<string>("AB", CellID.NumberToColumn(28));
            Assert.AreEqual<string>("ZZ", CellID.NumberToColumn(702));
        }
    }
}