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
    public class PageTests
    {
        [TestMethod()]
        public void Page_PageTest()
        {
            Page p = new Page("PageTest", 5, 6);
            Assert.AreEqual<int>(5, p.peekWidth);
            Assert.AreEqual<int>(6, p.peekHeight);
            Assert.AreEqual<int>(30, p.numberOfCells);
            Assert.AreEqual<string>("PageTest", p.peekTitle);
        }

        [TestMethod()]
        public void Page_GetSetValueTest()
        {
            Page p1 = new Page("GetSetValueTest_Page1", 5, 6);
            p1.SetValue("C3", "GetValueTest 1");
            Assert.AreEqual<string>("GetValueTest 1", p1.GetValue("C3"));

            Page p2 = new Page("GetSetValueTest_Page2", 5, 6);
            p2.SetValue(2, 2, "GetValueTest 2");
            Assert.AreEqual<string>("GetValueTest 2", p2.GetValue("C3"));

            Page p3 = new Page("GetSetValueTest_Page3", 5, 6);
            p3.SetValue("A1", "GetValueTest 3");
            Assert.AreEqual<string>("GetValueTest 3", p3.GetValue(0,0));

            Page p4 = new Page("GetSetValueTest_Page4", 5, 6);
            p4.SetValue(3, 3, "GetValueTest 4");
            Assert.AreEqual<string>("GetValueTest 4", p4.GetValue("D4"));
        }

        [TestMethod()]
        public void Page_AddRowTest()
        {
            Page p = new Page("AddRowTest_Page", 4, 4);
            Assert.AreEqual<int>(16, p.numberOfCells);
            Assert.IsNull(p.GetValue(4, 3));
            p.AddRow();
            Assert.AreEqual<int>(20, p.numberOfCells);
            Assert.IsNotNull(p.GetValue(4, 3));
        }

        [TestMethod()]
        public void Page_AddColumnTest()
        {
            Page p = new Page("AddColumnTest_Page", 4, 4);
            Assert.AreEqual<int>(16, p.numberOfCells);
            Assert.IsNull(p.GetValue(3, 4));
            p.AddColumn();
            Assert.AreEqual<int>(20, p.numberOfCells);
            Assert.IsNotNull(p.GetValue(3, 4));
        }

        [TestMethod()]
        public void Page_InsertRowAtTest()
        {
            Page p = new Page("InsertRowAtTest_Page", 2, 2);
            Assert.AreEqual<int>(4, p.numberOfCells);
            p.SetValue(0,0, "InsertRowAtTest 0,0");
            p.SetValue(0,1, "InsertRowAtTest 0,1");
            p.SetValue(1,0, "InsertRowAtTest 1,0");
            p.SetValue(1,1, "InsertRowAtTest 1,1");

            p.InsertRowAt(1);

            Assert.AreEqual<int>(6, p.numberOfCells);
            Assert.AreEqual("InsertRowAtTest 0,0", p.GetValue(0, 0));
            Assert.AreEqual("InsertRowAtTest 0,1", p.GetValue(0, 1));
            Assert.AreEqual("", p.GetValue(1, 0));
            Assert.AreEqual("", p.GetValue(1, 1));
            Assert.AreEqual("InsertRowAtTest 1,0", p.GetValue(2, 0));
            Assert.AreEqual("InsertRowAtTest 1,1", p.GetValue(2, 1));
        }

        [TestMethod()]
        public void Page_InsertColumnAtTest()
        {
            Page p = new Page("InsertColumnAtTest_Page", 2, 2);
            Assert.AreEqual<int>(4, p.numberOfCells);
            p.SetValue(0, 0, "InsertColumnAtTest 0,0");
            p.SetValue(0, 1, "InsertColumnAtTest 0,1");
            p.SetValue(1, 0, "InsertColumnAtTest 1,0");
            p.SetValue(1, 1, "InsertColumnAtTest 1,1");

            p.InsertColumnAt(1);

            Assert.AreEqual<int>(6, p.numberOfCells);
            Assert.AreEqual("InsertColumnAtTest 0,0", p.GetValue(0, 0));
            Assert.AreEqual("", p.GetValue(0, 1));
            Assert.AreEqual("InsertColumnAtTest 0,1", p.GetValue(0, 2));
            Assert.AreEqual("InsertColumnAtTest 1,0", p.GetValue(1, 0));
            Assert.AreEqual("", p.GetValue(1, 1));
            Assert.AreEqual("InsertColumnAtTest 1,1", p.GetValue(1, 2));
        }

        [TestMethod()]
        public void Page_RemoveLastRowTest()
        {
            Page p = new Page("RemoveLastRowTest", 2, 2);
            Assert.AreEqual<int>(4, p.numberOfCells);
            Assert.IsNotNull(p.GetValue(0, 0));
            Assert.IsNotNull(p.GetValue(0, 1));
            Assert.IsNotNull(p.GetValue(1, 0));
            Assert.IsNotNull(p.GetValue(1, 1));
            p.RemoveLastRow();
            Assert.AreEqual<int>(2, p.numberOfCells);
            Assert.IsNotNull(p.GetValue(0, 0));
            Assert.IsNotNull(p.GetValue(0, 1));
            Assert.IsNull(p.GetValue(1, 0));
            Assert.IsNull(p.GetValue(1, 1));
        }

        [TestMethod()]
        public void Page_RemoveRowAtTest()
        {
            Page p = new Page("RemoveRowAtTest_Page", 2, 3);
            Assert.AreEqual<int>(6, p.numberOfCells);
            p.SetValue(0, 0, "RemoveRowAtTest 0,0");
            p.SetValue(0, 1, "RemoveRowAtTest 0,1");
            p.SetValue(1, 0, "RemoveRowAtTest 1,0");
            p.SetValue(1, 1, "RemoveRowAtTest 1,1");
            p.SetValue(2, 0, "RemoveRowAtTest 2,0");
            p.SetValue(2, 1, "RemoveRowAtTest 2,1");

            p.RemoveRowAt(1);

            Assert.AreEqual<int>(4, p.numberOfCells);
            Assert.AreEqual("RemoveRowAtTest 0,0", p.GetValue(0, 0));
            Assert.AreEqual("RemoveRowAtTest 0,1", p.GetValue(0, 1));
            Assert.AreEqual("RemoveRowAtTest 2,0", p.GetValue(1, 0));
            Assert.AreEqual("RemoveRowAtTest 2,1", p.GetValue(1, 1));
            Assert.IsNull(p.GetValue(2, 0));
            Assert.IsNull(p.GetValue(2, 1));
        }

        [TestMethod()]
        public void Page_RemoveLastColumnTest()
        {
            Page p = new Page("RemoveColumnAtTest", 2, 2);
            Assert.AreEqual<int>(4, p.numberOfCells);
            Assert.IsNotNull(p.GetValue(0, 0));
            Assert.IsNotNull(p.GetValue(0, 1));
            Assert.IsNotNull(p.GetValue(1, 0));
            Assert.IsNotNull(p.GetValue(1, 1));
            p.RemoveLastColumn();
            Assert.AreEqual<int>(2, p.numberOfCells);
            Assert.IsNotNull(p.GetValue(0, 0));
            Assert.IsNull(p.GetValue(0, 1));
            Assert.IsNotNull(p.GetValue(1, 0));
            Assert.IsNull(p.GetValue(1, 1));
        }

        [TestMethod()]
        public void Page_RemoveColumnAtTest()
        {
            Page p = new Page("RemoveColumnAtTest_Page", 3, 2);
            Assert.AreEqual<int>(6, p.numberOfCells);
            p.SetValue(0, 0, "RemoveColumnAtTest 0,0");
            p.SetValue(0, 1, "RemoveColumnAtTest 0,1");
            p.SetValue(0, 2, "RemoveColumnAtTest 0,2");
            p.SetValue(1, 0, "RemoveColumnAtTest 1,0");
            p.SetValue(1, 1, "RemoveColumnAtTest 1,1");
            p.SetValue(1, 2, "RemoveColumnAtTest 1,2");

            p.RemoveColumnAt(1);

            Assert.AreEqual<int>(4, p.numberOfCells);

            Assert.AreEqual("RemoveColumnAtTest 0,0", p.GetValue(0, 0));
            Assert.AreEqual("RemoveColumnAtTest 0,2", p.GetValue(0, 1));
            Assert.IsNull(p.GetValue(0, 2));
            
            Assert.AreEqual("RemoveColumnAtTest 1,0", p.GetValue(1, 0));
            Assert.AreEqual("RemoveColumnAtTest 1,2", p.GetValue(1, 1));
            Assert.IsNull(p.GetValue(1, 2));       
        }

        [TestMethod()]
        public void Page_GetRowTest()
        {
            Page p = new Page("GetRowTest_Page", 2, 3);
            p.SetValue(0, 0, "GetRowTest 0,0");
            p.SetValue(0, 1, "GetRowTest 0,1");
            p.SetValue(1, 0, "GetRowTest 1,0");
            p.SetValue(1, 1, "GetRowTest 1,1");
            p.SetValue(2, 0, "GetRowTest 2,0");
            p.SetValue(2, 1, "GetRowTest 2,1");

            List<Cell> row0 = p.GetRow(0);
            Assert.AreEqual<int>(2, row0.Count);
            Assert.AreEqual<string>("GetRowTest 0,0", row0[0].Value);
            Assert.AreEqual<string>("GetRowTest 0,1", row0[1].Value);

            List<Cell> row1 = p.GetRow(1);
            Assert.AreEqual<int>(2, row1.Count);
            Assert.AreEqual<string>("GetRowTest 1,0", row1[0].Value);
            Assert.AreEqual<string>("GetRowTest 1,1", row1[1].Value);

            List<Cell> row2 = p.GetRow(2);
            Assert.AreEqual<int>(2, row2.Count);
            Assert.AreEqual<string>("GetRowTest 2,0", row2[0].Value);
            Assert.AreEqual<string>("GetRowTest 2,1", row2[1].Value);
        }

        [TestMethod()]
        public void Page_GetColumnTest()
        {
            Page p = new Page("GetColumnTest_Page", 3, 2);
            Assert.AreEqual<int>(6, p.numberOfCells);
            p.SetValue(0, 0, "GetColumnTest 0,0");
            p.SetValue(0, 1, "GetColumnTest 0,1");
            p.SetValue(0, 2, "GetColumnTest 0,2");
            p.SetValue(1, 0, "GetColumnTest 1,0");
            p.SetValue(1, 1, "GetColumnTest 1,1");
            p.SetValue(1, 2, "GetColumnTest 1,2");

            List<Cell> col0 = p.GetColumn(0);
            Assert.AreEqual<int>(2, col0.Count);
            Assert.AreEqual<string>("GetColumnTest 0,0", col0[0].Value);
            Assert.AreEqual<string>("GetColumnTest 1,0", col0[1].Value);

            List<Cell> col1 = p.GetColumn(1);
            Assert.AreEqual<int>(2, col1.Count);
            Assert.AreEqual<string>("GetColumnTest 0,1", col1[0].Value);
            Assert.AreEqual<string>("GetColumnTest 1,1", col1[1].Value);

            List<Cell> col2 = p.GetColumn(2);
            Assert.AreEqual<int>(2, col2.Count);
            Assert.AreEqual<string>("GetColumnTest 0,2", col2[0].Value);
            Assert.AreEqual<string>("GetColumnTest 1,2", col2[1].Value);
        }

        [TestMethod()]
        public void Page_FromObjectListTest()
        {
            Page p = Page.FromObjectList("FromObjectListTest_Page", sampleObjList());
            Assert.AreEqual<int>(9, p.numberOfCells);
            Assert.AreEqual<string>("FromObjectListTest_Page", p.peekTitle);
            Assert.AreEqual<string>("A1", p.GetValue(0, 0));
            Assert.AreEqual<string>("B1", p.GetValue(0, 1));
            Assert.AreEqual<string>("C1", p.GetValue(0, 2));
            Assert.AreEqual<string>("A2", p.GetValue(1, 0));
            Assert.AreEqual<string>("B2", p.GetValue(1, 1));
            Assert.AreEqual<string>("C2", p.GetValue(1, 2));
            Assert.AreEqual<string>("A3", p.GetValue(2, 0));
            Assert.AreEqual<string>("B3", p.GetValue(2, 1));
            Assert.AreEqual<string>("C3", p.GetValue(2, 2));
        }

        [TestMethod()]
        public void Page_ToObjectListTest()
        {
            IList<IList<Object>> testSample = sampleObjList();
            Page p = Page.FromObjectList("ToObjectListTest", testSample);
            IList<IList<Object>> resultList = Page.ToObjectList(p);

            Assert.IsNotNull(resultList);
            Assert.IsNotNull(resultList[0]);
            Assert.AreEqual<int>(9, resultList.Count * resultList[0].Count);       

            for (int row = 0; row < resultList.Count; row++)
                for (int col = 0; col < resultList[0].Count; col++)
                    Assert.AreEqual<Object>(testSample[row][col], resultList[row][col]);
        }

        private IList<IList<Object>> sampleObjList()
        {
            return (IList<IList<Object>>) new List<IList<Object>>
            {
                new List<Object> { "A1", "B1", "C1"},
                new List<Object> { "A2", "B2", "C2"},
                new List<Object> { "A3", "B3", "C3"}
            };
        }
    }
}
