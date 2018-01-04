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
    public class MatrixTests
    {
        [TestMethod()]
        public void MatrixTest()
        {
            Matrix mtx = new Matrix(5, 6);
            Assert.AreEqual<int>(5, mtx.peekWidth);
            Assert.AreEqual<int>(6, mtx.peekHeight);
            Assert.AreEqual<int>(30, mtx.numberOfCells);
        }

        [TestMethod()]
        public void GetSetValueTest()
        {
            Matrix mtx1 = new Matrix(5, 6);
            mtx1.SetValue("C3", "GetValueTest 1");
            Assert.AreEqual<string>("GetValueTest 1", mtx1.GetValue("C3"));

            Matrix mtx2 = new Matrix(5, 6);
            mtx2.SetValue(2, 2, "GetValueTest 2");
            Assert.AreEqual<string>("GetValueTest 2", mtx2.GetValue("C3"));

            Matrix mtx3 = new Matrix(5, 6);
            mtx3.SetValue("A1", "GetValueTest 3");
            Assert.AreEqual<string>("GetValueTest 3", mtx3.GetValue(0,0));

            Matrix mtx4 = new Matrix(5, 6);
            mtx4.SetValue(3, 3, "GetValueTest 4");
            Assert.AreEqual<string>("GetValueTest 4", mtx4.GetValue("D4"));
        }

        [TestMethod()]
        public void AddRowTest()
        {
            Matrix mtx1 = new Matrix(4, 4);
            Assert.AreEqual<int>(16, mtx1.numberOfCells);
            Assert.IsNull(mtx1.GetValue(4, 3));
            mtx1.AddRow();
            Assert.AreEqual<int>(20, mtx1.numberOfCells);
            Assert.IsNotNull(mtx1.GetValue(4, 3));
        }

        [TestMethod()]
        public void AddColumnTest()
        {
            Matrix mtx1 = new Matrix(4, 4);
            Assert.AreEqual<int>(16, mtx1.numberOfCells);
            Assert.IsNull(mtx1.GetValue(3, 4));
            mtx1.AddColumn();
            Assert.AreEqual<int>(20, mtx1.numberOfCells);
            Assert.IsNotNull(mtx1.GetValue(3, 4));
        }

        [TestMethod()]
        public void InsertRowAtTest()
        {
            Matrix mtx1 = new Matrix(2, 2);
            Assert.AreEqual<int>(4, mtx1.numberOfCells);
            mtx1.SetValue(0,0, "InsertRowAtTest 0,0");
            mtx1.SetValue(0,1, "InsertRowAtTest 0,1");
            mtx1.SetValue(1,0, "InsertRowAtTest 1,0");
            mtx1.SetValue(1,1, "InsertRowAtTest 1,1");

            mtx1.InsertRowAt(1);

            Assert.AreEqual<int>(6, mtx1.numberOfCells);
            Assert.AreEqual("InsertRowAtTest 0,0", mtx1.GetValue(0, 0));
            Assert.AreEqual("InsertRowAtTest 0,1", mtx1.GetValue(0, 1));
            Assert.AreEqual("", mtx1.GetValue(1, 0));
            Assert.AreEqual("", mtx1.GetValue(1, 1));
            Assert.AreEqual("InsertRowAtTest 1,0", mtx1.GetValue(2, 0));
            Assert.AreEqual("InsertRowAtTest 1,1", mtx1.GetValue(2, 1));
        }

        [TestMethod()]
        public void InsertColumnAtTest()
        {
            Matrix mtx1 = new Matrix(2, 2);
            Assert.AreEqual<int>(4, mtx1.numberOfCells);
            mtx1.SetValue(0, 0, "InsertColumnAtTest 0,0");
            mtx1.SetValue(0, 1, "InsertColumnAtTest 0,1");
            mtx1.SetValue(1, 0, "InsertColumnAtTest 1,0");
            mtx1.SetValue(1, 1, "InsertColumnAtTest 1,1");

            mtx1.InsertColumnAt(1);

            Assert.AreEqual<int>(6, mtx1.numberOfCells);
            Assert.AreEqual("InsertColumnAtTest 0,0", mtx1.GetValue(0, 0));
            Assert.AreEqual("", mtx1.GetValue(0, 1));
            Assert.AreEqual("InsertColumnAtTest 0,1", mtx1.GetValue(0, 2));
            Assert.AreEqual("InsertColumnAtTest 0,0", mtx1.GetValue(1, 0));
            Assert.AreEqual("", mtx1.GetValue(1, 1));
            Assert.AreEqual("InsertColumnAtTest 1,1", mtx1.GetValue(1, 2));
        }

        [TestMethod()]
        public void RemoveLastRowTest()
        {
            Matrix mtx1 = new Matrix(2, 2);
            Assert.AreEqual<int>(4, mtx1.numberOfCells);
            Assert.IsNotNull(mtx1.GetValue(0, 0));
            Assert.IsNotNull(mtx1.GetValue(0, 1));
            Assert.IsNotNull(mtx1.GetValue(1, 0));
            Assert.IsNotNull(mtx1.GetValue(1, 1));
            mtx1.RemoveLastRow();
            Assert.AreEqual<int>(2, mtx1.numberOfCells);
            Assert.IsNotNull(mtx1.GetValue(0, 0));
            Assert.IsNotNull(mtx1.GetValue(0, 1));
            Assert.IsNull(mtx1.GetValue(1, 0));
            Assert.IsNull(mtx1.GetValue(1, 1));
        }

        [TestMethod()]
        public void RemoveRowAtTest()
        {
            Matrix mtx1 = new Matrix(2, 3);
            Assert.AreEqual<int>(6, mtx1.numberOfCells);
            mtx1.SetValue(0, 0, "RemoveRowAtTest 0,0");
            mtx1.SetValue(0, 1, "RemoveRowAtTest 0,1");
            mtx1.SetValue(1, 0, "RemoveRowAtTest 1,0");
            mtx1.SetValue(1, 1, "RemoveRowAtTest 1,1");
            mtx1.SetValue(2, 0, "RemoveRowAtTest 2,0");
            mtx1.SetValue(2, 1, "RemoveRowAtTest 2,1");

            mtx1.RemoveRowAt(1);

            Assert.AreEqual<int>(4, mtx1.numberOfCells);
            Assert.AreEqual("RemoveRowAtTest 0,0", mtx1.GetValue(0, 0));
            Assert.AreEqual("RemoveRowAtTest 0,1", mtx1.GetValue(0, 1));
            Assert.AreEqual("RemoveRowAtTest 2,0", mtx1.GetValue(1, 0));
            Assert.AreEqual("RemoveRowAtTest 2,1", mtx1.GetValue(1, 1));
            Assert.IsNull(mtx1.GetValue(2, 0));
            Assert.IsNull(mtx1.GetValue(2, 1));
        }

        [TestMethod()]
        public void RemoveLastColumnTest()
        {
            Matrix mtx1 = new Matrix(2, 2);
            Assert.AreEqual<int>(4, mtx1.numberOfCells);
            Assert.IsNotNull(mtx1.GetValue(0, 0));
            Assert.IsNotNull(mtx1.GetValue(0, 1));
            Assert.IsNotNull(mtx1.GetValue(1, 0));
            Assert.IsNotNull(mtx1.GetValue(1, 1));
            mtx1.RemoveLastColumn();
            Assert.AreEqual<int>(2, mtx1.numberOfCells);
            Assert.IsNotNull(mtx1.GetValue(0, 0));
            Assert.IsNull(mtx1.GetValue(0, 1));
            Assert.IsNotNull(mtx1.GetValue(1, 0));
            Assert.IsNull(mtx1.GetValue(1, 1));
        }

        [TestMethod()]
        public void RemoveColumnAtTest()
        {
            Matrix mtx1 = new Matrix(2, 3);
            Assert.AreEqual<int>(6, mtx1.numberOfCells);
            mtx1.SetValue(0, 0, "RemoveColumnAtTest 0,0");
            mtx1.SetValue(0, 1, "RemoveColumnAtTest 0,1");
            mtx1.SetValue(0, 2, "RemoveColumnAtTest 0,2");
            mtx1.SetValue(1, 0, "RemoveColumnAtTest 1,0");
            mtx1.SetValue(1, 1, "RemoveColumnAtTest 1,1");
            mtx1.SetValue(1, 2, "RemoveColumnAtTest 1,2");

            mtx1.RemoveColumnAt(1);

            Assert.AreEqual<int>(3, mtx1.numberOfCells);

            Assert.AreEqual("RemoveColumnAtTest 0,0", mtx1.GetValue(0, 0));
            Assert.AreEqual("RemoveColumnAtTest 0,2", mtx1.GetValue(0, 1));
            Assert.IsNull(mtx1.GetValue(0, 2));
            
            Assert.AreEqual("RemoveColumnAtTest 1,0", mtx1.GetValue(1, 0));
            Assert.AreEqual("RemoveColumnAtTest 1,2", mtx1.GetValue(1, 1));
            Assert.IsNull(mtx1.GetValue(1, 2));       
        }

        [TestMethod()]
        public void GetRowTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void GetColumnTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void FromObjectListTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void ToObjectListTest()
        {
            Assert.Fail();
        }
    }
}
