using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using fr.avh.braille.addin;
using fr.avh.braille.dictionnaire.Entities;
using fr.avh.braille.dictionnaire;

namespace fr.avh.braille.tests
{
    [TestClass]
    public class TestAddinUpdater
    {
        [TestMethod]
        [DataRow("0.0.0.0", "v1.3.4.5", true)]
        [DataRow("1.3.4.6", "v1.3.4.5", false)]
        [DataRow("1.3.4.5", "v1.3.4.5", false)]
        [DataRow("1.3.4.5", "v1.3.4.6", true)]
        [DataRow("1.3.4.5", "v1.3.6.6", true)]
        [DataRow("1.3.4.5", "v1.4", true)]
        [DataRow("1.3.5", "v1.4", true)]
        [DataRow("1.4", "v2", true)]
        public void TestVersionComparaison(string v1, string v2, bool expectV1isLowerThanV2)
        {
            ulong v1Int = AddinUpdater.computeVersionComparator(v1);
            ulong v2Int = AddinUpdater.computeVersionComparator(v2);
            Assert.AreEqual(
                expectV1isLowerThanV2,
                v1Int < v2Int,
                "v1:<{0}>={1} , v2:{2}={3}",
                new object[] { v1, v1Int, v2, v2Int });
        }
    }
}
