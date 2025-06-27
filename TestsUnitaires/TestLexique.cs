using fr.avh.braille.dictionnaire;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace fr.avh.braille.tests
{
    [TestClass]
    public class TestLexique
    {
        [TestMethod]
        [DataRow("chamallow", true)]
        [DataRow("croyez-moi", true)]
        [DataRow("peut-être", true)]
        [DataRow("puis-je", true)]
        [DataRow("dis-moi", true)]
        public void TestMotFrançais(string mot, bool expected)
        {
            Assert.AreEqual(
                expected,
                LexiqueFrance.EstFrancais(mot),
                "Mot:<{0}>",
                new object[] { mot });
        }
    }
}
