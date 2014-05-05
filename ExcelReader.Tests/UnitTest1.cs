using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReader.Tests
{
    [TestClass]
    public class UidGenerationTest
    {
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestForInvalidInputException_nulls()
        {
            UidGenerator.GenerateUid(null, null, null, 1);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestForInvalidInputException_too_short()
        {
            UidGenerator.GenerateUid("ee", "ee", "ee", 1);
        }

        [TestMethod]
        public void TestFor_basic_uid_generation()
        {
            var uid = UidGenerator.GenerateUid("afsal ahmed", "test", "test", 1);
            Assert.AreEqual(uid, "tete1af");
        }

        [TestMethod]
        public void TestFor_lowercase_onlyid()
        {
            var uid = UidGenerator.GenerateUid("Afsal Ahmed", "Center Test", "Location Test", 1);
            Assert.AreEqual(uid, "celo1af");
        }

        [TestMethod]
        public void TestFor_alphabet_onlyid()
        {
            var uid = UidGenerator.GenerateUid("M.S Afsal Ahmed", "Center Test", "Location Test", 1);
            Assert.AreEqual(uid, "celo1ms");

            uid = UidGenerator.GenerateUid(" M.S Afsal Ahmed", "Center Test", "Location Test", 1);
            Assert.AreEqual(uid, "celo1ms");
        }
    }
}
