using System;
using Helper;
using NUnit.Framework;

namespace Tests
{
    public class ExcelHelperTest
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void Test1()
        {
            var dynamics = ExcelHelper.ReadExcel("", 0, false);

            foreach (var item in dynamics)
            {
                Console.WriteLine(item);
            }
            Assert.AreNotEqual(null, dynamics);
        }
    }
}