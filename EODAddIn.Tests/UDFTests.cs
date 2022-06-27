using EODAddIn.BL;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace EODAddIn.Tests
{
    [TestClass]
    public class UDFTests
    {
        [TestMethod]
        public void EODOpen_AAPL()
        {
            // act
            UDF udf = new UDF();
            double? actual = udf.EOD_GetOpen("AAPL.US", new DateTime(2022, 03, 24));

            // assert
            Assert.AreEqual(171.06, actual);
        }

        [TestMethod]
        public void EODHigh_AAPL()
        {
            // act
            UDF udf = new UDF();
            double? actual = udf.EOD_GetHigh("AAPL.US", new DateTime(2022, 03, 24));

            // assert
            Assert.AreEqual(174.14, actual);
        }

        [TestMethod]
        public void EODLow_AAPL()
        {
            // act
            UDF udf = new UDF();
            double? actual = udf.EOD_GetLow("AAPL.US", new DateTime(2022, 03, 24));

            // assert
            Assert.AreEqual(170.21, actual);
        }

        [TestMethod]
        public void EODClose_AAPL()
        {
            // act
            UDF udf = new UDF();
            double? actual = udf.EOD_GetClose("AAPL.US", new DateTime(2022, 03, 24));

            // assert
            Assert.AreEqual(174.07, actual);
        }

        [TestMethod]
        public void EODAdjustedClose_AAPL()
        {
            // act
            UDF udf = new UDF();
            double? actual = udf.EOD_GetAdjustedClose("AAPL.US", new DateTime(2022, 03, 24));

            // assert
            Assert.AreEqual(174.07, actual);
        }

        [TestMethod]
        public void EODVolume_AAPL()
        {
            // act
            UDF udf = new UDF();
            double? actual = udf.EOD_GetVolume("AAPL.US", new DateTime(2022, 03, 24));

            // assert
            Assert.AreEqual(90131418, actual);
        }

        [TestMethod]
        public void IntradayOpen_AAPL()
        {
            // act
            UDF udf = new UDF();
            double? actual = udf.Intraday_GetOpen("AAPL.US", new DateTime(2022, 03, 14, 12, 01, 00));

            // assert
            Assert.AreEqual(154.2, actual);
        }

        [TestMethod]
        public void IntradayHigh_AAPL()
        {
            // act
            UDF udf = new UDF();
            double? actual = udf.Intraday_GetHigh("AAPL.US", new DateTime(2022, 03, 14, 12, 01, 00));

            // assert
            Assert.AreEqual(154.21, actual);
        }

        [TestMethod]
        public void IntradayLow_AAPL()
        {
            // act
            UDF udf = new UDF();
            double? actual = udf.Intraday_GetLow("AAPL.US", new DateTime(2022, 03, 14, 12, 01, 00));

            // assert
            Assert.AreEqual(154.2, actual);
        }

        [TestMethod]
        public void IntradayClose_AAPL()
        {
            // act
            UDF udf = new UDF();
            double? actual = udf.Intraday_GetClose("AAPL.US", new DateTime(2022, 03, 14, 12, 01, 00));

            // assert
            Assert.AreEqual(154.21, actual);
        }

        [TestMethod]
        public void IntradayVolume_AAPL()
        {
            // act
            UDF udf = new UDF();
            decimal? actual = udf.Intraday_GetVolume("AAPL.US", new DateTime(2022, 03, 14, 12, 01, 00));

            // assert
            Assert.AreEqual(947, actual);
        }
    }
}
