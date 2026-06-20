using Microsoft.VisualStudio.TestTools.UnitTesting;
using yuseok.kim.dw2docs.Common.Positioning;
using System.Drawing;

namespace yuseok.kim.dw2docs.test.Common.Positioning
{
    [TestClass]
    public class CoordinateMapperTests
    {
        private CoordinateMapper _coordinateMapper = null!;

        [TestInitialize]
        public void Setup()
        {
            _coordinateMapper = new CoordinateMapper();
        }

        [TestMethod]
        public void ConvertToPdfUnits_PowerBuilderUnitsToPoints_CorrectConversion()
        {
            // Arrange
            int pbUnits = 1000; // 1 inch in PowerBuilder units

            // Act
            double pdfPoints = _coordinateMapper.ConvertToPdfUnits(pbUnits);

            // Assert
            Assert.AreEqual(72.0, pdfPoints, 0.001, "1000 PB units (1 inch) should equal 72 PDF points");
        }

        [TestMethod]
        public void ConvertToWordUnits_PowerBuilderUnitsToTwips_CorrectConversion()
        {
            // Arrange
            int pbUnits = 1000; // 1 inch in PowerBuilder units

            // Act
            double wordTwips = _coordinateMapper.ConvertToWordUnits(pbUnits);

            // Assert
            Assert.AreEqual(1440.0, wordTwips, 0.001, "1000 PB units (1 inch) should equal 1440 Word twips");
        }

        [TestMethod]
        public void ConvertToPixels_PowerBuilderUnitsToPixels_CorrectConversion()
        {
            // Arrange
            int pbUnits = 1000; // 1 inch in PowerBuilder units
            double dpi = 96.0; // Standard screen DPI

            // Act
            double pixels = _coordinateMapper.ConvertToPixels(pbUnits, dpi);

            // Assert
            Assert.AreEqual(96.0, pixels, 0.001, "1000 PB units (1 inch) at 96 DPI should equal 96 pixels");
        }

        [TestMethod]
        public void ConvertToEMU_PowerBuilderUnitsToEMU_CorrectConversion()
        {
            // Arrange
            int pbUnits = 1000; // 1 inch in PowerBuilder units

            // Act
            long emuUnits = _coordinateMapper.ConvertToEMU(pbUnits);

            // Assert
            Assert.AreEqual(914400L, emuUnits, "1000 PB units (1 inch) should equal 914400 EMU");
        }

        [TestMethod]
        public void ConvertFromPixels_PixelsToPowerBuilderUnits_CorrectConversion()
        {
            // Arrange
            double pixels = 96.0; // 1 inch at 96 DPI
            double dpi = 96.0;

            // Act
            int pbUnits = _coordinateMapper.ConvertFromPixels(pixels, dpi);

            // Assert
            Assert.AreEqual(1000, pbUnits, "96 pixels at 96 DPI should equal 1000 PB units");
        }

        [TestMethod]
        public void ConvertFromPoints_PointsToPowerBuilderUnits_CorrectConversion()
        {
            // Arrange
            double points = 72.0; // 1 inch in points

            // Act
            int pbUnits = _coordinateMapper.ConvertFromPoints(points);

            // Assert
            Assert.AreEqual(1000, pbUnits, "72 points (1 inch) should equal 1000 PB units");
        }

        [TestMethod]
        public void ConvertFromTwips_TwipsToPowerBuilderUnits_CorrectConversion()
        {
            // Arrange
            double twips = 1440.0; // 1 inch in twips

            // Act
            int pbUnits = _coordinateMapper.ConvertFromTwips(twips);

            // Assert
            Assert.AreEqual(1000, pbUnits, "1440 twips (1 inch) should equal 1000 PB units");
        }

        [TestMethod]
        public void RoundTripConversion_PdfUnits_MaintainsAccuracy()
        {
            // Arrange
            int originalPbUnits = 2500; // 2.5 inches

            // Act
            double pdfPoints = _coordinateMapper.ConvertToPdfUnits(originalPbUnits);
            int convertedBack = _coordinateMapper.ConvertFromPoints(pdfPoints);

            // Assert
            Assert.AreEqual(originalPbUnits, convertedBack, "Round-trip conversion should maintain accuracy");
        }

        [TestMethod]
        public void RoundTripConversion_WordUnits_MaintainsAccuracy()
        {
            // Arrange
            int originalPbUnits = 3600; // 3.6 inches

            // Act
            double wordTwips = _coordinateMapper.ConvertToWordUnits(originalPbUnits);
            int convertedBack = _coordinateMapper.ConvertFromTwips(wordTwips);

            // Assert
            Assert.AreEqual(originalPbUnits, convertedBack, "Round-trip conversion should maintain accuracy");
        }
    }
}