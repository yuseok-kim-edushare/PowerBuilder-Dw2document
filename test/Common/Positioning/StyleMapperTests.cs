using Microsoft.VisualStudio.TestTools.UnitTesting;
using yuseok.kim.dw2docs.Common.Positioning;
using System.Drawing;

namespace yuseok.kim.dw2docs.test.Common.Positioning
{
    [TestClass]
    public class StyleMapperTests
    {
        private StyleMapper _styleMapper = null!;

        [TestInitialize]
        public void Setup()
        {
            _styleMapper = new StyleMapper();
        }

        [TestMethod]
        public void MapColor_PowerBuilderColorInteger_CorrectRGBExtraction()
        {
            // Arrange - PowerBuilder stores color as BGR format
            int pbColor = 0x00FF0000; // Red in BGR format (Blue=0, Green=0, Red=255)

            // Act
            var documentColor = _styleMapper.MapColor(pbColor);

            // Assert
            Assert.AreEqual(255, documentColor.Red, "Red component should be extracted correctly");
            Assert.AreEqual(0, documentColor.Green, "Green component should be extracted correctly");
            Assert.AreEqual(0, documentColor.Blue, "Blue component should be extracted correctly");
        }

        [TestMethod]
        public void MapColor_SystemDrawingColor_CorrectConversion()
        {
            // Arrange
            var systemColor = Color.FromArgb(128, 255, 100, 50); // Alpha=128, R=255, G=100, B=50

            // Act
            var documentColor = _styleMapper.MapColor(systemColor);

            // Assert
            Assert.AreEqual(255, documentColor.Red);
            Assert.AreEqual(100, documentColor.Green);
            Assert.AreEqual(50, documentColor.Blue);
            Assert.AreEqual(128, documentColor.Alpha);
        }

        [TestMethod]
        public void MapFont_PowerBuilderHeightUnits_CorrectPointConversion()
        {
            // Arrange
            string fontFace = "Arial";
            int heightPBUnits = 250; // 0.25 inch in PowerBuilder units
            int weight = 700; // Bold weight

            // Act
            var documentFont = _styleMapper.MapFont(fontFace, heightPBUnits, weight);

            // Assert
            Assert.AreEqual("Arial", documentFont.FontFace);
            Assert.AreEqual(18.0, documentFont.PointSize, 0.001, "250 PB units should equal 18 points");
            Assert.IsTrue(documentFont.IsBold, "Weight 700 should be interpreted as bold");
        }

        [TestMethod]
        public void MapFont_ByteSizeWeight_CorrectMapping()
        {
            // Arrange
            string fontFace = "Times New Roman";
            byte fontSize = 12; // Already in points
            short weight = 400; // Normal weight
            bool isItalic = true;
            bool isUnderline = true;

            // Act
            var documentFont = _styleMapper.MapFont(fontFace, fontSize, weight, isItalic, isUnderline);

            // Assert
            Assert.AreEqual("Times New Roman", documentFont.FontFace);
            Assert.AreEqual(12.0, documentFont.PointSize);
            Assert.IsFalse(documentFont.IsBold, "Weight 400 should not be bold");
            Assert.IsTrue(documentFont.IsItalic);
            Assert.IsTrue(documentFont.IsUnderline);
        }

        [TestMethod]
        public void MapFont_NullFontFace_UsesDefaultFont()
        {
            // Arrange
            string? fontFace = null;
            int heightPBUnits = 180;
            int weight = 400;

            // Act
            var documentFont = _styleMapper.MapFont(fontFace, heightPBUnits, weight);

            // Assert
            Assert.AreEqual("Arial", documentFont.FontFace, "Null font face should default to Arial");
        }

        [TestMethod]
        public void GetFallbackFont_CommonFonts_CorrectSubstitution()
        {
            // Act & Assert
            Assert.AreEqual("Times New Roman", _styleMapper.GetFallbackFont("Times"));
            Assert.AreEqual("Times New Roman", _styleMapper.GetFallbackFont("serif"));
            Assert.AreEqual("Courier New", _styleMapper.GetFallbackFont("Courier"));
            Assert.AreEqual("Courier New", _styleMapper.GetFallbackFont("mono"));
            Assert.AreEqual("Arial", _styleMapper.GetFallbackFont("Helvetica"));
            Assert.AreEqual("Arial", _styleMapper.GetFallbackFont("UnknownFont"));
            Assert.AreEqual("Arial", _styleMapper.GetFallbackFont(null));
            Assert.AreEqual("Arial", _styleMapper.GetFallbackFont(""));
        }

        [TestMethod]
        public void CalculateColorDifference_SameColors_ZeroDifference()
        {
            // Arrange
            var color1 = new DocumentColor(255, 100, 50);
            var color2 = new DocumentColor(255, 100, 50);

            // Act
            double difference = _styleMapper.CalculateColorDifference(color1, color2);

            // Assert
            Assert.AreEqual(0.0, difference, 0.001, "Same colors should have zero difference");
        }

        [TestMethod]
        public void CalculateColorDifference_DifferentColors_PositiveDifference()
        {
            // Arrange
            var color1 = new DocumentColor(0, 0, 0); // Black
            var color2 = new DocumentColor(255, 255, 255); // White

            // Act
            double difference = _styleMapper.CalculateColorDifference(color1, color2);

            // Assert
            Assert.IsTrue(difference > 0, "Different colors should have positive difference");
            Assert.AreEqual(441.67, difference, 0.5, "Black to white should have maximum difference");
        }

        [TestMethod]
        public void DocumentColor_ToHex_CorrectHexString()
        {
            // Arrange
            var color = new DocumentColor(255, 128, 64);

            // Act
            string hex = color.ToHex();

            // Assert
            Assert.AreEqual("#FF8040", hex, "Hex conversion should be correct");
        }

        [TestMethod]
        public void DocumentColor_ColorClamping_ClampsBounds()
        {
            // Arrange & Act
            var color = new DocumentColor(-10, 300, 128, 400);

            // Assert
            Assert.AreEqual(0, color.Red, "Negative values should be clamped to 0");
            Assert.AreEqual(255, color.Green, "Values over 255 should be clamped to 255");
            Assert.AreEqual(128, color.Blue, "Valid values should remain unchanged");
            Assert.AreEqual(255, color.Alpha, "Alpha over 255 should be clamped to 255");
        }

        [TestMethod]
        public void DocumentFont_ToString_FormatsCorrectly()
        {
            // Arrange
            var font = new DocumentFont("Arial", 12.0, true, true, true, true);

            // Act
            string fontString = font.ToString();

            // Assert
            Assert.IsTrue(fontString.Contains("Arial"), "Should contain font face");
            Assert.IsTrue(fontString.Contains("12pt"), "Should contain point size");
            Assert.IsTrue(fontString.Contains("Bold"), "Should contain Bold");
            Assert.IsTrue(fontString.Contains("Italic"), "Should contain Italic");
            Assert.IsTrue(fontString.Contains("Underline"), "Should contain Underline");
            Assert.IsTrue(fontString.Contains("Strikethrough"), "Should contain Strikethrough");
        }
    }
}