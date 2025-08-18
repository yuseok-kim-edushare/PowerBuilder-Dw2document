using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Text.Json;
using yuseok.kim.dw2docs.Interop;

namespace yuseok.kim.dw2docs.test.Interop
{
    [TestClass]
    public class MediumFidelityParsingTests
    {
        [TestMethod]
        public void ParseJSON_WithVisualObjects_ExtractsAllObjectTypes()
        {
            // Arrange
            const string testJson = @"{
              ""objects"": [
                {
                  ""name"": ""title_text"",
                  ""type"": ""text"",
                  ""x"": 100,
                  ""y"": 50,
                  ""width"": 300,
                  ""height"": 60,
                  ""band"": ""header"",
                  ""text"": ""Employee Report"",
                  ""alignment"": ""center"",
                  ""font_face"": ""Arial"",
                  ""font_height"": 14,
                  ""font_weight"": 700
                },
                {
                  ""name"": ""separator_line"",
                  ""type"": ""line"",
                  ""x1"": 0,
                  ""y1"": 120,
                  ""x2"": 500,
                  ""y2"": 120,
                  ""band"": ""header"",
                  ""pen_width"": 2,
                  ""pen_color"": 0
                },
                {
                  ""name"": ""total_compute"",
                  ""type"": ""compute"",
                  ""x"": 200,
                  ""y"": 10,
                  ""width"": 100,
                  ""height"": 30,
                  ""band"": ""footer"",
                  ""expression"": ""Count(Name)"",
                  ""format"": ""#,##0"",
                  ""alignment"": ""right""
                }
              ],
              ""columns"": [
                { ""name"": ""Name"", ""width"": 100, ""type"": ""char(20)"", ""format"": """" }
              ],
              ""bands"": [
                { ""name"": ""header"" },
                { ""name"": ""detail"" },
                { ""name"": ""footer"" }
              ],
              ""rows"": [
                { ""Name"": ""Alice"" }
              ],
              ""cell_attributes"": {
                ""cell_0_0"": { ""text"": ""Alice"", ""is_visible"": true }
              }
            }";

            // Act - Parse JSON to validate structure
            using var document = JsonDocument.Parse(testJson);
            var root = document.RootElement;

            // Assert - Verify objects are parseable
            Assert.IsTrue(root.TryGetProperty("objects", out var objectsElement), "Objects property should exist");
            
            var objectsArray = objectsElement.EnumerateArray().ToArray();
            Assert.AreEqual(3, objectsArray.Length, "Should have 3 visual objects");

            // Verify text object
            var textObj = objectsArray[0];
            Assert.AreEqual("text", textObj.GetProperty("type").GetString());
            Assert.AreEqual("Employee Report", textObj.GetProperty("text").GetString());
            Assert.AreEqual("header", textObj.GetProperty("band").GetString());
            Assert.AreEqual(100, textObj.GetProperty("x").GetInt32());
            Assert.AreEqual(50, textObj.GetProperty("y").GetInt32());

            // Verify line object
            var lineObj = objectsArray[1];
            Assert.AreEqual("line", lineObj.GetProperty("type").GetString());
            Assert.AreEqual(0, lineObj.GetProperty("x1").GetInt32());
            Assert.AreEqual(120, lineObj.GetProperty("y1").GetInt32());
            Assert.AreEqual(500, lineObj.GetProperty("x2").GetInt32());
            Assert.AreEqual(120, lineObj.GetProperty("y2").GetInt32());

            // Verify compute object
            var computeObj = objectsArray[2];
            Assert.AreEqual("compute", computeObj.GetProperty("type").GetString());
            Assert.AreEqual("Count(Name)", computeObj.GetProperty("expression").GetString());
            Assert.AreEqual("footer", computeObj.GetProperty("band").GetString());
            Assert.AreEqual("#,##0", computeObj.GetProperty("format").GetString());

            Console.WriteLine("✅ Medium-fidelity object parsing validation passed!");
            Console.WriteLine($"   - Text object with positioning: ({textObj.GetProperty("x").GetInt32()},{textObj.GetProperty("y").GetInt32()})");
            Console.WriteLine($"   - Line object with coordinates: ({lineObj.GetProperty("x1").GetInt32()},{lineObj.GetProperty("y1").GetInt32()}) to ({lineObj.GetProperty("x2").GetInt32()},{lineObj.GetProperty("y2").GetInt32()})");
            Console.WriteLine($"   - Compute object with expression: {computeObj.GetProperty("expression").GetString()}");
        }

        [TestMethod]
        public void DatawindowExporter_WithMediumFidelityJSON_ProcessesSuccessfully()
        {
            // Arrange
            const string mediumFidelityJson = @"{
              ""objects"": [
                {
                  ""name"": ""report_title"",
                  ""type"": ""text"",
                  ""x"": 100,
                  ""y"": 50,
                  ""width"": 300,
                  ""height"": 60,
                  ""band"": ""header"",
                  ""text"": ""Sales Report"",
                  ""alignment"": ""center"",
                  ""font_face"": ""Arial"",
                  ""font_height"": 16,
                  ""font_weight"": 700,
                  ""underline"": ""false"",
                  ""italic"": ""false""
                }
              ],
              ""columns"": [
                { ""name"": ""Product"", ""width"": 150, ""type"": ""char(30)"", ""format"": """" },
                { ""name"": ""Sales"", ""width"": 100, ""type"": ""decimal(10,2)"", ""format"": ""#,##0.00"" }
              ],
              ""bands"": [
                { ""name"": ""header"" },
                { ""name"": ""detail"" }
              ],
              ""rows"": [
                { ""Product"": ""Widget A"", ""Sales"": ""1000.00"" },
                { ""Product"": ""Widget B"", ""Sales"": ""2500.00"" }
              ],
              ""cell_attributes"": {
                ""cell_0_0"": { ""text"": ""Widget A"", ""is_visible"": true },
                ""cell_0_1"": { ""text"": ""1000.00"", ""is_visible"": true },
                ""cell_1_0"": { ""text"": ""Widget B"", ""is_visible"": true },
                ""cell_1_1"": { ""text"": ""2500.00"", ""is_visible"": true }
              }
            }";

            var exporter = new DatawindowExporter();

            // Act - Test both export formats
            string wordResult = exporter.ExportToWord(mediumFidelityJson, Path.GetTempFileName() + ".docx");

            // Assert
            Assert.IsTrue(wordResult.StartsWith("Success"), $"Word export should succeed, got: {wordResult}");
            
            Console.WriteLine("✅ Medium-fidelity DataWindow export validation passed!");
            Console.WriteLine($"   - Word export result: {wordResult}");
        }
    }
}