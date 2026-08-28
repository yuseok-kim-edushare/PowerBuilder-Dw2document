using System;
using System.Drawing;

namespace yuseok.kim.dw2docs.Common.Positioning
{
    /// <summary>
    /// Represents a document color with RGB values
    /// </summary>
    public class DocumentColor
    {
        public int Red { get; }
        public int Green { get; }
        public int Blue { get; }
        public int Alpha { get; }

        public DocumentColor(int red, int green, int blue, int alpha = 255)
        {
            Red = Math.Max(0, Math.Min(255, red));
            Green = Math.Max(0, Math.Min(255, green));
            Blue = Math.Max(0, Math.Min(255, blue));
            Alpha = Math.Max(0, Math.Min(255, alpha));
        }

        public override string ToString()
        {
            return $"RGB({Red}, {Green}, {Blue})";
        }

        public string ToHex()
        {
            return $"#{Red:X2}{Green:X2}{Blue:X2}";
        }
    }

    /// <summary>
    /// Represents a document font with metrics and styling
    /// </summary>
    public class DocumentFont
    {
        public string FontFace { get; }
        public double PointSize { get; }
        public bool IsBold { get; }
        public bool IsItalic { get; }
        public bool IsUnderline { get; }
        public bool IsStrikethrough { get; }

        public DocumentFont(string fontFace, double pointSize, bool isBold = false, 
            bool isItalic = false, bool isUnderline = false, bool isStrikethrough = false)
        {
            FontFace = fontFace ?? "Arial";
            PointSize = Math.Max(1.0, pointSize);
            IsBold = isBold;
            IsItalic = isItalic;
            IsUnderline = isUnderline;
            IsStrikethrough = isStrikethrough;
        }

        public override string ToString()
        {
            return $"{FontFace} {PointSize}pt" + 
                   (IsBold ? " Bold" : "") + 
                   (IsItalic ? " Italic" : "") +
                   (IsUnderline ? " Underline" : "") +
                   (IsStrikethrough ? " Strikethrough" : "");
        }
    }

    /// <summary>
    /// Maps PowerBuilder styling properties to document format styling
    /// </summary>
    public class StyleMapper
    {
        /// <summary>
        /// Maps PowerBuilder color integer to DocumentColor
        /// PowerBuilder stores color as BGR format in a 32-bit integer
        /// </summary>
        /// <param name="pbColor">PowerBuilder color integer</param>
        /// <returns>DocumentColor object</returns>
        public DocumentColor MapColor(int pbColor)
        {
            // PowerBuilder color format is BGR (Blue-Green-Red)
            // Extract RGB values from PowerBuilder color integer
            int r = (pbColor & 0xFF0000) >> 16;
            int g = (pbColor & 0x00FF00) >> 8;
            int b = (pbColor & 0x0000FF);
            
            return new DocumentColor(r, g, b);
        }

        /// <summary>
        /// Maps System.Drawing.Color to DocumentColor
        /// </summary>
        /// <param name="color">System.Drawing.Color</param>
        /// <returns>DocumentColor object</returns>
        public DocumentColor MapColor(Color color)
        {
            return new DocumentColor(color.R, color.G, color.B, color.A);
        }

        /// <summary>
        /// Maps PowerBuilder font properties to DocumentFont
        /// </summary>
        /// <param name="fontFace">Font face name</param>
        /// <param name="heightPBUnits">Font height in PowerBuilder units (1/1000 inch)</param>
        /// <param name="weight">Font weight (400=normal, 700=bold)</param>
        /// <param name="isItalic">Italic flag</param>
        /// <param name="isUnderline">Underline flag</param>
        /// <param name="isStrikethrough">Strikethrough flag</param>
        /// <returns>DocumentFont object</returns>
        public DocumentFont MapFont(string? fontFace, int heightPBUnits, int weight, 
            bool isItalic = false, bool isUnderline = false, bool isStrikethrough = false)
        {
            // Convert PB height (1/1000 inch) to points (1/72 inch)
            double pointSize = heightPBUnits * (72.0 / 1000.0);
            
            // Map weight (400=normal, 700=bold, but can be any value)
            bool isBold = weight >= 700;
            
            return new DocumentFont(fontFace ?? "Arial", pointSize, isBold, 
                isItalic, isUnderline, isStrikethrough);
        }

        /// <summary>
        /// Maps PowerBuilder font properties using byte values
        /// </summary>
        /// <param name="fontFace">Font face name</param>
        /// <param name="fontSize">Font size in points (already converted)</param>
        /// <param name="weight">Font weight as short</param>
        /// <param name="isItalic">Italic flag</param>
        /// <param name="isUnderline">Underline flag</param>
        /// <param name="isStrikethrough">Strikethrough flag</param>
        /// <returns>DocumentFont object</returns>
        public DocumentFont MapFont(string? fontFace, byte fontSize, short weight, 
            bool isItalic = false, bool isUnderline = false, bool isStrikethrough = false)
        {
            // Font size is already in points
            double pointSize = fontSize;
            
            // Map weight (400=normal, 700=bold)
            bool isBold = weight >= 700;
            
            return new DocumentFont(fontFace ?? "Arial", pointSize, isBold, 
                isItalic, isUnderline, isStrikethrough);
        }

        /// <summary>
        /// Gets a suitable fallback font for the given font face
        /// </summary>
        /// <param name="fontFace">Original font face</param>
        /// <returns>Fallback font face name</returns>
        public string GetFallbackFont(string? fontFace)
        {
            if (string.IsNullOrEmpty(fontFace))
                return "Arial";

            // Common font substitutions
            var lowerFont = fontFace.ToLowerInvariant();
            return lowerFont switch
            {
                var f when f.Contains("times") || f.Contains("serif") => "Times New Roman",
                var f when f.Contains("courier") || f.Contains("mono") => "Courier New",
                var f when f.Contains("helvetica") => "Arial",
                var f when f.Contains("comic") => "Comic Sans MS",
                var f when f.Contains("impact") => "Impact",
                var f when f.Contains("georgia") => "Georgia",
                var f when f.Contains("verdana") => "Verdana",
                var f when f.Contains("tahoma") => "Tahoma",
                _ => "Arial" // Default fallback
            };
        }

        /// <summary>
        /// Calculates color difference using Delta E formula (simplified)
        /// Values less than 2.0 are considered visually equivalent
        /// </summary>
        /// <param name="color1">First color</param>
        /// <param name="color2">Second color</param>
        /// <returns>Color difference value</returns>
        public double CalculateColorDifference(DocumentColor color1, DocumentColor color2)
        {
            // Simplified Delta E calculation using RGB values
            double deltaR = color1.Red - color2.Red;
            double deltaG = color1.Green - color2.Green;
            double deltaB = color1.Blue - color2.Blue;
            
            return Math.Sqrt(deltaR * deltaR + deltaG * deltaG + deltaB * deltaB);
        }
    }
}