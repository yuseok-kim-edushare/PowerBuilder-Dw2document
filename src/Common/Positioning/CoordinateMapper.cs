using System;

namespace yuseok.kim.dw2docs.Common.Positioning
{
    /// <summary>
    /// Provides coordinate mapping between PowerBuilder units and various document format units
    /// PowerBuilder uses 1/1000 inch units, this class handles conversion to target formats
    /// </summary>
    public class CoordinateMapper
    {
        /// <summary>
        /// Converts PowerBuilder units (1/1000 inch) to PDF points (1/72 inch)
        /// </summary>
        /// <param name="pbUnits">PowerBuilder units</param>
        /// <returns>PDF points</returns>
        public double ConvertToPdfUnits(int pbUnits)
        {
            // PowerBuilder uses 1/1000 inch, PDF uses 1/72 inch (points)
            return pbUnits * (72.0 / 1000.0);
        }

        /// <summary>
        /// Converts PowerBuilder units (1/1000 inch) to Word twips (1/1440 inch)
        /// </summary>
        /// <param name="pbUnits">PowerBuilder units</param>
        /// <returns>Word twips</returns>
        public double ConvertToWordUnits(int pbUnits)
        {
            // Word uses twips (1/1440 inch)
            return pbUnits * (1440.0 / 1000.0);
        }

        /// <summary>
        /// Converts PowerBuilder units to pixels based on DPI
        /// </summary>
        /// <param name="pbUnits">PowerBuilder units</param>
        /// <param name="dpi">Dots per inch (default 96 DPI for screen)</param>
        /// <returns>Pixels</returns>
        public double ConvertToPixels(int pbUnits, double dpi = 96.0)
        {
            // PowerBuilder uses 1/1000 inch, convert to inches then to pixels
            double inches = pbUnits / 1000.0;
            return inches * dpi;
        }

        /// <summary>
        /// Converts PowerBuilder units to Excel EMU (English Metric Units)
        /// </summary>
        /// <param name="pbUnits">PowerBuilder units</param>
        /// <returns>EMU units</returns>
        public long ConvertToEMU(int pbUnits)
        {
            // PowerBuilder uses 1/1000 inch, EMU uses 914400 per inch
            return (long)(pbUnits * (914400.0 / 1000.0));
        }

        /// <summary>
        /// Converts pixels back to PowerBuilder units
        /// </summary>
        /// <param name="pixels">Pixel value</param>
        /// <param name="dpi">Dots per inch (default 96 DPI for screen)</param>
        /// <returns>PowerBuilder units</returns>
        public int ConvertFromPixels(double pixels, double dpi = 96.0)
        {
            // Convert pixels to inches, then to PowerBuilder units
            double inches = pixels / dpi;
            return (int)Math.Round(inches * 1000.0);
        }

        /// <summary>
        /// Converts points to PowerBuilder units
        /// </summary>
        /// <param name="points">Point value (1/72 inch)</param>
        /// <returns>PowerBuilder units</returns>
        public int ConvertFromPoints(double points)
        {
            // Points are 1/72 inch, PowerBuilder is 1/1000 inch
            return (int)Math.Round(points * (1000.0 / 72.0));
        }

        /// <summary>
        /// Converts twips to PowerBuilder units
        /// </summary>
        /// <param name="twips">Twip value (1/1440 inch)</param>
        /// <returns>PowerBuilder units</returns>
        public int ConvertFromTwips(double twips)
        {
            // Twips are 1/1440 inch, PowerBuilder is 1/1000 inch
            return (int)Math.Round(twips * (1000.0 / 1440.0));
        }
    }
}