using NPOI.Util;
using yuseok.kim.dw2docs.Common.Positioning;

namespace yuseok.kim.dw2docs.Xlsx.Helpers;

internal static class UnitConversion
{
    private static readonly CoordinateMapper _coordinateMapper = new();

    public static int PixelsToXlsxCellUnits(int pixels) => pixels * 256 / 6;
    //public static int PixelsToXlsxCellUnits(int pixels) => pixels * 256 / 6;
    //public static double PixelsToCharacterFraction(int pixels) => pixels / 7.001699924468994 / 255;
    public static double PixelsToColumnWidth(int pixels) => Units.PixelToEMU(pixels) * 255.0 / 66691.0;
    public static int PixelsToEMU(int pixels) => Units.PixelToEMU(pixels);
    public static double PixelsToPoints(int pixels) => Units.PixelToPoints(pixels);

    /// <summary>
    /// Converts PowerBuilder units directly to pixels for XLSX operations
    /// </summary>
    /// <param name="pbUnits">PowerBuilder units (1/1000 inch)</param>
    /// <param name="dpi">Dots per inch (default 96 DPI)</param>
    /// <returns>Pixels</returns>
    public static double PowerBuilderUnitsToPixels(int pbUnits, double dpi = 96.0)
    {
        return _coordinateMapper.ConvertToPixels(pbUnits, dpi);
    }

    /// <summary>
    /// Converts PowerBuilder units to EMU for Excel operations
    /// </summary>
    /// <param name="pbUnits">PowerBuilder units (1/1000 inch)</param>
    /// <returns>EMU units</returns>
    public static long PowerBuilderUnitsToEMU(int pbUnits)
    {
        return _coordinateMapper.ConvertToEMU(pbUnits);
    }

    /// <summary>
    /// Converts PowerBuilder units to points for font sizing
    /// </summary>
    /// <param name="pbUnits">PowerBuilder units (1/1000 inch)</param>
    /// <returns>Points</returns>
    public static double PowerBuilderUnitsToPoints(int pbUnits)
    {
        return _coordinateMapper.ConvertToPdfUnits(pbUnits);
    }
}