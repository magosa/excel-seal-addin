namespace SealAddIn.Services
{
    /// <summary>
    /// mm / pt / px の単位変換ユーティリティ。Excelの図形サイズはポイント(pt)基準です。
    /// </summary>
    public static class UnitConversion
    {
        private const double PointsPerMm = 72.0 / 25.4; // 1mm = 2.83464566...pt
        private const double PointsPerPx = 72.0 / 96.0; // 96dpi前提

        public static double MmToPoints(double mm) => mm * PointsPerMm;

        public static double PointsToMm(double pt) => pt / PointsPerMm;

        public static double PixelsToPoints(double px) => px * PointsPerPx;

        public static double PointsToPixels(double pt) => pt / PointsPerPx;

        public static double ConvertToPoints(double value, string unit)
        {
            switch (unit)
            {
                case "mm":
                    return MmToPoints(value);
                case "px":
                    return PixelsToPoints(value);
                default:
                    return value; // pt
            }
        }

        public static double ConvertFromPoints(double pt, string unit)
        {
            switch (unit)
            {
                case "mm":
                    return PointsToMm(pt);
                case "px":
                    return PointsToPixels(pt);
                default:
                    return pt;
            }
        }
    }
}
