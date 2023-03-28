using IronXL;
using IronXL.Styles;
using Schedule.Maker.Models.Entity.Excel;

namespace Schedule.Maker.Models.Helper.Library
{
    /// <summary>
    ///  Helper class used to store functions that used the IronXL library.
    /// </summary>
    internal abstract class LibraryHelper
    {
        /**
         * <summary>
         *  Function used to apply borders using IronXL.
         * </summary>
         * <param name="worksheet">The worksheet that is currently edited.</param>
         * <param name="cells">The list of cells to treat.</param>
         */
        internal static void Apply_Borders(WorkSheet worksheet, List<ExcelCell> cells)
        {
            foreach(ExcelCell cell in cells)
            {
                if (cell.borders.Contains("TOP")) worksheet[$"{cell.range}"].Style.TopBorder.Type = BorderType.Thin;

                if (cell.borders.Contains("BOTTOM")) worksheet[$"{cell.range}"].Style.BottomBorder.Type = BorderType.Thin;

                if (cell.borders.Contains("LEFT")) worksheet[$"{cell.range}"].Style.LeftBorder.Type = BorderType.Thin;
                
                if(cell.borders.Contains("RIGHT")) worksheet[$"{cell.range}"].Style.RightBorder.Type = BorderType.Thin;
            }
        }

        /**
         * <summary>
         *  Function used to apply borders using IronXL.
         * </summary>
         * <param name="worksheet">The worksheet that is currently edited.</param>
         * <param name="cells">The list of cells to treat.</param>
         */
        internal static void Apply_Merge(WorkSheet worksheet, List<ExcelCell> cells)
        {
            foreach(ExcelCell cell in cells)
            {
                worksheet.Merge($"{cell.range}");

                worksheet[$"{cell.range}"].Style.SetBackgroundColor(cell.color);
            }
        }
    }
}
