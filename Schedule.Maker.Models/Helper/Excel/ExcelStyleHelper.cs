using IronXL;
using Schedule.Maker.Enums.Enum;
using Schedule.Maker.Models.Entity.Excel;
using Schedule.Maker.Models.Helper.Library;

namespace Schedule.Maker.Models.Helper.Excel
{
    /// <summary>
    ///  Helper class that contains functions concerning the writing of cell's style.
    /// </summary>
    internal abstract class ExcelStyleHelper
    {
        static ExcelConstant excel_constant = new ExcelConstant();

        /**
         * <summary>
         *  Function used to write the username of the person that will own the file after the generation.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         * <return>The updated index after the writing.</return>
         */
        internal static int Write_Username_Style(WorkSheet worksheet, int index)
        {
            string ranges = $"A{index}:A{index};A{index}:H{index};D{index}:E{index};H{index}:H{index}";
            string properties = $"{Border.LEFT};{Border.TOP};{Border.LEFT}|{Border.RIGHT}|{Border.BOTTOM};{Border.RIGHT}";

            List<ExcelCell> cells = GenerateCells(ranges, properties, GenerationType.BORDER);

            LibraryHelper.Apply_Borders(worksheet, cells);

            ranges = $"A{index}:C{index};D{index}:E{index};F{index}:H{index}";
            properties = $"{excel_constant.secondary};{excel_constant.primary};{excel_constant.secondary}";

            cells = GenerateCells(ranges, properties, GenerationType.MERGE);

            LibraryHelper.Apply_Merge(worksheet, cells);

            index++;

            return index;
        }

        /**
         * <summary>
         *  Function used to write the month and the date of the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         * <return>The updated index after the writing.</return>
         */
        internal static int Write_Generated_Month_Style(WorkSheet worksheet, int index)
        {
            string ranges = $"A{index}:A{index};A{index}:H{index};D{index}:E{index};H{index}:H{index}";
            string properties = $"{Border.LEFT};{Border.BOTTOM};{Border.LEFT}|{Border.RIGHT};{Border.RIGHT}";

            List<ExcelCell> cells = GenerateCells(ranges, properties, GenerationType.BORDER);

            LibraryHelper.Apply_Borders(worksheet, cells);

            ranges = $"A{index}:C{index};D{index}:E{index};F{index}:H{index}";
            properties = $"{excel_constant.secondary};{excel_constant.primary};{excel_constant.secondary}";

            cells = GenerateCells(ranges, properties, GenerationType.MERGE);

            LibraryHelper.Apply_Merge(worksheet, cells);

            index++;

            return index;
        }

        /**
         * <summary>
         *  Function used to write all the reference hours of the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         * <return>The updated index after the writing.</return>
         */
        internal static int Write_Reference_Hour_Style(WorkSheet worksheet, int index)
        {
            string ranges = $"A{index}:B{index};C{index}:H{index}";
            string properties = $"{Border.LEFT}|{Border.RIGHT}|{Border.BOTTOM};{Border.LEFT}|{Border.RIGHT}|{Border.BOTTOM}";

            List<ExcelCell> cells = GenerateCells(ranges, properties, GenerationType.BORDER);

            LibraryHelper.Apply_Borders(worksheet, cells);

            ranges = $"A{index}:B{index}";
            properties = $"{excel_constant.primary}";

            cells = GenerateCells(ranges, properties, GenerationType.MERGE);

            LibraryHelper.Apply_Merge(worksheet, cells);

            index++;

            return index;
        }

        /**
         * <summary>
         *  Function used to write all schedule section of the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         * <return>The updated index after the writing.</return>
         */
        internal static int Write_Schedule_Section_Style(WorkSheet worksheet, int index)
        {
            string ranges = $"A{index}:B{index};C{index}:D{index};E{index}:F{index};G{index}:H{index};A{index}:H{index}";
            string properties = $"${Border.LEFT}|{Border.RIGHT};{Border.LEFT}|{Border.RIGHT};{Border.LEFT}|{Border.RIGHT};{Border.LEFT}|{Border.RIGHT};{Border.BOTTOM}";

            List<ExcelCell> cells = GenerateCells(ranges, properties, GenerationType.BORDER);

            LibraryHelper.Apply_Borders(worksheet, cells);

            ranges = $"A{index}:B{index};C{index}:D{index};E{index}:F{index};G{index}:H{index}";
            properties = $"{excel_constant.secondary};{excel_constant.primary};{excel_constant.primary};{excel_constant.secondary}";

            cells = GenerateCells(ranges, properties, GenerationType.MERGE);

            LibraryHelper.Apply_Merge(worksheet, cells);

            index++;

            return index;
        }

        /**
         * <summary>
         *  Function used to write header columns of the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         * <return>The updated index after the writing.</return>
         */
        internal static int Write_Header_Column_Style(WorkSheet worksheet, int index)
        {
            string ranges = $"A{index}:H{index}";
            string properties = $"{Border.LEFT}|{Border.RIGHT}|{Border.BOTTOM}";

            List<ExcelCell> cells = GenerateCells(ranges, properties, GenerationType.BORDER);

            LibraryHelper.Apply_Borders(worksheet, cells);

            index++;

            return index;
        }

        /**
         * <summary>
         *  Function used to write a weekend line in the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         * <return>The updated index after the writing.</return>
         */
        internal static int Write_Non_Working_Style(WorkSheet worksheet, int index)
        {
            string ranges = $"A{index}:A{index};B{index}:F{index};G{index}:G{index};H{index}:H{index}";
            string properties = $"{Border.LEFT}|{Border.RIGHT}|{Border.BOTTOM};{Border.LEFT}|{Border.RIGHT};{Border.LEFT}|{Border.RIGHT}|{Border.BOTTOM};{Border.LEFT}|{Border.RIGHT}|{Border.BOTTOM}";

            List<ExcelCell> cells = GenerateCells(ranges, properties, GenerationType.BORDER);

            LibraryHelper.Apply_Borders(worksheet, cells);

            ranges = $"B{index}:F{index}";
            properties = $"{excel_constant.secondary}";

            cells = GenerateCells(ranges, properties, GenerationType.MERGE);

            LibraryHelper.Apply_Merge(worksheet, cells);

            index++;

            return index;
        }

        /**
         * <summary>
         *  Function used to write a working line in the file.
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         * </summary>
         */
        internal static int Write_Working_Style(WorkSheet worksheet, int index)
        {
            string ranges = $"A{index}:H{index}";
            string properties = $"{Border.LEFT}|{Border.RIGHT}|{Border.BOTTOM}|{Border.TOP}";

            List<ExcelCell> cells = GenerateCells(ranges, properties, GenerationType.BORDER);

            LibraryHelper.Apply_Borders(worksheet, cells);

            index++;

            return index;
        }

        /**
         * <summary>
         *  Function used to write the signature line of the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         */
        internal static void Write_Signature_Style(WorkSheet worksheet, int index)
        {
            string ranges = $"A{index}:H{index}";
            string properties = $"{Border.LEFT}|{Border.RIGHT}|{Border.BOTTOM}|{Border.TOP}";

            List<ExcelCell> cells = GenerateCells(ranges, properties, GenerationType.BORDER);

            LibraryHelper.Apply_Borders(worksheet, cells);

            ranges = $"A{index}:B{index};C{index}:D{index};E{index}:F{index};G{index}:H{index}";
            properties = $"{excel_constant.secondary};{excel_constant.primary};{excel_constant.primary};{excel_constant.secondary}";

            cells = GenerateCells(ranges, properties, GenerationType.MERGE);

            LibraryHelper.Apply_Merge(worksheet , cells);
        }

        /**
         * <summary>
         *  Function used to generate excell cells base on a string of ranges and properties.
         * </summary>
         * <param name="ranges">A string that contains all the ranges of cells.</param>
         * <param name="properties">A string that contains all the reanges of cells.</param>
         * <param name="type">The generation type to use.</param>
         * <return>cells - The list of cells that has been generated.</return>
         */
        private static List<ExcelCell> GenerateCells(string ranges, string properties, GenerationType type)
        {
            List<string> range_array = ranges.Split(';').ToList();
            List<string> property_array = properties.Split(";").ToList();

            List<ExcelCell> cells = new List<ExcelCell>();

            for (int i = 0; i < range_array.Count(); i++)
            {
                if (type == GenerationType.BORDER)
                    cells.Add(new ExcelCell(range_array.ElementAt(i), "", property_array.ElementAt(i)));

                if (type == GenerationType.MERGE)
                    cells.Add(new ExcelCell(range_array.ElementAt(i), property_array.ElementAt(i)));
            }

            return cells;
        }
    }
}
