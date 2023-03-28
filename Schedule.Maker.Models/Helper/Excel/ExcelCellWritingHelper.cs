using IronXL;
using IronXL.Styles;
using Schedule.Maker.Models.Entity.Excel;

namespace Schedule.Maker.Models.Helper.Excel
{
    public abstract class ExcelCellWritingHelper
    {
        static ExcelConstant excel_constant = new ExcelConstant();

        /**
         * <summary>
         *  Function used to write the content of a given cell.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         */
        internal static void Write_Cell_Content(IronXL.Range range, string value, string format = "", string color = "")
        {
            if (format != "")
            {
                range.FormatString = format;
            }

            if (color != "")
            {
                range.Style.BackgroundColor = color;
                range.Style.Font.Color = "#FFFFFF";
            }

            range.Value = value;

            range.Style.Font.Name = "Consolas";
            range.Style.Font.Height = 8;
            range.Style.Font.Bold = true;
            range.Style.HorizontalAlignment = HorizontalAlignment.Center;
            range.Style.VerticalAlignment = VerticalAlignment.Center;
        }

        /**
         * <summary>
         *  Function used to write reference hours of the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         */
        internal static void Write_Reference_Hour_Values(WorkSheet worksheet, int index)
        {
            Write_Cell_Content(worksheet[$"A{index}"], excel_constant.hour_reference_title, excel_constant.hour_format);

            Write_Cell_Content(worksheet[$"C{index}"], excel_constant.morning_start, excel_constant.hour_format);
            Write_Cell_Content(worksheet[$"D{index}"], excel_constant.morning_end, excel_constant.hour_format);

            Write_Cell_Content(worksheet[$"E{index}"], excel_constant.afternoon_start, excel_constant.hour_format);
            Write_Cell_Content(worksheet[$"F{index}"], excel_constant.afternoon_end, excel_constant.hour_format);

            Write_Cell_Content(worksheet[$"H{index}"], excel_constant.hour_total, excel_constant.hour_format);
        }

        /**
         * <summary>
         *  Function used to write schedule section of the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         */
        internal static void Write_Schedule_Section_Values(WorkSheet worksheet, int index)
        {
            Write_Cell_Content(worksheet[$"C{index}"], excel_constant.morning_section);
            Write_Cell_Content(worksheet[$"E{index}"], excel_constant.afternoon_section);
        }

        /**
         * <summary>
         *  Function used to write weekend line in the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         * <param name="comment">The comment to add on the line</param>
         * <param name="day_number">The number corresponding to the current day of the month.</param>
         */
        internal static void Write_Non_Working_Values(WorkSheet worksheet, int index, int day_number, string comment)
        {
            Write_Cell_Content(worksheet[$"A{index}"], $"{day_number}");
            Write_Cell_Content(worksheet[$"G{index}"], comment);
            Write_Cell_Content(worksheet[$"H{index}"], excel_constant.hour_none_total, excel_constant.hour_format, excel_constant.invalid_color);
        }

        /**
         * <summary>
         *  Function used to write a working line in the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         * <param name="comment">The comment to add on the line</param>
         * <param name="day_number">The number corresponding to the current day of the month.</param>
         * <param name="isRemote">A boolean valu that indicates if a given day is a remote day or not.</param>
         */
        internal static void Write_Working_Values(WorkSheet worksheet, int index, int day_number, bool isRemote, string comment)
        {
            Write_Cell_Content(worksheet[$"A{index}"], $"{day_number}");

            if (isRemote)
            {
                Write_Cell_Content(worksheet[$"B{index}"], $"OUI", "", excel_constant.valid_color);
            }
            else
            {
                Write_Cell_Content(worksheet[$"B{index}"], $"NON", "", excel_constant.invalid_color);
            }

            Write_Cell_Content(worksheet[$"C{index}"], excel_constant.morning_start_user, "hh:mm");
            Write_Cell_Content(worksheet[$"D{index}"], excel_constant.morning_end_user, "hh:mm");
            Write_Cell_Content(worksheet[$"E{index}"], excel_constant.afternoon_start_user, "hh:mm");
            Write_Cell_Content(worksheet[$"F{index}"], excel_constant.afternoon_end_user, "hh:mm");
            Write_Cell_Content(worksheet[$"G{index}"], comment);
            Write_Cell_Content(worksheet[$"H{index}"], excel_constant.hour_total, "hh:mm", excel_constant.valid_color);
        }

        /**
         * <summary>
         *  Function used to write the signature line of the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         */
        internal static void Write_Signature_Values(WorkSheet worksheet, int index)
        {
            Write_Cell_Content(worksheet[$"C{index}"], excel_constant.signature_label);
        }

        /**
         * <summary>
         *  Function used to write header columns of the file.
         * </summary>
         * <param name="worksheet">The worksheet that is beeing updated.</param>
         * <param name="index">The current line that has to be written.</param>
         */
        internal static void Write_Header_Column_Values(WorkSheet worksheet, int index)
        {
            Write_Cell_Content(worksheet[$"A{index}"], excel_constant.day_column);
            Write_Cell_Content(worksheet[$"B{index}"], excel_constant.remote_column);
            Write_Cell_Content(worksheet[$"C{index}"], excel_constant.start_column);
            Write_Cell_Content(worksheet[$"D{index}"], excel_constant.end_column);
            Write_Cell_Content(worksheet[$"E{index}"], excel_constant.start_column);
            Write_Cell_Content(worksheet[$"F{index}"], excel_constant.end_column);
            Write_Cell_Content(worksheet[$"G{index}"], excel_constant.comment_column);
            Write_Cell_Content(worksheet[$"H{index}"], excel_constant.hour_summary_column);
        }
    }
}
