using IronXL;
using Schedule.Maker.Models.Entity.Data;
using Schedule.Maker.Models.Helper.Excel;

namespace Schedule.Maker.Models.Entity.Excel
{
    /// <summary>
    ///  Class that contains excel variables & excel constants.
    /// </summary>
    public class ExcelEntity
    {
        User user { get; set; }
        public WorkBook workbook { get; set; }
        private WorkSheet worksheet { get; set; }
        private int index { get; set; }

        private const string weekend_comment = "WEEK-END";
        private const string days_off_comment = "CONGES";
        private const string public_day_comment = "FERIE";
        private const string none_comment = "-";

        public ExcelEntity(User user)
        {
            this.user = user;

            this.workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            this.worksheet = this.workbook.CreateWorkSheet($"CRA-{this.user.month}");

            this.index = 1;
        }

        /**
         * <summary>
         *  Function used to write the excel header that contains username, current month, reference hours, schedule sections, column names.
         * </summary>
         */
        public void Write_Excel_Header()
        {
            index = ExcelStyleHelper.Write_Username_Style(this.worksheet, index);
            ExcelCellWritingHelper.Write_Cell_Content(this.worksheet["D1"], $"{this.user.first_name} {this.user.last_name}");

            index = ExcelStyleHelper.Write_Generated_Month_Style(this.worksheet, index);
            ExcelCellWritingHelper.Write_Cell_Content(this.worksheet["D2"], $"{this.user.month.ToUpper()} {DateTime.Now.Year}", "MMMM yyyy");

            index = ExcelStyleHelper.Write_Reference_Hour_Style(this.worksheet, index);
            ExcelCellWritingHelper.Write_Reference_Hour_Values(this.worksheet, index - 1);

            index = ExcelStyleHelper.Write_Schedule_Section_Style(this.worksheet, index);
            ExcelCellWritingHelper.Write_Schedule_Section_Values(this.worksheet, index - 1);

            index = ExcelStyleHelper.Write_Header_Column_Style(this.worksheet, index);
            ExcelCellWritingHelper.Write_Header_Column_Values(this.worksheet, index - 1);
        }

        /**
         * <summary>
         *  Function used to write the excel main content that contains all the working hours.
         * </summary>
         * <param name="days_in_month">The list that contains all the month of the generated month.</param>
         * <param name="user">The user object that contains all the data.</param>
         */
        public void Write_Excel_Content(List<DateTime> days_in_month, User user)
        {
            for (int i = 0; i < days_in_month.Count; i++)
            {
                if (days_in_month[i].DayOfWeek == DayOfWeek.Sunday || days_in_month[i].DayOfWeek == DayOfWeek.Saturday)
                {
                    index = ExcelStyleHelper.Write_Non_Working_Style(this.worksheet, index);
                    ExcelCellWritingHelper.Write_Non_Working_Values(this.worksheet, index - 1, i + 1, weekend_comment);
                }

                else if (user.days_off.Contains(days_in_month[i]))
                {
                    index = ExcelStyleHelper.Write_Non_Working_Style(this.worksheet, index);
                    ExcelCellWritingHelper.Write_Non_Working_Values(this.worksheet, index - 1, i + 1, days_off_comment);
                }

                else if (user.public_days.Contains(days_in_month[i]))
                {
                    index = ExcelStyleHelper.Write_Non_Working_Style(this.worksheet, index);
                    ExcelCellWritingHelper.Write_Non_Working_Values(this.worksheet, index - 1, i + 1, public_day_comment);
                }

                else if (user.remote_days.Contains(days_in_month[i]))
                {
                    index = ExcelStyleHelper.Write_Working_Style(this.worksheet, index);
                    ExcelCellWritingHelper.Write_Working_Values(this.worksheet, index - 1, i + 1, true, none_comment);
                }

                else if (!user.remote_days.Contains(days_in_month[i]))
                {
                    index = ExcelStyleHelper.Write_Working_Style(this.worksheet, index);
                    ExcelCellWritingHelper.Write_Working_Values(this.worksheet, index - 1, i + 1, false, none_comment);
                }
            }
        }

        /**
         * <summary>
         *  Function used to write the excel footer that contains the owner signature.
         * </summary>
         */
        public void Write_Excel_Footer()
        {
            ExcelStyleHelper.Write_Signature_Style(this.worksheet, index);
            ExcelCellWritingHelper.Write_Signature_Values(this.worksheet, index);
        }
    }
}
