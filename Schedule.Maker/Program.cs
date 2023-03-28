using Schedule.Maker.Models.Entity.Data;
using Schedule.Maker.Models.Entity.Excel;
using Schedule.Maker.Models.Helper.Terminal;
using Schedule.Maker.Models.Helper.TreeStructure;
using System.Globalization;

namespace Schedule.Maker
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // ***** Variable declarations ***** //

            bool is_asking = true;

            string excel_name = string.Empty;

            List<DateTime> days_in_month = new List<DateTime>();

            User user = new User(string.Empty, string.Empty, new List<DateTime>(), new List<DateTime>(), new List<DateTime>(), string.Empty, 0);

            // ***** User interactions & retrieval of user's information ***** //

            TerminalHelper.Write_Program_Information();
            TerminalHelper.Write_User_Instructions();

            while (is_asking)
            {
                string first_name = TerminalHelper.Ask_First_Name();
                string last_name = TerminalHelper.Ask_Last_Name();

                string month = TerminalHelper.Ask_Month();

                int month_number = DateTime.ParseExact($"{month}", "MMMM", new CultureInfo("en-US")).Month;

                excel_name = $"{month_number}.{month}_{DateTime.Now.Year}.xlsx";

                days_in_month = Enumerable.Range(1, DateTime.DaysInMonth(DateTime.Now.Year, month_number))
                                          .Select(day => new DateTime(DateTime.Now.Year, month_number, day, 0, 0, 0))
                                          .ToList();

                List<DateTime> rest_of_days = days_in_month.Where(day => day.DayOfWeek != DayOfWeek.Sunday && day.DayOfWeek != DayOfWeek.Saturday)
                                                            .ToList();

                List<DateTime> days_off = TerminalHelper.Ask_Days_Off(rest_of_days, month, month_number);

                rest_of_days = days_in_month.Where(day => !days_off.Contains(day) && day.DayOfWeek != DayOfWeek.Sunday && day.DayOfWeek != DayOfWeek.Saturday)
                                                           .ToList();

                List<DateTime> public_days = TerminalHelper.Ask_Public_Days(rest_of_days, month, month_number);

                rest_of_days = days_in_month.Where(day => !days_off.Contains(day) && !public_days.Contains(day) && day.DayOfWeek != DayOfWeek.Sunday && day.DayOfWeek != DayOfWeek.Saturday)
                                            .ToList();

                List<DateTime>  remote_days = TerminalHelper.Ask_Remote_Days(rest_of_days, month, month_number);

                user = new User(first_name, last_name, days_off, public_days, remote_days, month, month_number);

                is_asking = TerminalHelper.Write_User_Summary(user, $"[dim]{month_number}.[/][bold springgreen2]{month}_{DateTime.Now.Year}.xlsx[/]");
            }


            // ***** Tree structure establishment ***** //

            string base_directory_path = $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\schedule-maker-files";

            if (!TreeStructureHelper.Find_Directory(base_directory_path))
                Directory.CreateDirectory(base_directory_path);

            if (TreeStructureHelper.Find_File($"{base_directory_path}\\{excel_name}"))
                File.Delete($"{base_directory_path}\\{excel_name}");


            // ***** Main content of the Excel file ***** //

            ExcelEntity excel_entity = new ExcelEntity(user);

            excel_entity.Write_Excel_Header();
            excel_entity.Write_Excel_Content(days_in_month, user);
            excel_entity.Write_Excel_Footer();

            excel_entity.workbook.SaveAs($"{base_directory_path}/{excel_name}");
        }
    }
}