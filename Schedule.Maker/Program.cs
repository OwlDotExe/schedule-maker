using Schedule.Maker.Models.Entity;
using Schedule.Maker.Models.Helper.Terminal;
using Schedule.Maker.Models.Helper.TreeStructure;
using System.Globalization;

namespace Schedule.Maker
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // ***** User interactions & retrieval of user's information ***** //

            TerminalHelper.Write_Program_Information();
            TerminalHelper.Write_User_Instructions();

            string first_name = TerminalHelper.Ask_First_Name();
            string last_name = TerminalHelper.Ask_Last_Name();

            string month = TerminalHelper.Ask_Month();

            int month_number = DateTime.ParseExact($"{month}", "MMMM", new CultureInfo("en-US")).Month;

            List<DateTime> days_in_month = Enumerable.Range(1, DateTime.DaysInMonth(DateTime.Now.Year, month_number))
                                                     .Select(day => new DateTime(DateTime.Now.Year, month_number, day, 0, 0, 0))
                                                     .ToList();

            List<DateTime> days_off = TerminalHelper.Ask_Days_Off(days_in_month, month, month_number);

            User user = new User(first_name, last_name, days_off);


            // ***** Tree structure establishment ***** //

            string base_directory_path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\schedule-maker-files";

            if (!TreeStructureHelper.Find_Directory(base_directory_path))
                Directory.CreateDirectory(base_directory_path);

            string excel_name = $"{month_number}.{month}_{DateTime.Now.Year}.xlsx";
            string pdf_name = $"{month_number}.{month}_{DateTime.Now.Year}.pdf";

            if (TreeStructureHelper.Find_File($"{base_directory_path}\\{excel_name}"))
                File.Delete($"{base_directory_path}\\{excel_name}");

            if (TreeStructureHelper.Find_File($"{base_directory_path}\\{pdf_name}"))
                File.Delete($"{base_directory_path}\\{pdf_name}");

            Console.ReadKey();
        }
    }
}