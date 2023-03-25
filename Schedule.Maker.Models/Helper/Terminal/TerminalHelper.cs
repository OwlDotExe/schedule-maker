using Spectre.Console;

namespace Schedule.Maker.Models.Helper.Terminal
{
    /// <summary>
    ///  Helper class used to ask or show information to the user.
    /// </summary>
    public abstract class TerminalHelper
    {
        private const string ansi_end_style = "[/]";
        private const string ansi_accentuated = "[bold]";
        private const string ansi_attenuated = "[bold grey]";
        private const string ansi_highlight_color = "[orange1]";

        /**
         * <summary>
         *  Function used to write the main presentation of the program.  
         * </summary>
         */
        public static void Write_Program_Information()
        {
            AnsiConsole.MarkupLine($"{ansi_accentuated}AUTHOR:{ansi_end_style} {ansi_attenuated}Takoune{ansi_end_style}");
            AnsiConsole.MarkupLine($"{ansi_accentuated}DESCRIPTION:{ansi_end_style} {ansi_attenuated}Automation project to generate Excel & PDF files that contain working hour for a given month.{ansi_end_style}");
            AnsiConsole.MarkupLine($"{ansi_accentuated}VERSION:{ansi_end_style} {ansi_attenuated}1.0{ansi_end_style}");
            AnsiConsole.MarkupLine($"{ansi_accentuated}GITHUB PROFIL:{ansi_end_style} {ansi_attenuated}[link]https://github.com/takoune{ansi_end_style}{ansi_end_style}\n");
        }

        /**
         * <summary>
         *  Function used to write the main instruction to the user.
         * </summary>
         */
        public static void Write_User_Instructions()
        {
            AnsiConsole.MarkupLine($"{ansi_accentuated}For the great generation of the {ansi_highlight_color}Excel & PDF files{ansi_end_style} this program will ask you some information :[/]");
            AnsiConsole.MarkupLine($"    {ansi_attenuated}1~{ansi_end_style} {ansi_accentuated}Your {ansi_highlight_color}first name{ansi_end_style} and {ansi_highlight_color}last name{ansi_end_style}.{ansi_end_style}");
            AnsiConsole.MarkupLine($"    {ansi_attenuated}2~{ansi_end_style} {ansi_accentuated}The {ansi_highlight_color}month{ansi_end_style} used for the generation.{ansi_end_style}");
            AnsiConsole.MarkupLine($"    {ansi_attenuated}3~{ansi_end_style} {ansi_accentuated}Potential {ansi_highlight_color}days off{ansi_end_style} taken during the month.{ansi_end_style}\n");

            AnsiConsole.MarkupLine($"{ansi_attenuated}(Press ENTER if you want to continue){ansi_end_style}");

            Wait_For_Key(ConsoleKey.Enter);
        }

        /**
         * <summary>
         *  Function used to ask the user for his first name.
         * </summary>
         * <return>first_name - The first name entered by the user.</return>
         */
        public static string Ask_First_Name()
        {
            bool is_asking = true;
            string first_name = string.Empty;

            while (is_asking)
            {
                Write_Step($"{ansi_attenuated}1~{ansi_end_style} {ansi_accentuated}[underline]Asking for the user identity.{ansi_end_style}{ansi_end_style}\n");

                first_name = AnsiConsole.Ask<string>($"{ansi_accentuated}What is the {ansi_highlight_color}first name{ansi_end_style} to use for the generation:{ansi_end_style}");

                AnsiConsole.MarkupLine($"\nCan you please confirm that the spelling of your first name is right: {ansi_highlight_color}{first_name}{ansi_end_style}\n");

                if (Wait_For_User_Agreement(ConsoleKey.Enter, "Press any other key if you want type again your first name"))
                    is_asking = false;
                else
                {
                    is_asking = true;
                }
            }

            return first_name;
        }

        /**
         * <summary>
         *  Function used to ask the user for his last name.
         * </summary>
         * <return>last_name - The last name entered by the user.</return>
         */
        public static string Ask_Last_Name()
        {
            bool is_asking = true;
            string last_name = string.Empty;

            while (is_asking)
            {
                Write_Step($"{ansi_attenuated}1~{ansi_end_style} {ansi_accentuated}[underline]Asking for the user identity.{ansi_end_style}{ansi_end_style}\n");

                last_name = AnsiConsole.Ask<string>($"{ansi_accentuated}What is the {ansi_highlight_color}last name{ansi_end_style} to use for the generation:{ansi_end_style}");

                AnsiConsole.MarkupLine($"\nCan you please confirm that the spelling of your last name is right: {ansi_highlight_color}{last_name}{ansi_end_style}\n");

                if (Wait_For_User_Agreement(ConsoleKey.Enter, "Press any other key if you want type again your last name"))
                    is_asking = false;
                else
                {
                    is_asking = true;
                }
            }

            return last_name;
        }

        /**
         * <summary>
         *  Function used to ask the user what month should be used for the generation.
         * </summary>
         * <return>month - The month that has been selected by the user.</return>
         */
        public static string Ask_Month()
        {
            bool is_asking = true;
            string month = string.Empty;

            string[] month_names = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.MonthNames;

            while (is_asking)
            {
                Write_Step($"{ansi_attenuated}2~{ansi_end_style} {ansi_accentuated}[underline]Asking for the month to use for the generation.{ansi_end_style}{ansi_end_style}\n");

                month = AnsiConsole.Prompt(new SelectionPrompt<string>()
                                   .Title($"What {ansi_highlight_color}month{ansi_end_style} should be used for the generation ?")
                                   .PageSize(12)
                                   .AddChoices(month_names.Take(month_names.Length - 1))); ;

                AnsiConsole.MarkupLine($"{ansi_accentuated}You decided to generate {ansi_highlight_color}{month}{ansi_end_style} work hours.{ansi_end_style}\n");

                if (Wait_For_User_Agreement(ConsoleKey.Enter, "Press any other key if you want to pick another month"))
                    is_asking = false;
                else
                {
                    is_asking = true;
                }
            }

            return month;
        }

        /**
         * <summary>
         *  Function used to ask the user if some days off has been taken during the selected month.
         * </summary>
         * <param>days_in_month - The list of days in the previous selected month.</param>
         * <param>month - The previous selected month.</param>
         * <param>month_number - The int value of the previous selected month.</param>
         * <return>days_off - The potential list of days off that has been taken by the user.</return>
         */
        public static List<DateTime> Ask_Days_Off(List<DateTime> days_in_month, string month, int month_number)
        {
            bool is_asking = true;
            List<DateTime> days_off = new List<DateTime>();

            while (is_asking)
            {
                Write_Step($"{ansi_attenuated}3~{ansi_end_style} {ansi_accentuated}[underline]Asking for the potential days off taken during the month.{ansi_end_style}{ansi_end_style}\n");

                days_off = AnsiConsole.Prompt(new MultiSelectionPrompt<DateTime>()
                                      .Title($"Feel free to select your {ansi_highlight_color}days off{ansi_end_style} in the current list :")
                                      .NotRequired()
                                      .PageSize(15)
                                      .MoreChoicesText($"{ansi_attenuated}(Move up and down to reveal all the day of {month}){ansi_end_style}")
                                      .InstructionsText($"{ansi_attenuated}(Press [blue]<SPACE>[/] to choose a day as a day off and [green]<ENTER>[/] to validate your selection){ansi_end_style}")
                                      .AddChoices(days_in_month));

                Write_Step($"{ansi_attenuated}3~{ansi_end_style} {ansi_accentuated}[underline]Asking for the potential days off taken during the month.{ansi_end_style}{ansi_end_style}\n");

                Show_Calendar(days_off, month_number);

                AnsiConsole.Markup($"\n{ansi_accentuated}(days off = {ansi_highlight_color}*{ansi_end_style}){ansi_end_style}\n\n");

                if (Wait_For_User_Agreement(ConsoleKey.Enter, "Press any key if you want to select again your days off"))
                    is_asking = false;
                else
                {
                    is_asking = true;
                }
            }

            return days_off;
        }

        /**
         * <summary>
         *  Function used to display the calendar with the potential days off.
         * </summary>
         * <param>days_off - The potential list of days off selected by the user.</param>
         * <param>month_number - The month selected by the user.</param>
         */
        private static void Show_Calendar(List<DateTime> days_off, int month_number)
        {
            Calendar calendar = new Calendar(DateTime.Now.Year, month_number).HeaderStyle(Style.Parse("orange1 bold")).HighlightStyle(Style.Parse("orange1 bold"));

            days_off.ForEach(day => calendar.AddCalendarEvent(day.Year, day.Month, day.Day));

            AnsiConsole.Write(calendar);
        }

        /**
         * <summary>
         *  Function used to write the header of the step.
         * </summary>
         * <param>step_sentence - The header sentence to use.</param>
         */
        private static void Write_Step(string step_sentence)
        {
            Console.Clear();

            AnsiConsole.MarkupLine(step_sentence);
        }

        /**
         * <summary>
         *  Function used to wait for a special key to be pressed by the user.
         * </summary>
         * <param>key - The key pressed by the user.</param>
         */
        private static void Wait_For_Key(ConsoleKey key)
        {
            while (Console.ReadKey(true).Key != key) {}
        }

        /**
         * <summary>
         *  Function used to handle the cancelling or the validation of a step by the user.
         * </summary>
         * <param>agreement_key - The key that validates the current step.</param>
         * <return>Boolean value that indicates if the user wants to confirm or cancel the current step.</return>
         */
        private static bool Wait_For_User_Agreement(ConsoleKey agreement_key, string cancel_sentence)
        {
            AnsiConsole.MarkupLine($"{ansi_attenuated}(Press {agreement_key.ToString().ToUpper()} if you want to continue){ansi_end_style}");
            AnsiConsole.MarkupLine($"{ansi_attenuated}({cancel_sentence}){ansi_end_style}");

            ConsoleKeyInfo key = Console.ReadKey(true);

            if (key.Key == agreement_key)
                return true;
            return false;
        }
    }
}
