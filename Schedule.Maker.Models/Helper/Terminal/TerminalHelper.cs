using Schedule.Maker.Models.Entity.Data;
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
        private const string ansi_highlight_color_2 = "[purple_1]";

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
            AnsiConsole.MarkupLine($"    {ansi_attenuated}3~{ansi_end_style} {ansi_accentuated}{ansi_highlight_color}Days off{ansi_end_style} / {ansi_highlight_color}public days{ansi_end_style} / {ansi_highlight_color}remote days{ansi_end_style}.{ansi_end_style}");
            AnsiConsole.MarkupLine($"    {ansi_attenuated}4~{ansi_end_style} {ansi_accentuated}Final {ansi_highlight_color}summary{ansi_end_style} that contains all the {ansi_highlight_color}user information{ansi_end_style}.{ansi_end_style}\n");


            AnsiConsole.MarkupLine($"{ansi_attenuated}(Press ENTER if you want to continue){ansi_end_style}");

            Wait_For_Key(ConsoleKey.Enter);
        }

        /**
         * <summary>
         *  Function used to write the final summary of all the user's information.
         * </summary>
         * <param name="user">The user object that contains user's information.</param>
         * <param name="excel_name">The excel file name with embedded style.</param>
         * <return>Boolean value which indicates if the user has confirmed or not the summary of his own information.</return>
         */
        public static bool Write_User_Summary(User user, string excel_name)
        {
            Write_Step($"{ansi_attenuated}4~{ansi_end_style} {ansi_accentuated}[underline]Summary of user information.{ansi_end_style}{ansi_end_style}\n");

            Write_User_Info(user);

            AnsiConsole.MarkupLine($"{ansi_attenuated}~~>{ansi_end_style} {ansi_accentuated}Preview of the {ansi_highlight_color}tree structure{ansi_end_style} that will be generated:{ansi_end_style}\n");

            Write_Tree_Structure(excel_name);

            return !Wait_For_User_Agreement(ConsoleKey.Enter, "Press any other key if you want enter new information");
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
         * <param name="days_in_month">The list of days in the previous selected month.</param>
         * <param name="month">The previous selected month.</param>
         * <param name="month_number">The int value of the previous selected month.</param>
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

                Show_Non_Working_Calendar(days_off, new List<DateTime> { }, month_number);

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
         *  Function used to ask the user if there is public days in the selected month.
         * </summary>
         * <param name="rest_of_days">The list of days that can be selected by the user (without days off)</param>
         * <param name="month">The previous selected month.</param>
         * <param name="month_number">The int value of the previous selected month.</param>
         * <return>public_days - The potential list of public days of the selected month.</return>
         */
        public static List<DateTime> Ask_Public_Days(List<DateTime> rest_of_days, string month, int month_number)
        {
            bool is_asking = true;
            List<DateTime> public_days = new List<DateTime>();

            while (is_asking)
            {
                Write_Step($"{ansi_attenuated}3~{ansi_end_style} {ansi_accentuated}[underline]Asking for the potential public days of the month.{ansi_end_style}{ansi_end_style}\n");

                public_days = AnsiConsole.Prompt(new MultiSelectionPrompt<DateTime>()
                                         .Title($"Feel free to select all the {ansi_highlight_color}public days{ansi_end_style} in the current list :")
                                         .NotRequired()
                                         .PageSize(15)
                                         .MoreChoicesText($"{ansi_attenuated}(Move up and down to reveal all the day of {month}){ansi_end_style}")
                                         .InstructionsText($"{ansi_attenuated}(Press [blue]<SPACE>[/] to choose a day as a public day and [green]<ENTER>[/] to validate your selection){ansi_end_style}")
                                         .AddChoices(rest_of_days));

                Write_Step($"{ansi_attenuated}3~{ansi_end_style} {ansi_accentuated}[underline]Asking for the potential public days of the month.{ansi_end_style}{ansi_end_style}\n");

                Show_Non_Working_Calendar(new List<DateTime> { }, public_days, month_number);

                AnsiConsole.Markup($"\n{ansi_accentuated}(public days = {ansi_highlight_color}*{ansi_end_style}){ansi_end_style}\n\n");

                if (Wait_For_User_Agreement(ConsoleKey.Enter, "Press any key if you want to select again public days"))
                    is_asking = false;
                else
                {
                    is_asking = true;
                }
            }

            return public_days;
        }

        /**
         * <summary>
         *  Function used to ask the user for the remote days taken during the month.
         * </summary>
         * <param name="rest_of_days">The list of days that can be selected by the user (without days off & public days)</param>
         * <param name="month">The previous selected month.</param>
         * <param name="month_number">The int value of the previous selected month.</param>
         * <return>remote_days - The list of remote days taken by the user.</return>
         */
        public static List<DateTime> Ask_Remote_Days(List<DateTime> rest_of_days, string month, int month_number)
        {
            bool is_asking = true;
            List<DateTime> remote_days = new List<DateTime>();

            while (is_asking)
            {
                Write_Step($"{ansi_attenuated}3~{ansi_end_style} {ansi_accentuated}[underline]Asking for the remote days taken during the month.{ansi_end_style}{ansi_end_style}\n");

                remote_days = AnsiConsole.Prompt(new MultiSelectionPrompt<DateTime>()
                                         .Title($"Feel free to select all the {ansi_highlight_color}remote days{ansi_end_style} in the current list :")
                                         .PageSize(15)
                                         .Required()
                                         .MoreChoicesText($"{ansi_attenuated}(Move up and down to reveal all the day of {month}){ansi_end_style}")
                                         .InstructionsText($"{ansi_attenuated}(Press [blue]<SPACE>[/] to choose a day as a remote day and [green]<ENTER>[/] to validate your selection){ansi_end_style}")
                                         .AddChoices(rest_of_days));

                Write_Step($"{ansi_attenuated}3~{ansi_end_style} {ansi_accentuated}[underline]Asking for the remote days taken during the month.{ansi_end_style}{ansi_end_style}\n");

                Show_Working_Calendar(remote_days, month_number);

                AnsiConsole.Markup($"\n{ansi_accentuated}(remote days = {ansi_highlight_color_2}*{ansi_end_style}){ansi_end_style}\n\n");

                if (Wait_For_User_Agreement(ConsoleKey.Enter, "Press any key if you want to select again public days"))
                    is_asking = false;
                else
                {
                    is_asking = true;
                }
            }

            return remote_days;
        }

        /**
         * <summary>
         *  Function used to display the calendar with days off and public days.
         * </summary>
         * <param name="days_off">The potential list of days off selected by the user.</param>
         * <param name="public_days">The potential list of public days selected by the user.</param>
         * <param name="month_number">The month selected by the user.</param>
         */
        private static void Show_Non_Working_Calendar(List<DateTime> days_off, List<DateTime> public_days, int month_number)
        {
            Calendar calendar = new Calendar(DateTime.Now.Year, month_number).HeaderStyle(Style.Parse("orange1 bold")).HighlightStyle(Style.Parse($"orange1 bold"));

            days_off.ForEach(day => calendar.AddCalendarEvent(day.Year, day.Month, day.Day));
            public_days.ForEach(day => calendar.AddCalendarEvent(day.Year, day.Month, day.Day));

            AnsiConsole.Write(calendar);
        }

        /**
         * <summary>
         *  Function used to display the calendar that contains all the remote days.
         * </summary>
         * <param name="remote_days">The potential list of remote days selected by the user.</param>
         * <param name="month_number">The month selected by the user.</param>
         */
        private static void Show_Working_Calendar(List<DateTime> remote_days, int month_number)
        {
            Calendar calendar = new Calendar(DateTime.Now.Year, month_number).HeaderStyle(Style.Parse("purple_1 bold")).HighlightStyle(Style.Parse($"purple_1 bold"));

            remote_days.ForEach(day => calendar.AddCalendarEvent(day.Year, day.Month, day.Day));

            AnsiConsole.Write(calendar);
        }

        /**
         * <summary>
         *  Function used to write the header of the step.
         * </summary>
         * <param name="step_sentence">The header sentence to use.</param>
         */
        private static void Write_Step(string step_sentence)
        {
            Console.Clear();

            AnsiConsole.MarkupLine(step_sentence);
        }

        /**
         * <summary>
         *  Function used to write the tree structure that will be generated.
         * </summary>
         * <param name="excel_name">The final name of the excel file that will be generated.</param>
         */
        private static void Write_Tree_Structure(string excel_name)
        {
            Tree root = new Tree($"{ansi_accentuated}Desktop{ansi_end_style}");

            TreeNode main_directory_node = root.AddNode($"{ansi_accentuated}[orange1]shedule-maker-files{ansi_end_style}{ansi_end_style} [bold dim](directory)[/]");

            main_directory_node.AddNode($"{excel_name} [bold dim](file)[/]");

            AnsiConsole.Write(root);
            AnsiConsole.Markup("\n");
        }

        /**
         * <summary>
         *  Function used to write user information.
         * </summary>
         * <param name="user">The user object that contains all the data.</param>
         */
        private static void Write_User_Info(User user)
        {
            AnsiConsole.MarkupLine($"{ansi_attenuated}~~>{ansi_end_style} {ansi_accentuated}Full name that will be used for the generation: {ansi_highlight_color}{user.first_name} {user.last_name}{ansi_end_style}{ansi_end_style}");
            AnsiConsole.MarkupLine($"{ansi_attenuated}~~>{ansi_end_style} {ansi_accentuated}Month that will be used for the generation: {ansi_highlight_color}{user.month}{ansi_end_style}{ansi_end_style}");
            AnsiConsole.MarkupLine($"{ansi_attenuated}~~>{ansi_end_style} {ansi_accentuated}Potential {ansi_highlight_color}days off{ansi_end_style} & {ansi_highlight_color}public days{ansi_end_style}:{ansi_end_style}\n");

            Show_Non_Working_Calendar(user.days_off, user.public_days, user.month_number);

            AnsiConsole.Markup($"\n{ansi_accentuated}(days off & public days = {ansi_highlight_color}*{ansi_end_style}){ansi_end_style}\n");

            AnsiConsole.MarkupLine($"\n{ansi_attenuated}~~>{ansi_end_style} {ansi_accentuated}Summary of the {ansi_highlight_color_2}remote days{ansi_end_style}:{ansi_end_style}\n");

            Show_Working_Calendar(user.remote_days, user.month_number);

            AnsiConsole.Markup($"\n{ansi_accentuated}(remote days = {ansi_highlight_color_2}*{ansi_end_style}){ansi_end_style}\n");

            AnsiConsole.Markup("\n");
        }

        /**
         * <summary>
         *  Function used to wait for a special key to be pressed by the user.
         * </summary>
         * <param name="key">The key pressed by the user.</param>
         */
        private static void Wait_For_Key(ConsoleKey key)
        {
            while (Console.ReadKey(true).Key != key) {}
        }

        /**
         * <summary>
         *  Function used to handle the cancelling or the validation of a step by the user.
         * </summary>
         * <param name="agreement_key">The key that validates the current step.</param>
         * <param name="cancel_sentence">The cancel sentence that is displayed to the user.</param>
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
