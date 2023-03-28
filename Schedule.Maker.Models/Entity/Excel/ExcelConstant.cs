namespace Schedule.Maker.Models.Entity.Excel
{
    /// <summary>
    ///  Class that contains all the excel readonlyant.
    /// </summary>
    internal class ExcelConstant
    {
        internal readonly string primary = "#FFFFFF";
        internal readonly string secondary = "#D9D9D9";

        internal readonly string hour_reference_title = "HEURES REFERENCE";
        internal readonly string morning_start = "09:00";
        internal readonly string morning_end = "12:00";
        internal readonly string afternoon_start = "13:00";
        internal readonly string afternoon_end = "17:00";
        internal readonly string hour_total = "07:00";
        internal readonly string hour_none_total = "00:00";
        internal readonly string hour_format = "hh:mm";

        internal readonly string morning_section = "MATIN";
        internal readonly string afternoon_section = "APRES-MIDI";

        internal readonly string day_column = "JOUR";
        internal readonly string remote_column = "TELETRAVAIL";
        internal readonly string start_column = "DEBUT";
        internal readonly string end_column = "FIN";
        internal readonly string comment_column = "COMMENTAIRE";
        internal readonly string hour_summary_column = "TOTAL (H)";

        internal readonly string valid_color = "#079992";
        internal readonly string invalid_color = "#b71540";

        internal readonly string morning_start_user = "07:30";
        internal readonly string morning_end_user = "12:30";
        internal readonly string afternoon_start_user = "14:00";
        internal readonly string afternoon_end_user = "16:00";

        internal readonly string signature_label = "SIGNATURE";
    }
}
