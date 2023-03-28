namespace Schedule.Maker.Models.Entity.Excel
{
    /// <summary>
    ///  Mother class of ExcelMergeCell and ExcelBorderCell that contains the range property.
    /// </summary>
    internal class ExcelCell
    {
        internal string range { get; set; }
        internal string color { get; set; }
        internal string borders { get; set; }

        public ExcelCell(string range, string color = "", string borders = "")
        {
            this.range = range;

            this.color = color;
            this.borders = borders;
        }
    }
}
