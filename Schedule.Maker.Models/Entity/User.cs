namespace Schedule.Maker.Models.Entity
{
    /// <summary>
    ///  Class that contains information about the user.
    /// </summary>
    public class User
    {
        public string first_name { get; set; }
        public string last_name { get; set; }
        public List<DateTime> days_off { get; set; }

        public User(string first_name, string last_name, List<DateTime> days_off)
        {
            this.first_name = first_name;
            this.last_name = last_name;
            this.days_off = days_off;
        }
    }
}
