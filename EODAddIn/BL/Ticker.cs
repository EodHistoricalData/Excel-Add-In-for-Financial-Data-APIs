namespace EODAddIn.BL
{
    public class Ticker
    {
        public string Name { get; set; }
        public string Exchange { get; set; }

        public string FullName => $"{Name}.{Exchange}";
    }
}
