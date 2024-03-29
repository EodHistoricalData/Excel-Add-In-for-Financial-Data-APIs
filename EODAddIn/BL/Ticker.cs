namespace EODAddIn.BL
{
    public class Ticker
    {
        public Ticker() { }
        public Ticker(string name, string exch)
        {
            Name = name;
            Exchange = exch;
        }

        public string Name { get; set; }
        public string Exchange { get; set; }

        public string FullName => $"{Name}.{Exchange}";
    }
}
