using System;

namespace EODAddIn.Program
{
    internal class Version
    {
        internal string Name;
        internal int Major;
        internal int Minor;
        internal int Build;
        internal int Revision;
        /// <summary>
        /// Версия программы текстом
        /// </summary>
        internal string Text { get { return $"{Major}.{Minor}.{Build}.{Revision}"; } }
        internal DateTime Date;
        internal string Description;
        internal string Link;
    }
}
