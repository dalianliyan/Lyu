using System;
using System.Collections.Generic;
using System.Linq;

namespace MarvellousWorks.PracticalPattern.JavaAndDotNet
{
    public enum Season
    {
        Spring,
        Summer,
        Autumn,
        Winter
    }

    public class SeasonCalendar
    {
        static IDictionary<Season, SeasonMonthEntry> monthLookup =
            new Dictionary<Season, SeasonMonthEntry>();

        struct SeasonMonthEntry
        {
            public Season Season { get; set; }
            public int StartMonth { get; set; }
            public int EndMonth { get; set; }
            public string Note { get; set; }
        }

        static SeasonCalendar()
        {
            monthLookup.Add(Season.Spring, new SeasonMonthEntry(){
                Season = Season.Spring, StartMonth = 1, EndMonth = 3, Note = "生机盎然"});
            monthLookup.Add(Season.Summer, new SeasonMonthEntry(){
                Season = Season.Summer, StartMonth = 4, EndMonth = 6, Note = "烈日炎炎"});
            monthLookup.Add(Season.Autumn, new SeasonMonthEntry(){
                Season = Season.Autumn, StartMonth = 7, EndMonth = 9, Note = "硕果累累"});
            monthLookup.Add(Season.Winter, new SeasonMonthEntry(){
                Season = Season.Winter, StartMonth = 10, EndMonth = 12, Note = "冰天雪地"});
        }

        SeasonMonthEntry entry;
        public SeasonCalendar(Season season)
        {
            if (season == null) throw new ArgumentNullException("season");
            entry = monthLookup[season];
        }

        public Season Season { get { return entry.Season; } }
        public int StartMonth { get { return entry.StartMonth; } }
        public int EndMonth { get { return entry.EndMonth; } }
        public string Note { get { return entry.Note; } }

        /// <summary>
        /// get season of month
        /// </summary>
        public static Season GetSeason(int month)
        {
            if ((month > 12) || (month < 1))
                throw new NotSupportedException();
            int startMonth = ((month - 1)/3)*3 + 1;
            return monthLookup.Values.FirstOrDefault(x => x.StartMonth == startMonth).Season;
        }
    }
}