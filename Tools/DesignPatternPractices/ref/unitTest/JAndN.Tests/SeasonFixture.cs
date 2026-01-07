using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MarvellousWorks.PracticalPattern.JavaAndDotNet.Tests
{
    [TestClass]
    public class SeasonFixture
    {
        [TestMethod]
        public void SeasonEnum()
        {
            var winter = new SeasonCalendar(Season.Winter);

            Assert.AreEqual<int>(10, winter.StartMonth);
            Assert.AreEqual<int>(12, winter.EndMonth);
            Assert.AreEqual<string>("冰天雪地", winter.Note);
            Assert.AreNotEqual<int>(11, new SeasonCalendar(Season.Autumn).StartMonth);

            Assert.AreEqual<Season>(Season.Winter, SeasonCalendar.GetSeason(12));
            Assert.AreEqual<Season>(Season.Winter, SeasonCalendar.GetSeason(11));
            Assert.AreEqual<Season>(Season.Autumn, SeasonCalendar.GetSeason(8));
        }
    }
}
