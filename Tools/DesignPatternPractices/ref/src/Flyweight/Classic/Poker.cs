namespace MarvellousWorks.PracticalPattern.Flyweight.Classic
{
    /// <summary>
    /// 包括两个享元类型成员的享元类型
    /// </summary>
    public class Poker
    {
        /// <summary>
        /// 花色
        /// </summary>
        public Suit Suit { get; set; }

        /// <summary>
        /// 点数
        /// </summary>
        public Rank Rank { get; set; }
    }
}
