using System;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Flyweight.Classic
{
    /// <summary>
    /// 享元工厂类型
    /// </summary>
    public class PokerFactory
    {
        const int MaxPoint = 13;
        const int MinPoint = 1;

        #region 可共享享元目标类型实例的缓存

        static IDictionary<SuitOptions, Suit> suits = new Dictionary<SuitOptions, Suit>()
                {
                    {SuitOptions.Spade, new Spade()},
                    {SuitOptions.Club, new Club()},
                    {SuitOptions.Diamond, new Diamond()},
                    {SuitOptions.Heart, new Heart()}
                };
        static IDictionary<int, Rank> points = new Dictionary<int, Rank>()
                {
                    {1, new Ace()},
                    {2, new Deuce()},
                    {3, new Three()},
                    {4, new Four()},
                    {5, new Five()},
                    {6, new Six()},
                    {7, new Seven()},
                    {8, new Eight()},
                    {9, new Nine()},
                    {10, new Ten()},
                    {11, new Jack()},
                    {12, new Queue()},
                    {13, new King()}
                };
        #endregion

        /// <summary>
        /// 享元类型实例缓存
        /// </summary>
        static readonly IDictionary<KeyValuePair<int, SuitOptions>, Poker> pokers =
            new Dictionary<KeyValuePair<int, SuitOptions>, Poker>();

        /// <summary>
        /// 根据享元目标类型特征生成享元类型实例
        /// </summary>
        /// <param name="suit">花色</param>
        /// <param name="rank">点数</param>
        /// <returns>扑克牌实例</returns>
        public Poker Create(int rank, SuitOptions suit)
        {
            if ((rank > MaxPoint) || (rank < MinPoint))
                throw new ArgumentOutOfRangeException("rank");
            
            Poker result;
            var key = new KeyValuePair<int, SuitOptions>(rank, suit);
            if (!pokers.TryGetValue(key, out result))
            {
                result = new Poker()
                {
                    Suit = suits[suit],
                    Rank = points[rank]
                };
                pokers.Add(key, result);
            }
            return result;
        }
    }
}
