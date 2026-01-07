namespace MarvellousWorks.PracticalPattern.Flyweight
{
    public enum SuitOptions
    {
        Spade, 
        Heart,
        Diamond,
        Club
    }

    /// <summary>
    /// 花色
    /// 享元类型
    /// </summary>
    public abstract class Suit{}
    class Spade : Suit { }
    class Heart : Suit { }
    class Diamond : Suit { }
    class Club : Suit { }

    /// <summary>
    /// 点数
    /// 享元类型
    /// </summary>
    public abstract class Rank { }
    class Ace : Rank { }
    class Deuce : Rank { }
    class Three : Rank { }
    class Four : Rank { }
    class Five : Rank { }
    class Six : Rank { }
    class Seven : Rank { }
    class Eight : Rank { }
    class Nine : Rank { }
    class Ten : Rank { }
    class Jack : Rank { }
    class Queue : Rank { }
    class King : Rank { }
}
