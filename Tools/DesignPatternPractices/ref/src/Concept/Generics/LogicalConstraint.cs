namespace MarvellousWorks.PracticalPattern.Concept.Generics
{
    /// <summary>
    /// 模拟 where T : A || B 
    /// </summary>
    /// <remarks>之前OrA和OrB是两个独立的类型，为了找到两者的共同点，额外定义了一个INewComer接口，从外围强制建立了一个新的“或”关系</remarks>
    interface INewComer { }
    class OrA : INewComer { }
    class OrB : INewComer { }
    class OrClient<T> where T : INewComer { }


    /// <summary>
    /// 模拟 where T : A && B
    /// </summary>
    interface ILayer11 { }
    interface ILayer12 { }
    interface ILayer2 : ILayer11, ILayer12 { }
    class AndA : ILayer11 { }
    class AndB : ILayer12 { }
    class AndClient<T> where T : ILayer2 { }
}
