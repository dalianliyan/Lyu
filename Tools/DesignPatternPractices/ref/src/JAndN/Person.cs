using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MarvellousWorks.PracticalPattern.JavaAndDotNet
{
    class Person{}
    interface IPerson{}

    interface IPersonWithAbility<TAbility>{}
    interface IPersonWithAbilityAndBackground<TAbility, TBackground>{}

    interface IAbility<TMind, TPhysical>{}
    interface IPersonWithAbilityDetailAndBackground
        <Tmind, TPhysical, TBackground> {}
}
