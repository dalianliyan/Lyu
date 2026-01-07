using System;
using System.Collections.Generic;
using System.Linq;

namespace MarvellousWorks.PracticalPattern.Common
{
    /// <summary>
    /// Utility for <see cref="IEnumerable<T>"/>
    /// </summary>
    public static class EnumerableHelper
    {
        /// <summary>
        /// Execute Action for each element in <see cref="IEnumerable<T>"/>
        /// </summary>
        public static void ForEach<T>(this IEnumerable<T> sequence, Action<T> action)
        {
            if (sequence == null)
                throw new ArgumentNullException("sequence");
            if (action == null)
                throw new ArgumentNullException("action");
            foreach (var item in sequence)
                action(item);
        }
    }
}
