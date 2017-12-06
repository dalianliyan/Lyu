using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace bubble
{
    class Program
    {
        static void Main(string[] args)
        {
            int[] arr = new int[] { 3, 2, 1, 4, 5, 7, 6 };
            string str = string.Empty;
            foreach (var member in arr)
            {
                str = str + " " + member.ToString();
            }
            Console.WriteLine(str);
            int[] val = bubble(arr);
            str = string.Empty;
            foreach (var member in val)
            {
                str = str + " " + member.ToString();
            }
            Console.WriteLine(str);
            Console.ReadLine();
        }

        static public int[] bubble(int[] array)
        {
            int num = array.Count();
            for (int i = num; i > 0; i--)
            {
                for (int j = 0; j < i - 1; j++)
                {
                    if (array[j + 1] < array[j])
                    {
                        int temp = array[j + 1];
                        array[j + 1] = array[j];
                        array[j] = temp;
                    }
                }
            }
            return array;
        }
    }
}
