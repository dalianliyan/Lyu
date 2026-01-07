using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bridge
{

    /// <summary>
    /// As Implementor.
    /// </summary>
    public abstract class DrawingProgramming
    {
        /// <summary>
        /// Abstract draw method.
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public abstract void DrawRetangle(int width, int height);

        public abstract void DrawTriangle(int Rows);
    }


    /// <summary>
    /// As concrete Implementor
    /// </summary>
    public class OperationalDP1 : DrawingProgramming
    {
        public override void DrawRetangle(int width, int height)
        {
            DP1.DrawRetangle(width, height);
        }

        public override void DrawTriangle(int Rows)
        {
            DP1.DrawTriangle(Rows);
        }
    }

    /// <summary>
    /// As concrete Implementor
    /// </summary>
    public class OperationalDP2 : DrawingProgramming
    {
        public override void DrawRetangle(int width, int height)
        {
            DP2.DrawRetangle(width, height);
        }

        public override void DrawTriangle(int Rows)
        {
            DP2.DrawTriangle(Rows);
        }
    }


    /// <summary>
    /// existed drawing programming.
    /// </summary>
    public class DP1
    {
        /// <summary>
        /// Concrete draw method.
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public static void DrawRetangle(int width, int height)
        {
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    Console.Write("■");
                }
                Console.WriteLine();
            }
        }

        public static void DrawTriangle(int Rows)
        {
            for (int i = 0; i < Rows; i++)
            {
                for (int j = 0; j < i + 1; j++)
                {
                    Console.Write("■");
                }
                Console.WriteLine();
            }
        }
    }

    /// <summary>
    /// existed drawing programming.
    /// </summary>
    public class DP2
    {
        /// <summary>
        /// Concrete draw method.
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public static void DrawRetangle(int width, int height)
        {
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    Console.Write("★");
                }
                Console.WriteLine();
            }
        }

        public static void DrawTriangle(int Rows)
        {
            for (int i = 0; i < Rows; i++)
            {
                for (int j = 0; j < i + 1; j++)
                {
                    Console.Write("★");
                }
                Console.WriteLine();
            }
        }
    }

}
