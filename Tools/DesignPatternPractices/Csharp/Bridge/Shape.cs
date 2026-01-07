using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bridge
{
    /// <summary>
    /// As Abstraction.
    /// </summary>
    public abstract class Shape
    {

        public Shape(DrawingProgramming dp)
        {
            this.Dp = dp;
        }

        public DrawingProgramming Dp { get; set; }

        public abstract void Draw();
    }

    /// <summary>
    /// As Refined Abstraction.
    /// </summary>
    public class Retangle : Shape
    {

        public Retangle(DrawingProgramming dp, int width, int height)
            : base(dp)
        {
            this.Width = width;
            this.Height = height;
        }
        public int Width { get; set; }

        public int Height { get; set; }

        public override void Draw()
        {
            this.DrawRetangle(Width, Height);
        }

        public void DrawRetangle(int width, int height)
        {
            this.Dp.DrawRetangle(width, height);
        }
    }

    /// <summary>
    /// As Refined Abstraction.
    /// </summary>
    public class Triangle : Shape
    {
    
        public Triangle(DrawingProgramming dp, int rows)
            : base(dp)
        {
            this.Rows = rows;
        }

        public override void Draw()
        {
            this.DrawTriangle(this.Rows);
        }

        public void DrawTriangle(int rows)
        {
            this.Dp.DrawTriangle(rows);
        }

        public int Rows { get; set; }
    }

}
