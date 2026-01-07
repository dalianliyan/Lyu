namespace MarvellousWorks.PracticalPattern.State.Raw
{
    public enum Color { Red, Green, Yellow }

    public class StopLight
    {
        Color current = Color.Yellow;    // default,用来保存当前状态

        /// <summary>
        /// 用一组排比的if else完成转换逻辑
        /// </summary>
        /// <returns></returns>
        public Color Next
        {
            get
            {
                if (current == Color.Green)
                    current = Color.Yellow;
                else if (current == Color.Yellow)
                    current = Color.Red;
                else if (current == Color.Red)
                    current = Color.Green;
                return current;
            }
        }
    }
}
