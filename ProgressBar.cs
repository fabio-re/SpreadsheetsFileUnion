namespace SpreadsheetsFileUnion
{
    public class ConsoleProgressBar : IProgressBar
    {
        private const ConsoleColor ForeColor = ConsoleColor.White;
        private const int TextMarginLeft = 2;
        private readonly int _total;
        private readonly int _widthOfBar;
        private bool _intited;
        private int _lastPosition;

        public ConsoleProgressBar(int total)
        {
            _total = total;
            _widthOfBar = (total + 1) * 2 + total + 2 + 3;
        }
        public void Init()
        {
            _lastPosition = 0;
            Console.CursorVisible = false;
            Console.CursorLeft = 0;
            Console.Write("[");
            Console.CursorLeft = _widthOfBar;
            Console.Write("]");
            Console.CursorLeft = 1;
            for (int position = 1; position < _widthOfBar; position++)
            {
                Console.CursorLeft = position;
                Console.Write(" ");
            }
        }

        public void ShowProgress(int currentCount)
        {
            if (!_intited)
            {
                Init();
                _intited = true;
            }
            DrawTextProgressBar(currentCount);
        }

        public void DrawTextProgressBar(int currentCount)
        {
            int position = currentCount * _widthOfBar / _total;
            if (position != _lastPosition)
            {
                _lastPosition = position;
                Console.BackgroundColor = ConsoleColor.White;
                Console.CursorLeft = ((position >= _widthOfBar) ? (_widthOfBar - 1) : position);
                Console.Write(" ");
            }
            Console.CursorLeft = _widthOfBar + 2;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Write(currentCount + " di " + _total + "    ");
        }
    }
    public interface IProgressBar
    {
        void ShowProgress(int currentCount);
    }
}
