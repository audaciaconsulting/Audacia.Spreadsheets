using System.Collections.Generic;

namespace Audacia.Spreadsheets
{
    public class SectionEnumerable<T> : IEnumerable<T>
    {
        private readonly IEnumerator<T> _enumerator;

        public SectionEnumerable(IEnumerator<T> enumerator, int sectionSize)
        {
            _enumerator = enumerator;
            Left = sectionSize;
        }

        public IEnumerator<T> GetEnumerator()
        {
            while (Left > 0)
            {
                Left--;
                yield return _enumerator.Current;
                if (Left > 0)
                    if (!_enumerator.MoveNext())
                        break;
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int Left { get; private set; }
    }
}