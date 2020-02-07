using System;
using System.Collections.Generic;

namespace Audacia.Spreadsheets.Extensions
{
    public static class Sections
    {
        public static IEnumerable<IEnumerable<TEntity>> Section<TEntity>(this IEnumerable<TEntity> collection, int sectionSize)
        {
            if (collection == null)
            {
                throw new ArgumentNullException(nameof(collection));
            }

            if (sectionSize < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(sectionSize));
            }

            using (var enumerator = collection.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    var enumerable = new SectionEnumerable<TEntity>(enumerator, sectionSize);
                    
                    yield return enumerable;

                    for (var index = 0; index < enumerable.Left; index++)
                    {
                        if (!enumerator.MoveNext()) 
                        { 
                            yield break;
                        }
                    }
                }
            }
        }
    }
}