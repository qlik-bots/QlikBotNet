using System;
using System.Collections.Generic;
using System.Linq;


namespace QlikSenseEasy
{
    public static class EnumerableExtension
    {
        public static T PickRandom<T>(this IEnumerable<T> source)
        {
            return source.PickRandom(1).Single();
        }

        public static IEnumerable<T> PickRandom<T>(this IEnumerable<T> source, int count)
        {
            return source.ShufflePick().Take(count);
        }

        private static IEnumerable<T> ShufflePick<T>(this IEnumerable<T> source)
        {
            return source.OrderBy(x => Guid.NewGuid());
        }

        public static IEnumerable<T> Shuffle<T>(this IEnumerable<T> list)
        {
            var r = new Random();
            int size = list.Count();

            var shuffledList =
                list.
                    Select(x => new { Number = r.Next(), Item = x }).
                    OrderBy(x => x.Number).
                    Select(x => x.Item).
                    Take(size); // Assume first @size items is fine

            return shuffledList.ToList();
        }

        //public static IEnumerable<T> Shuffle<T>(this IEnumerable<T> source)
        //{
        //    return source.Shuffle(new Random());
        //}

        //public static IEnumerable<T> Shuffle<T>(this IEnumerable<T> source, Random rng)
        //{
        //    if (source == null) throw new ArgumentNullException("source");
        //    if (rng == null) throw new ArgumentNullException("rng");

        //    return source.ShuffleIterator(rng);
        //}

        //private static IEnumerable<T> ShuffleIterator<T>(
        //    this IEnumerable<T> source, Random rng)
        //{
        //    var buffer = source.ToList();
        //    for (int i = 0; i < buffer.Count; i++)
        //    {
        //        int j = rng.Next(i, buffer.Count);
        //        yield return buffer[j];

        //        buffer[j] = buffer[i];
        //    }
        //}
    }
}
