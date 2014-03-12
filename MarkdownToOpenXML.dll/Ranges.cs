
namespace MarkdownToOpenXML
{
    using System;
    using System.Collections.Generic;
    
    public class Ranges<T> where T : IComparable<T>
    {
        private List<Range<T>> rangelist = new List<Range<T>>();

        public void add(Range<T> range)
        {
            rangelist.Add(range);
        }

        public int Count()
        {
            return rangelist.Count;
        }

        public Boolean ContainsValue(T value)
        {
            foreach (Range<T> range in rangelist)
            {
                if (range.ContainsValue(value)) return true;
            }

            return false;
        }
    }
}
