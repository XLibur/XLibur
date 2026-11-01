using System.Diagnostics;

namespace XLibur.Extensions;

internal static class IntegerExtensions
{
    public static bool Between(this int val, int from, int to)
    {
        return val >= from && val <= to;
    }

    extension(uint value)
    {
        /// <summary>
        /// Get index of the highest set bit &lt;= to <paramref name="maximalIndex"/> or -1 if no such bit.
        /// </summary>
        internal int GetHighestSetBitBelow(int maximalIndex)
        {
            Debug.Assert(maximalIndex is >= 0 and < 32);
            const uint highestBit = 0x80000000;
            value <<= 31 - maximalIndex;
            while (value != 0)
            {
                if ((value & highestBit) != 0)
                    return maximalIndex;
                value <<= 1;
                maximalIndex--;
            }

            return -1;
        }

        /// <summary>
        /// Get index of lowest set bit &gt;= to <paramref name="minimalIndex"/> or -1 if no such bit.
        /// </summary>
        internal int GetLowestSetBitAbove(int minimalIndex)
        {
            value >>= minimalIndex;
            while (value != 0)
            {
                if ((value & 1) == 1)
                    return minimalIndex;
                value >>= 1;
                minimalIndex++;
            }

            return -1;
        }

        /// <summary>
        /// Get the highest set bit index or -1 if no bit is set.
        /// </summary>
        internal int GetHighestSetBit()
        {
            var highestSetBitIndex = -1;
            while (value != 0)
            {
                value >>= 1;
                highestSetBitIndex++;
            }

            return highestSetBitIndex;
        }
    }
}
