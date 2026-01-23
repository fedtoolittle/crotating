using System.Globalization;

namespace Crotating.Services
{
    internal static class CellNormalization
    {
        /// <summary>
        /// Safely converts an Excel cell value to double.
        /// Returns 0 for null, empty, or non-numeric values.
        /// </summary>
        public static double GetDoubleOrZero(object cellValue)
        {
            if (cellValue == null)
                return 0d;

            // Native Excel numeric
            if (cellValue is double d)
                return d;

            // EPPlus may surface decimals or integers
            if (cellValue is decimal dec)
                return (double)dec;

            if (cellValue is int i)
                return i;

            if (cellValue is long l)
                return l;

            // String fallback (formatted numbers, formulas, etc.)
            if (double.TryParse(
                cellValue.ToString()?.Trim(),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out var parsed))
            {
                return parsed;
            }

            return 0d;
        }
    }
}
