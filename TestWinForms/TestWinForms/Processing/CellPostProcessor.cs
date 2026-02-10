using OfficeOpenXml;

namespace Crotating.Services
{
    internal static class CellPostProcessor
    {
        public static void FillBlankCellsWithZero(
            ExcelWorksheet worksheet,
            int startRow,
            int startColumn,
            int endRow,
            int endColumn)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int column = startColumn; column <= endColumn; column++)
                {
                    var cell = worksheet.Cells[row, column];
                    if (IsBlank(cell.Value))
                    {
                        cell.Value = 0m;
                    }
                }
            }
        }

        public static void FillBlankCellsWithZero(ExcelWorksheet worksheet, ExcelRangeBase range)
        {
            FillBlankCellsWithZero(
                worksheet,
                range.Start.Row,
                range.Start.Column,
                range.End.Row,
                range.End.Column);
        }

        private static bool IsBlank(object value)
        {
            if (value == null)
                return true;

            if (value is string text)
                return string.IsNullOrWhiteSpace(text);

            return false;
        }
    }
}
