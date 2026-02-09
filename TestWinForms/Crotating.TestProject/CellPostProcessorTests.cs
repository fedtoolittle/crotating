using Crotating.Services;
using NUnit.Framework;
using OfficeOpenXml;
using System;

namespace Crotating.Tests.Services
{
    [TestFixture]
    public class CellPostProcessorTests
    {
        [Test]
        public void FillBlankCellsWithZero_ReplacesBlankCellsInRandomized4x4Range()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Crotating.Tests");

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");

            var random = new Random(42);
            var expected = new object[4, 4];

            for (int row = 1; row <= 4; row++)
            {
                for (int col = 1; col <= 4; col++)
                {
                    var valueRoll = random.NextDouble();
                    object value = valueRoll switch
                    {
                        < 0.33 => null,
                        < 0.66 => "  ",
                        _ => Math.Round((decimal)random.NextDouble() * 10m, 2)
                    };

                    worksheet.Cells[row, col].Value = value;
                    expected[row - 1, col - 1] = value;
                }
            }

            CellPostProcessor.FillBlankCellsWithZero(worksheet, 1, 1, 4, 4);

            for (int row = 1; row <= 4; row++)
            {
                for (int col = 1; col <= 4; col++)
                {
                    var original = expected[row - 1, col - 1];
                    var current = worksheet.Cells[row, col].Value;

                    if (original == null || original is string text && string.IsNullOrWhiteSpace(text))
                    {
                        Assert.That(current, Is.EqualTo(0m), $"Expected cell ({row},{col}) to be 0m.");
                    }
                    else
                    {
                        Assert.That(current, Is.EqualTo(original));
                    }
                }
            }
        }
    }
}
