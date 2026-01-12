using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Crotating.Models;
using Crotating.Services;

namespace Crotating.Tests.Services
{
    [TestFixture]
    public class ExcelReaderTests
    {
        private ExcelReader _excelReader;
        private string _testFilesDirectory;

        [SetUp]
        public void SetUp()
        {
            _excelReader = new ExcelReader();
            _testFilesDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestFiles");
            
            if (!Directory.Exists(_testFilesDirectory))
                Directory.CreateDirectory(_testFilesDirectory);
        }

        [Test]
        public void ReadWorkEntries_WithSingleHeaderRow_SkipsHeaderAndParsesData()
        {
            // Arrange
            var filePath = CreateTestFile("SingleHeader.xlsx", new[]
            {
                new[] { "Team member", "Date", "Duration", "Hours" },
                new[] { "John", "01/15/2024", "08:30:00", "8.5" },
                new[] { "John", "01/16/2024", "09:00:00", "9" }
            });

            // Act
            var results = _excelReader.ReadWorkEntries(filePath);

            // Assert
            Assert.That(results.Count, Is.EqualTo(2));
            Assert.That(results[0].Name, Is.EqualTo("John"));
            Assert.That(results[0].StartTime.Date, Is.EqualTo(new DateTime(2024, 1, 15)));
            Assert.That(results[0].HoursWorked, Is.EqualTo(8.5m));
        }

        [Test]
        public void ReadWorkEntries_WithMultipleHeaderRows_SkipsAllHeaders()
        {
            // Arrange
            var filePath = CreateTestFile("MultipleHeaders.xlsx", new[]
            {
                new[] { "Team member", "Date", "Duration", "Hours" },
                new[] { "Name", "Start Date", "Work Time", "Total Hours" },
                new[] { "Alice", "01/20/2024", "07:30:00", "7.5" },
                new[] { "Alice", "01/21/2024", "08:00:00", "8" }
            });

            // Act
            var results = _excelReader.ReadWorkEntries(filePath);

            // Assert
            Assert.That(results.Count, Is.EqualTo(2));
            Assert.That(results[0].Name, Is.EqualTo("Alice"));
            Assert.That(results[0].StartTime.Date, Is.EqualTo(new DateTime(2024, 1, 20)));
        }

        [Test]
        public void ReadWorkEntries_WithEmptyRowsBetweenData_SkipsEmptyRows()
        {
            // Arrange
            var filePath = CreateTestFile("WithEmptyRows.xlsx", new[]
            {
                new[] { "Team member", "Date", "Duration", "Hours" },
                new[] { "Bob", "01/10/2024", "06:00:00", "6" },
                new[] { "", "", "", "" },
                new[] { "Bob", "01/11/2024", "07:00:00", "7" }
            });

            // Act
            var results = _excelReader.ReadWorkEntries(filePath);

            // Assert
            Assert.That(results.Count, Is.EqualTo(2));
            Assert.That(results[0].Name, Is.EqualTo("Bob"));
            Assert.That(results[1].Name, Is.EqualTo("Bob"));
        }

        [Test]
        public void ReadWorkEntries_WithPersonNameCarryover_MaintainsCurrentName()
        {
            // Arrange
            var filePath = CreateTestFile("NameCarryover.xlsx", new[]
            {
                new[] { "Team member", "Date", "Duration", "Hours" },
                new[] { "Charlie", "01/05/2024", "08:00:00", "8" },
                new[] { "", "01/06/2024", "08:30:00", "8.5" },
                new[] { "Diana", "01/07/2024", "09:00:00", "9" },
                new[] { "", "01/08/2024", "07:30:00", "7.5" }
            });

            // Act
            var results = _excelReader.ReadWorkEntries(filePath);

            // Assert
            Assert.That(results.Count, Is.EqualTo(4));
            Assert.That(results[0].Name, Is.EqualTo("Charlie"));
            Assert.That(results[1].Name, Is.EqualTo("Charlie"));
            Assert.That(results[2].Name, Is.EqualTo("Diana"));
            Assert.That(results[3].Name, Is.EqualTo("Diana"));
        }

        [Test]
        public void ReadWorkEntries_WithInvalidDateFormat_ThrowsInvalidDataException()
        {
            // Arrange
            var filePath = CreateTestFile("InvalidDate.xlsx", new[]
            {
                new[] { "Team member", "Date", "Duration", "Hours" },
                new[] { "Eve", "2024-01-15", "08:00:00", "8" }
            });

            // Act & Assert
            Assert.Throws<InvalidOperationException>(() => _excelReader.ReadWorkEntries(filePath));
        }

        [Test]
        public void ReadWorkEntries_WithInvalidDurationFormat_ThrowsInvalidDataException()
        {
            // Arrange
            var filePath = CreateTestFile("InvalidDuration.xlsx", new[]
            {
                new[] { "Team member", "Date", "Duration", "Hours" },
                new[] { "Frank", "01/15/2024", "8 hours", "8" }
            });

            // Act & Assert
            Assert.Throws<InvalidOperationException>(() => _excelReader.ReadWorkEntries(filePath));
        }

        [Test]
        public void ReadWorkEntries_WithInvalidHoursFormat_ThrowsInvalidDataException()
        {
            // Arrange
            var filePath = CreateTestFile("InvalidHours.xlsx", new[]
            {
                new[] { "Team member", "Date", "Duration", "Hours" },
                new[] { "Grace", "01/15/2024", "08:00:00", "eight" }
            });

            // Act & Assert
            Assert.Throws<InvalidOperationException>(() => _excelReader.ReadWorkEntries(filePath));
        }

        [Test]
        public void ReadWorkEntries_WithMissingRequiredFields_SkipsRow()
        {
            // Arrange
            var filePath = CreateTestFile("MissingFields.xlsx", new[]
            {
                new[] { "Team member", "Date", "Duration", "Hours" },
                new[] { "Henry", "01/15/2024", "", "8" },
                new[] { "Henry", "01/16/2024", "09:00:00", "9" }
            });

            // Act
            var results = _excelReader.ReadWorkEntries(filePath);

            // Assert
            Assert.That(results.Count, Is.EqualTo(1));
            Assert.That(results[0].StartTime.Date, Is.EqualTo(new DateTime(2024, 1, 16)));
        }

        [Test]
        public void ReadWorkEntries_WithNullFilePath_ThrowsArgumentException()
        {
            // Act & Assert
            Assert.Throws<ArgumentException>(() => _excelReader.ReadWorkEntries(null));
        }

        [Test]
        public void ReadWorkEntries_WithWhitespaceFilePath_ThrowsArgumentException()
        {
            // Act & Assert
            Assert.Throws<ArgumentException>(() => _excelReader.ReadWorkEntries("   "));
        }

        private string CreateTestFile(string fileName, string[][] rowData)
        {
            var filePath = Path.Combine(_testFilesDirectory, fileName);
            
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");
                
                for (int row = 0; row < rowData.Length; row++)
                {
                    for (int col = 0; col < rowData[row].Length; col++)
                    {
                        worksheet.Cell(row + 1, col + 1).Value = rowData[row][col];
                    }
                }
                
                workbook.SaveAs(filePath);
            }
            
            return filePath;
        }

        [TearDown]
        public void TearDown()
        {
            try
            {
                if (Directory.Exists(_testFilesDirectory))
                {
                    foreach (var file in Directory.GetFiles(_testFilesDirectory))
                    {
                        File.Delete(file);
                    }
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }
}