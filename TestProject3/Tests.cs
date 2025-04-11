using BK_Details_App.ViewModels;
using ClosedXML.Excel;

namespace TestProject3
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
            string s = AppContext.BaseDirectory;
            testFilePath = Path.Combine(AppContext.BaseDirectory.Substring(0,
               AppContext.BaseDirectory.IndexOf("TestProject3") - 1),
               "BK_Details_App\\bin\\Debug\\net8.0\\Materials\\test.xlsx");
        }

        private readonly string logFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test_results.txt");
        private string testFilePath;

        void TestAddToFavourite(string _material)
        {
            var service = new DetailsVM(skipInit: true);
            var Favs = service.ReadFavorites(testFilePath);
            if (Favs.Any(x => x == _material))
            {
                return;
            }
            XLWorkbook workbook;
            if (File.Exists(testFilePath))
            {
                workbook = new XLWorkbook(testFilePath);
            }
            else
            {
                workbook = new XLWorkbook();
            }

            string sheetName = "Избранное";
            var sheet = workbook.Worksheets.Contains(sheetName)
            ? workbook.Worksheet(sheetName)
            : workbook.AddWorksheet(sheetName);

            //Определяем первую пустую строку
            int lastRow = sheet.LastRowUsed()?.RowNumber() + 1 ?? 1;
            sheet.Cell(lastRow, 1).Value = _material;

            //Сохраняем файл
            workbook.SaveAs(testFilePath);
        }

        [Test, Order(0)]
        public void Test_FileDoesExist_ReturnsListWithValue()
        {
            using (var workbook = new XLWorkbook(testFilePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    worksheet.Clear();
                }
                workbook.Save();
            }
            string testName = "Test_FileDoesExist_ReturnsListWithValue";
            string expected = "[1]";
            try
            {
                var service = new DetailsVM(skipInit: true);
                TestAddToFavourite("1");
                var result = service.ReadFavorites(testFilePath);

                result = service.ReadFavorites(testFilePath);

                string actual = "[" + string.Join(", ", result) + "]";
                bool passed = result.Count == 1;

                string status = passed ? "Passed" : "Failed";

                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Datetime now: {System.DateTime.Now}{Environment.NewLine}" +
                             $"Expected: {expected}{Environment.NewLine}" +
                             $"Actual: {actual}{Environment.NewLine}" +
                             $"Status: {status}{Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";

                File.AppendAllText(logFile, log);

                Assert.IsTrue(passed);
            }
            catch (Exception ex)
            {
                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Exception: {ex.Message}{Environment.NewLine}" +
                             $"Status: Failed (Exception){Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";
                File.AppendAllText(logFile, log);
                Assert.Fail("Exception thrown: " + ex.ToString());
            }
        }


        [Test, Order(1)]
        public void Test_FileDoesExist_ReturnsListWithValues()
        {
            using (var workbook = new XLWorkbook(testFilePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    worksheet.Clear();
                }
                workbook.Save();
            }
            string testName = "Test_FileDoesExist_ReturnsListWithValues";
            string expected = "[1, 2, 3, 4, 5]";
            try
            {
                var service = new DetailsVM(skipInit: true);
                TestAddToFavourite("1");
                TestAddToFavourite("2");
                TestAddToFavourite("3");
                TestAddToFavourite("4");
                TestAddToFavourite("5");
                var result = service.ReadFavorites(testFilePath);

                string actual = "[" + string.Join(", ", result) + "]";
                bool passed = result.Count == 5;

                string status = passed ? "Passed" : "Failed";

                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Datetime now: {System.DateTime.Now}{Environment.NewLine}" +
                             $"Expected: {expected}{Environment.NewLine}" +
                             $"Actual: {actual}{Environment.NewLine}" +
                             $"Status: {status}{Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";

                File.AppendAllText(logFile, log);

                Assert.IsTrue(passed);
            }
            catch (Exception ex)
            {
                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Exception: {ex.Message}{Environment.NewLine}" +
                             $"Status: Failed (Exception){Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";
                File.AppendAllText(logFile, log);
                Assert.Fail("Exception thrown: " + ex.ToString());
            }
        }

        [Test, Order(2)]
        public void Test_FileExist_ReturnsEmptyList()
        {
            using (var workbook = new XLWorkbook(testFilePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    worksheet.Clear();
                }
                workbook.Save();
            }
            string testName = "Test_FileExist_ReturnsEmptyList";
            string expected = "[]";
            try
            {
                var service = new DetailsVM(skipInit: true);
                var result = service.ReadFavorites(testFilePath);

                result = service.ReadFavorites(testFilePath);

                string actual = "[" + string.Join(", ", result) + "]";
                bool passed = result.Count == 0;

                string status = passed ? "Passed" : "Failed";

                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Datetime now: {System.DateTime.Now}{Environment.NewLine}" +
                             $"Expected: {expected}{Environment.NewLine}" +
                             $"Actual: {actual}{Environment.NewLine}" +
                             $"Status: {status}{Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";

                File.AppendAllText(logFile, log);

                Assert.IsTrue(passed);
            }
            catch (Exception ex)
            {
                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Exception: {ex.Message}{Environment.NewLine}" +
                             $"Status: Failed (Exception){Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";
                File.AppendAllText(logFile, log);
                Assert.Fail("Exception thrown: " + ex.ToString());
            }
        }

        [Test, Order(3)]
        public void Test_FileWithDuplicateFavorites_ReturnsUniqueList()
        {
            using (var workbook = new XLWorkbook(testFilePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    worksheet.Clear();
                }
                workbook.Save();
            }

            string testName = "Test_FileWithDuplicateFavorites_ReturnsUniqueList";
            string expected = "[1, 2]";
            try
            {
                var service = new DetailsVM(skipInit: true);
                TestAddToFavourite("1");
                TestAddToFavourite("1");
                TestAddToFavourite("2");
                var result = service.ReadFavorites(testFilePath);

                var actual = "[" + string.Join(", ", result.Distinct()) + "]";
                bool passed = result.Count == 2;

                string status = passed ? "Passed" : "Failed";

                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Datetime now: {DateTime.Now}{Environment.NewLine}" +
                             $"Expected: {expected}{Environment.NewLine}" +
                             $"Actual: {actual}{Environment.NewLine}" +
                             $"Status: {status}{Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";

                File.AppendAllText(logFile, log);
                Assert.IsTrue(passed);
            }
            catch (Exception ex)
            {
                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Exception: {ex.Message}{Environment.NewLine}" +
                             $"Status: Failed (Exception){Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";
                File.AppendAllText(logFile, log);
                Assert.Fail("Exception thrown: " + ex.ToString());
            }
        }

        [Test, Order(4)]
        public void Test_FileWithExtraColumns_IgnoresOtherColumns()
        {
            using (var workbook = new XLWorkbook(testFilePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    worksheet.Clear();
                }
                workbook.Save();
            }

            using (var workbook = new XLWorkbook(testFilePath))
            {
                var worksheet = workbook.Worksheet("Избранное");

                worksheet.Cell(1, 1).Value = "1";
                worksheet.Cell(1, 2).Value = "IgnoreMe";
                worksheet.Cell(2, 1).Value = "2";
                worksheet.Cell(2, 2).Value = "AlsoIgnore";

                workbook.SaveAs(testFilePath);
            }

            string testName = "Test_FileWithExtraColumns_IgnoresOtherColumns";
            string expected = "[1, 2]";
            try
            {
                var service = new DetailsVM(skipInit: true);
                var result = service.ReadFavorites(testFilePath);

                string actual = "[" + string.Join(", ", result) + "]";
                bool passed = result.SequenceEqual(new List<string> { "1", "2" });

                string status = passed ? "Passed" : "Failed";

                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Datetime now: {DateTime.Now}{Environment.NewLine}" +
                             $"Expected: {expected}{Environment.NewLine}" +
                             $"Actual: {actual}{Environment.NewLine}" +
                             $"Status: {status}{Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";
                File.AppendAllText(logFile, log);
                Assert.IsTrue(passed);
            }
            catch (Exception ex)
            {
                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Exception: {ex.Message}{Environment.NewLine}" +
                             $"Status: Failed (Exception){Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";
                File.AppendAllText(logFile, log);
                Assert.Fail("Exception thrown: " + ex.ToString());
            }
        }
    }
}