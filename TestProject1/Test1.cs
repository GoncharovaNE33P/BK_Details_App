using BK_Details_App.Models;
using BK_Details_App.ViewModels;
using ClosedXML.Excel;
using Microsoft.CodeCoverage.Core.Reports.Coverage;

namespace TestProject1
{
    [TestClass]
    public sealed class Test1
    {
        private readonly string logFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test_results.txt");
        private string testFilePath;

        void TestAddToFavourite(string _material)
        {
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

        [TestInitialize]
        public void Setup()
        {
            testFilePath = Path.Combine(AppContext.BaseDirectory.Substring(0,
               AppContext.BaseDirectory.IndexOf("TestProject1") - 1),
               "BK_Details_App\\bin\\Debug\\net8.0\\Materials\\test.xlsx");
        }

        //[TestMethod]
        //public void Test_FileDoesNotExist_ReturnsEmptyList()
        //{            
        //    string testName = "Test_FileDoesNotExist_ReturnsEmptyList";
        //    string expected = "[]";
        //    try
        //    {
        //        // Удаляем файл, если он есть, чтобы проверить поведение при его отсутствии
        //        if (File.Exists(testFilePath))
        //            File.Delete(testFilePath);
        //        var service = new DetailsVM(skipInit: true); // замени на реальный класс, где находится ReadFavorites
        //        var result = service.ReadFavorites(testFilePath);

        //        string actual = "[" + string.Join(", ", result) + "]";
        //        bool passed = result.Count == 0;

        //        string status = passed ? "Passed" : "Failed";

        //        string log = $"Test Name: {testName}{Environment.NewLine}" + 
        //                     $"Datetime now: {System.DateTime.Now}{Environment.NewLine}" +
        //                     $"Expected: {expected}{Environment.NewLine}" +
        //                     $"Actual: {actual}{Environment.NewLine}" +
        //                     $"Status: {status}{Environment.NewLine}" +
        //                     $"----------------------{Environment.NewLine}";

        //        File.AppendAllText(logFile, log);

        //        Assert.IsTrue(passed);
        //    }
        //    catch (Exception ex)
        //    {
        //        string log = $"Test Name: {testName}{Environment.NewLine}" +
        //                     $"Exception: {ex.Message}{Environment.NewLine}" +
        //                     $"Status: Failed (Exception){Environment.NewLine}" +
        //                     $"----------------------{Environment.NewLine}";
        //        File.AppendAllText(logFile, log);
        //        Assert.Fail("Exception thrown: " + ex.ToString());
        //    }
        //}

        [TestMethod]
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

        //[TestMethod]
        //public void Test_ReturnValueIsNotNull()
        //{
        //    if (!File.Exists(testFilePath))
        //    {
        //        File.Create(testFilePath);
        //    }
        //    string testName = "Test_ReturnValueIsNotNull";
        //    string expected = "[1]";
        //    try
        //    {
        //        var service = new DetailsVM(skipInit: true);
        //        var result = service.ReadFavorites(testFilePath);

        //        string actual = "[" + string.Join(", ", result) + "]";

        //        string status = result is null ? "Passed" : "Failed";

        //        string log = $"Test Name: {testName}{Environment.NewLine}" +
        //                     $"Datetime now: {System.DateTime.Now}{Environment.NewLine}" +
        //                     $"Expected: {expected}{Environment.NewLine}" +
        //                     $"Actual: {actual}{Environment.NewLine}" +
        //                     $"Status: {status}{Environment.NewLine}" +
        //                     $"----------------------{Environment.NewLine}";

        //        File.AppendAllText(logFile, log);

        //        Assert.IsNotNull(result);
        //    }
        //    catch (Exception ex)
        //    {
        //        string log = $"Test Name: {testName}{Environment.NewLine}" +
        //                     $"Exception: {ex.Message}{Environment.NewLine}" +
        //                     $"Status: Failed (Exception){Environment.NewLine}" +
        //                     $"----------------------{Environment.NewLine}";
        //        File.AppendAllText(logFile, log);
        //        Assert.Fail("Exception thrown: " + ex.ToString());
        //    }
        //}
        /*
                [TestMethod]
                public void ReadFavorites_FileInUse_ReturnsEmptyListWithoutCrash()
                {
                    var wb = new XLWorkbook();
                    var sheet = wb.AddWorksheet("Избранное");
                    sheet.Cell(1, 1).Value = "Locked";
                    wb.SaveAs(testFilePath);
                    var service = new DetailsVM();

                    using (var stream = new FileStream(testFilePath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        var result = service.ReadFavorites(testFilePath);
                        Assert.IsNotNull(result);
                        Assert.AreEqual(0, result.Count); // метод должен поймать исключение
                    }
                }*/

        [TestMethod]
        public void Test_FileDoesExist_ReturnsListWithValues()
        {
            using(var workbook = new XLWorkbook(testFilePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    worksheet.Clear();
                }
                workbook.Save();
            }
            string testName = "Test_FileDoesExist_ReturnsListWithValues";
            string expected = "[1, 1, 2, 3, 4, 5]";
            try
            {
                var service = new DetailsVM(skipInit: true);
                TestAddToFavourite("1");
                TestAddToFavourite("1");
                TestAddToFavourite("2");
                TestAddToFavourite("3");
                TestAddToFavourite("4");
                TestAddToFavourite("5");
                var result = service.ReadFavorites(testFilePath);

                result = service.ReadFavorites(testFilePath);

                string actual = "[" + string.Join(", ", result) + "]";
                bool passed = result.Count == 6;

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
    }
}
