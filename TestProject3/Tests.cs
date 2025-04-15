using BK_Details_App;
using BK_Details_App.Models;
using BK_Details_App.ViewModels;
using ClosedXML.Excel;
using CsvHelper;
using ExcelDataReader.Log;
using Microsoft.VisualStudio.TestPlatform.TestHost;
using Microsoft.Extensions.Logging;
using Aspose.Cells;
using System.Reflection;

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
        public List<Materials> _materialsList = new List<Materials>();

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

        public List<Materials> TestGetMaterials()
        {
            try
            {
                List<string> buf = [.. new DetailsVM(skipInit: true).ReadFavorites(testFilePath)];
                TestReadFromExcelFile();
                if (buf.Count > 0)
                {
                    List<Materials> FavsList = _materialsList.Where(x => buf.Contains(x.Name)).ToList();
                    return FavsList;
                }
                else
                {
                    return new List<Materials>();
                }
            }
            catch (Exception ex)
            {
                using ILoggerFactory factory = LoggerFactory.Create(builder => builder.AddConsole());
                ILogger logger = factory.CreateLogger<Program>();
                logger.LogInformation($":::::EXCEPTION:::::::::::::::EXCEPTION:::::::::::::::EXCEPTION::::::::{ex.ToString()}.", "what");
                return new List<Materials>();
            }
        }

        public void TestReadFromExcelFile()
        {
            try
            {
                string filePath = Path.Combine(AppContext.BaseDirectory.Substring(0,
                AppContext.BaseDirectory.IndexOf("TestProject3") - 1),
               "BK_Details_App\\Materials\\materials.xlsx");
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(filePath);
                WorksheetCollection collection = wb.Worksheets;
                Random random = new Random();

                for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
                {
                    Aspose.Cells.Worksheet worksheet = collection[worksheetIndex];

                    bool lastWasBold = false;
                    int rows = worksheet.Cells.MaxDataRow;
                    int cols = worksheet.Cells.MaxDataColumn;

                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= cols; j++)
                        {
                            var cell = worksheet.Cells[i, j];
                            var style = cell.GetStyle();
                            var value = cell.StringValue;

                            if (string.IsNullOrWhiteSpace(value))
                                break;

                            var material = new Materials()
                            {
                                IdNumber = worksheet.Cells[i, j - 1].StringValue == "" ? random.Next(1, 1000) : worksheet.Cells[i, j - 1].IntValue,
                                Name = cell.StringValue,
                                Measurement = worksheet.Cells[i, j + 1].StringValue,
                                Analogs = string.IsNullOrWhiteSpace(worksheet.Cells[i, j + 2].StringValue) ? "Аналогов нет" : worksheet.Cells[i, j + 2].StringValue,
                                Note = string.IsNullOrWhiteSpace(worksheet.Cells[i, j + 3].StringValue) ? "Примечание отсутствует" : worksheet.Cells[i, j + 3].StringValue,
                                GroupNavigation = new Groups(),
                                Group = 0,
                                CategoryNavigation = new Category(),
                                Category = 0
                            };
                            _materialsList.Add(material);
                            lastWasBold = false;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //ShowError("ReadFromExcelFile: Ошибка!", ex.ToString());
                using ILoggerFactory factory = LoggerFactory.Create(builder => builder.AddConsole());
                Microsoft.Extensions.Logging.ILogger logger = factory.CreateLogger<Program>();
                logger.LogInformation($":::::EXCEPTION:::::::::::::::EXCEPTION:::::::::::::::EXCEPTION::::::::{ex.ToString()}.", "what");
            }
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

        [Test, Order(5)]
        public void Test_OneFavoriteMaterial_ReturnsOneMaterial()
        {
            var expected = "[Двутавр 10Б1-ГК ГОСТ Р 57837-2017 / Ст3сп ГОСТ 535-88]";
            var testName = "Test_OneFavoriteMaterial_ReturnsOneMaterial";

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

                worksheet.Cell(1, 1).Value = "Двутавр 10Б1-ГК ГОСТ Р 57837-2017 / Ст3сп ГОСТ 535-88";

                workbook.SaveAs(testFilePath);
            }

            var result = TestGetMaterials();
            var actual = "[" + string.Join(", ", result.Select(x => x.Name)) + "]";
            bool passed = result.Count == 1 && result[0].Name == "Двутавр 10Б1-ГК ГОСТ Р 57837-2017 / Ст3сп ГОСТ 535-88";

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

        [Test, Order(6)]
        public void Test_EmptyFavorites_ReturnsEmptyList()
        {
            var testName = "Test_EmptyFavorites_ReturnsEmptyList";
            var expected = "[]";

            using (var workbook = new XLWorkbook(testFilePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    worksheet.Clear();
                }
                workbook.Save();
            }

            var result = TestGetMaterials();
            var actual = result == null ? "Лист не создан" : "[" + string.Join(", ", result.Select(x => x.Name)) + "]";
            bool passed =  result!= null && result.Count == 0;

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

        [Test, Order(7)]
        public void Test_AllMaterialsAreFavorite_ReturnsFewMaterials()
        {
            var testName = "Test_AllMaterialsAreFavorite_ReturnsFewMaterials";
            var expected = "[Винт 1-3,5х25 Н Хим.Фос.прп. ТУ 16 40-015-55798700-2006, Винт 2-3,5х16 Н Хим.Фос.прп. ТУ 16 40-015-55798700-2006]";

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

                worksheet.Cell(1, 1).Value = "Винт 1-3,5х25 Н Хим.Фос.прп. ТУ 16 40-015-55798700-2006";
                worksheet.Cell(2, 1).Value = "Винт 2-3,5х16 Н Хим.Фос.прп. ТУ 16 40-015-55798700-2006";

                workbook.SaveAs(testFilePath);
            }
            _materialsList.Clear();
            var result = TestGetMaterials();
            var actual = "[" + string.Join(", ", result.Select(x => x.Name)) + "]";
            bool passed = result.Count == 2 && result[0].Name == "Винт 1-3,5х25 Н Хим.Фос.прп. ТУ 16 40-015-55798700-2006"
                && result[1].Name == "Винт 2-3,5х16 Н Хим.Фос.прп. ТУ 16 40-015-55798700-2006";

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

        [Test, Order(8)]
        public void Test_WeAreExpectingTheListDataType_ReturnsListDataType()
        {
            var testName = "Test_WeAreExpectingTheListDataType_ReturnsListDataType";
            var expected = typeof(List<Materials>);

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

                worksheet.Cell(1, 1).Value = "Винт 1-3,5х25 Н Хим.Фос.прп. ТУ 16 40-015-55798700-2006";

                workbook.SaveAs(testFilePath);
            }

            var result = TestGetMaterials();
            var actual = "[" + string.Join(", ", result.Select(x => x.Name)) + "]";
            bool passed = result.GetType() == expected;

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

        [Test, Order(9)]
        public void Test_WeAreNotExpectingTheListStringDataType_ReturnsListStringDataType()
        {
            var testName = "Test_WeAreNotExpectingTheListStringDataType_ReturnsListStringDataType";
            var expected = typeof(string);

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

                worksheet.Cell(1, 1).Value = "Винт 1-3,5х25 Н Хим.Фос.прп. ТУ 16 40-015-55798700-2006";

                workbook.SaveAs(testFilePath);
            }

            var result = TestGetMaterials();
            var actual = "[" + string.Join(", ", result.Select(x => x.Name)) + "]";
            bool passed = result.GetType() != expected;

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
    }
}